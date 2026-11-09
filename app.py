import io
import os
import re
import time
from datetime import datetime
import pandas as pd
from bs4 import BeautifulSoup

import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service


# ==========================================
# 웹페이지 기본 설정
# ==========================================
st.set_page_config(
    page_title="CSI 현장점검 결과 자동 수집기",
    page_icon="🏗️",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1400px; }
    .main-title { font-size: 24px; font-weight: 700; color: #0f172a; margin-bottom: 4px; }
    .sub-title { font-size: 13px; color: #475569; margin-bottom: 16px; border-bottom: 1px solid #e2e8f0; padding-bottom: 10px; }
    .step-card { background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px; margin-bottom: 14px; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================
# 1. 크롬 브라우저 생성 (화면 표출 모드)
# ==========================================
def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--start-maximized")

    # 브라우저 이미지 로딩 유지 (사람이 화면 볼 때 자연스럽게 렌더링)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    chromium_paths = ["/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome"]
    driver_paths = ["/usr/bin/chromedriver", "/usr/lib/chromium-browser/chromedriver"]

    binary_loc = next((p for p in chromium_paths if os.path.exists(p)), None)
    driver_loc = next((p for p in driver_paths if os.path.exists(p)), None)

    if binary_loc:
        options.binary_location = binary_loc

    if driver_loc:
        service = Service(executable_path=driver_loc)
        driver = webdriver.Chrome(service=service, options=options)
    else:
        driver = webdriver.Chrome(options=options)

    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def dismiss_alerts(driver):
    try:
        alert = driver.switch_to.alert
        txt = alert.text
        alert.accept()
        return txt
    except Exception:
        return None


# ==========================================
# 2. CSI 자동 로그인 (선택 사항)
# ==========================================
def auto_login_csi(driver, user_id, user_pw):
    driver.get("https://www.csi.go.kr/main.do")
    dismiss_alerts(driver)

    js_code = """
    var uid = arguments[0];
    var upw = arguments[1];
    var idEl = document.querySelector('input[name="user_id"], input[name="userId"], input#user_id, input#userId, input[placeholder*="아이디"]');
    var pwEl = document.querySelector('input[name="pswd"], input[name="password"], input[name="user_pw"], input[type="password"]');

    if (!idEl || !pwEl) {
        var loginBtn = document.querySelector('a.login-btn, a[href*="login"], button.login-btn');
        if (loginBtn) { loginBtn.click(); }
        idEl = document.querySelector('input[name="user_id"], input[name="userId"], input#user_id, input#userId, input[placeholder*="아이디"]');
        pwEl = document.querySelector('input[name="pswd"], input[name="password"], input[name="user_pw"], input[type="password"]');
    }

    if (idEl && pwEl) {
        idEl.value = uid;
        pwEl.value = upw;
        idEl.dispatchEvent(new Event('change', { bubbles: true }));
        pwEl.dispatchEvent(new Event('change', { bubbles: true }));
        if (typeof go_login === 'function') { go_login(); return true; }
        if (document.forms['loginForm']) { document.forms['loginForm'].submit(); return true; }
        return 'filled';
    }
    return false;
    """
    try:
        driver.execute_script(js_code, user_id, user_pw)
    except Exception:
        pass

    # 비밀번호 엔터 전송
    try:
        pw_el = driver.find_elements(By.CSS_SELECTOR, "input[type='password']")
        if pw_el:
            pw_el[0].send_keys(Keys.RETURN)
    except Exception:
        pass


# ==========================================
# 3. 전체 페이지 수 감지
# ==========================================
def get_total_page_info(driver):
    try:
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        help_inline = soup.find("div", class_="help-inline")
        if help_inline:
            text = help_inline.get_text(strip=True)
            match = re.search(r"/\s*(\d+)\s*페이지\s*\)\s*/\s*총\s*(\d+)", text)
            if match:
                return int(match.group(1)), int(match.group(2))
    except Exception:
        pass
    return 1, 0


# ==========================================
# 4. 텍스트 정리 및 테이블 파싱
# ==========================================
def clean_text(text):
    if not text:
        return ""
    text = re.sub(r"[ \t\r\f\v]+", " ", str(text))
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return " ".join(lines).strip()


def extract_attachment_info(td_elem):
    if not td_elem:
        return ""
    files = []
    for a in td_elem.find_all("a"):
        t = clean_text(a.get_text())
        if t:
            files.append(t)
    for img in td_elem.find_all("img"):
        src = img.get("src", "")
        alt = clean_text(img.get("alt", ""))
        if src and "noimg" not in src and "ico_btn" not in src:
            files.append(alt if alt else src)
    if files:
        return ", ".join(files)
    return clean_text(td_elem.get_text())


def parse_key_value_table(table):
    info = {}
    if not table:
        return info

    for tr in table.find_all("tr"):
        cells = tr.find_all(["th", "td"], recursive=False)
        if not cells:
            continue
        idx = 0
        while idx < len(cells):
            cell = cells[idx]
            classes = cell.get("class") or []
            is_head = (cell.name == "th") or ("td-head" in classes)
            if is_head and idx + 1 < len(cells):
                next_cell = cells[idx + 1]
                next_classes = next_cell.get("class") or []
                if (next_cell.name == "th") or ("td-head" in next_classes):
                    idx += 1
                    continue

                key = clean_text(cell.get_text())
                if key in ("사진", "첨부파일", "현장점검 확인서") or next_cell.find("a") or next_cell.find("img"):
                    val = extract_attachment_info(next_cell)
                else:
                    val = clean_text(next_cell.get_text())

                if key:
                    info[key] = val
                idx += 2
                continue
            idx += 1

    return info


# ==========================================
# 5. 상세 페이지 HTML 파싱 (현장점검 결과 탭)
# ==========================================
def parse_csi_detail_html(html):
    soup = BeautifulSoup(html, "html.parser")
    tab = soup.find("div", id="tab") or soup

    parsed = {
        "meta": {},
        "check_confirm_title": "",
        "overview": {},
        "progress": [],
        "permit": [],
        "constructors": [],
        "supervisors": [],
        "engineers_con": [],
        "engineers_spv": [],
        "confirm_file": "",
        "memos": [],
        "inspections": [],
    }

    for form_id in ("viewForm", "main"):
        form = soup.find("form", id=form_id)
        if form:
            for inp in form.find_all("input", {"type": "hidden"}):
                name = inp.get("name")
                val = inp.get("value", "").strip()
                if name and not parsed["meta"].get(name):
                    parsed["meta"][name] = val

    for h1 in tab.find_all("h1"):
        if h1.find_parent("div", id="accordion"):
            continue

        title_text = clean_text(h1.get_text())
        next_div = h1.find_next_sibling("div")
        table = next_div.find("table") if next_div else h1.find_next("table")

        if title_text == "현장점검 확인서":
            kv = parse_key_value_table(table)
            if "점검명" in kv:
                parsed["check_confirm_title"] = kv.get("점검명", "")
            elif "현장점검 확인서" in kv:
                parsed["confirm_file"] = kv.get("현장점검 확인서", "")

        elif "공사 및 시공" in title_text:
            parsed["overview"] = parse_key_value_table(table)

        elif title_text == "공사현황" and table:
            current_sector = ""
            for tbody in table.find_all("tbody"):
                for tr in tbody.find_all("tr"):
                    tds = tr.find_all(["td", "th"], recursive=False)
                    if not tds:
                        continue
                    if len(tds) == 6:
                        current_sector = clean_text(tds[0].get_text())
                        sub_type = clean_text(tds[1].get_text())
                        vals = [clean_text(c.get_text()) for c in tds[2:]]
                    elif len(tds) == 5:
                        sub_type = clean_text(tds[0].get_text())
                        vals = [clean_text(c.get_text()) for c in tds[1:]]
                    else:
                        continue
                    parsed["progress"].append(
                        {
                            "부문": current_sector,
                            "구분": sub_type,
                            "계약일": vals[0] if len(vals) > 0 else "",
                            "도급액": vals[1] if len(vals) > 1 else "",
                            "공사기간": vals[2] if len(vals) > 2 else "",
                            "현재공정(%)": vals[3] if len(vals) > 3 else "",
                        }
                    )

        elif "사업 인허가" in title_text and table:
            tbody = table.find("tbody") or table
            for tr in tbody.find_all("tr"):
                tds = tr.find_all("td", recursive=False)
                if len(tds) >= 4:
                    vals = [clean_text(c.get_text()) for c in tds[:4]]
                    if any(vals):
                        parsed["permit"].append(
                            {
                                "승인일자": vals[0],
                                "승인기관": vals[1],
                                "담당부서": vals[2],
                                "전화번호": vals[3],
                            }
                        )

        elif title_text == "시공업체" and table:
            tbodies = table.find_all("tbody") or [table]
            for tb in tbodies:
                kv = parse_key_value_table(tb)
                if kv:
                    parsed["constructors"].append(kv)

        elif "건설사업관리용역사업자" in title_text and table:
            tbodies = table.find_all("tbody") or [table]
            for tb in tbodies:
                kv = parse_key_value_table(tb)
                if kv:
                    parsed["supervisors"].append(kv)

        elif title_text == "건설기술인" and table:
            tbody = table.find("tbody") or table
            for tr in tbody.find_all("tr"):
                tds = tr.find_all("td", recursive=False)
                if len(tds) >= 6:
                    vals = [clean_text(c.get_text()) for c in tds[:6]]
                    if any(vals):
                        parsed["engineers_con"].append(
                            {
                                "기술인 구분": vals[0],
                                "성명": vals[1],
                                "소속회사": vals[2],
                                "자격종류": vals[3],
                                "자격증번호": vals[4],
                                "휴대전화번호": vals[5],
                            }
                        )

        elif "감독·건설사업관리기술자" in title_text and table:
            tbody = table.find("tbody") or table
            for tr in tbody.find_all("tr"):
                tds = tr.find_all("td", recursive=False)
                if len(tds) >= 6:
                    vals = [clean_text(c.get_text()) for c in tds[:6]]
                    if any(vals):
                        parsed["engineers_spv"].append(
                            {
                                "기술인 구분": vals[0],
                                "성명": vals[1],
                                "소속회사": vals[2],
                                "자격종류": vals[3],
                                "자격증번호": vals[4],
                                "휴대전화번호": vals[5],
                            }
                        )

        elif "현장점검 메모" in title_text and table:
            tbody = table.find("tbody", id="memo")
            if tbody:
                for tr in tbody.find_all("tr"):
                    tds = tr.find_all("td", recursive=False)
                    if len(tds) >= 4:
                        inp = tds[0].find("input")
                        memo_txt = inp.get("value", "").strip() if inp else clean_text(tds[0].get_text())
                        vals = [
                            memo_txt,
                            clean_text(tds[1].get_text()),
                            clean_text(tds[2].get_text()),
                            clean_text(tds[3].get_text()),
                        ]
                        if any(vals):
                            parsed["memos"].append(
                                {
                                    "내용": vals[0],
                                    "점검일": vals[1],
                                    "작성자": vals[2],
                                    "작성일": vals[3],
                                }
                            )

    accordion = tab.find("div", id="accordion")
    if accordion:
        panels = accordion.find_all("div", class_="panel", recursive=False)
        if not panels:
            panels = accordion.find_all("div", class_="panel-default")

        for p_idx, panel in enumerate(panels, start=1):
            heading_el = panel.find("h4", class_="panel-title")
            panel_title = clean_text(heading_el.get_text()).lstrip("·").strip() if heading_el else ""
            matter_inp = panel.find("input", {"name": "impropt_matter_no"})
            matter_no = matter_inp.get("value", "").strip() if matter_inp else ""

            def _extract_sub_div(div_class):
                sub_div = panel.find("div", class_=div_class)
                if sub_div and sub_div.find("table"):
                    return parse_key_value_table(sub_div.find("table"))
                return {}

            parsed["inspections"].append(
                {
                    "순번": p_idx,
                    "지적사항번호": matter_no,
                    "패널제목": panel_title,
                    "지적사항": _extract_sub_div("cnfmn"),
                    "이의제기": _extract_sub_div("fobjc"),
                    "이의제기결과": _extract_sub_div("fobjcRslt"),
                    "조치결과": _extract_sub_div("managtDtls"),
                    "조치결과확인": _extract_sub_div("managtRslt"),
                    "행정조치": _extract_sub_div("administ"),
                    "조치미이행": _extract_sub_div("failure"),
                }
            )

    return parsed


# ==========================================
# 6. 상세 페이지 '현장점검 결과' 탭 활성화 및 파싱
# ==========================================
def parse_detail_page(driver):
    time.sleep(1.0)

    # 탭 클릭 또는 해시 변경으로 #cnfrmnView 탭 활성화
    try:
        driver.execute_script(
            "var link = document.querySelector('a[href*=\"#cnfrmnView\"]'); "
            "if(link) { link.click(); } "
            "window.location.hash = '#cnfrmnView'; "
            "if(typeof loadTab === 'function') { loadTab('#cnfrmnView'); } "
            "else if(window.onhashchange) { window.onhashchange(); }"
        )
    except Exception:
        pass

    # 내용 렌더링 대기
    start_w = time.time()
    while time.time() - start_w < 6.0:
        html = driver.page_source
        if "accordion" in html or "시공업체" in html or "현장점검 확인서" in html:
            break
        time.sleep(0.4)

    return parse_csi_detail_html(driver.page_source)


# ==========================================
# 7. 엑셀 스키마 빌더 및 바이트 생성
# ==========================================
def build_column_schema_and_rows(collected_sites):
    max_con_cnt = max((len(s["parsed"].get("constructors", [])) for s in collected_sites), default=1) or 1
    max_spv_cnt = max((len(s["parsed"].get("supervisors", [])) for s in collected_sites), default=1) or 1
    max_eng_con_cnt = max((len(s["parsed"].get("engineers_con", [])) for s in collected_sites), default=1) or 1
    max_eng_spv_cnt = max((len(s["parsed"].get("engineers_spv", [])) for s in collected_sites), default=1) or 1
    max_memo_cnt = max((len(s["parsed"].get("memos", [])) for s in collected_sites), default=1) or 1
    max_insp_cnt = max((len(s["parsed"].get("inspections", [])) for s in collected_sites), default=0)

    has_list_info = any(bool(s.get("list_info")) for s in collected_sites)
    col_defs = []

    def add_col(group_name, sub_name, key):
        col_defs.append((group_name, sub_name, key))

    if has_list_info:
        add_col("목록 기본정보", "No", "목록_No")
        add_col("목록 기본정보", "목록 공사기간", "목록_공사기간")
        add_col("목록 기본정보", "현장점검예정일", "목록_현장점검예정일")
        add_col("목록 기본정보", "현장점검종류", "목록_현장점검종류")
        add_col("목록 기본정보", "지적사항 조치완료", "목록_지적사항_조치완료")

    add_col("현장점검 확인서", "점검명", "현장점검확인서_점검명")

    grp_ov = "공사 및 사업자 현황"
    add_col(grp_ov, "공사명", "현황_공사명")
    add_col(grp_ov, "현장주소", "현황_현장주소")
    add_col(grp_ov, "발주자", "현황_발주자")
    add_col(grp_ov, "전화번호", "현황_발주자_전화번호")
    add_col(grp_ov, "사업자등록번호", "현황_사업자등록번호")
    add_col(grp_ov, "법인번호", "현황_법인번호")
    add_col(grp_ov, "공공/민간 구분", "현황_공공민간구분")
    add_col(grp_ov, "공동도급", "현황_공동도급")

    grp_prog_con = "공사현황 (시공부문)"
    add_col(grp_prog_con, "구분", "공사현황_시공_구분")
    add_col(grp_prog_con, "계약일", "공사현황_시공_계약일")
    add_col(grp_prog_con, "도급액", "공사현황_시공_도급액")
    add_col(grp_prog_con, "공사기간", "공사현황_시공_공사기간")
    add_col(grp_prog_con, "현재공정(%)", "공사현황_시공_현재공정")

    grp_prog_spv = "공사현황 (건설사업관리부문)"
    add_col(grp_prog_spv, "구분", "공사현황_건설사업관리_구분")
    add_col(grp_prog_spv, "계약일", "공사현황_건설사업관리_계약일")
    add_col(grp_prog_spv, "도급액", "공사현황_건설사업관리_도급액")
    add_col(grp_prog_spv, "공사기간", "공사현황_건설사업관리_공사기간")
    add_col(grp_prog_spv, "현재공정(%)", "공사현황_건설사업관리_현재공정")

    grp_perm = "사업 인허가 관련"
    add_col(grp_perm, "승인일자", "사업인허가_승인일자")
    add_col(grp_perm, "승인기관", "사업인허가_승인기관")
    add_col(grp_perm, "담당부서", "사업인허가_담당부서")
    add_col(grp_perm, "전화번호", "사업인허가_전화번호")

    for idx in range(1, max_con_cnt + 1):
        sfx = f" {idx}" if max_con_cnt > 1 else ""
        k_sfx = f"_{idx}" if max_con_cnt > 1 else ""
        grp = f"시공업체{sfx}"
        for field in ("업체명", "대표자", "업종", "면허번호", "주소", "전화번호", "팩스번호", "분담공종", "지분(%)"):
            add_col(grp, field, f"시공업체{k_sfx}_{field}")

    for idx in range(1, max_spv_cnt + 1):
        sfx = f" {idx}" if max_spv_cnt > 1 else ""
        k_sfx = f"_{idx}" if max_spv_cnt > 1 else ""
        grp = f"건설사업관리용역사업자{sfx}"
        for field in ("업체명", "대표자", "업종", "면허번호", "주소", "전화번호", "팩스번호", "분담공종", "지분(%)"):
            add_col(grp, field, f"건설사업관리{k_sfx}_{field}")

    for idx in range(1, max_eng_con_cnt + 1):
        sfx = f" {idx}" if max_eng_con_cnt > 1 else ""
        k_sfx = f"_{idx}" if max_eng_con_cnt > 1 else ""
        grp = f"건설기술인{sfx}"
        for field in ("기술인 구분", "성명", "소속회사", "자격종류", "자격증번호", "휴대전화번호"):
            add_col(grp, field, f"건설기술인{k_sfx}_{field}")

    for idx in range(1, max_eng_spv_cnt + 1):
        sfx = f" {idx}" if max_eng_spv_cnt > 1 else ""
        k_sfx = f"_{idx}" if max_eng_spv_cnt > 1 else ""
        grp = f"감독·건설사업관리기술자{sfx}"
        for field in ("기술인 구분", "성명", "소속회사", "자격종류", "자격증번호", "휴대전화번호"):
            add_col(grp, field, f"감독기술자{k_sfx}_{field}")

    add_col("현장점검 확인서(하단)", "현장점검 확인서", "하단_현장점검확인서")
    for idx in range(1, max_memo_cnt + 1):
        sfx = f" {idx}" if max_memo_cnt > 1 else ""
        k_sfx = f"_{idx}" if max_memo_cnt > 1 else ""
        grp = f"현장점검 메모{sfx}"
        for field in ("내용", "점검일", "작성자", "작성일"):
            add_col(grp, field, f"현장점검메모{k_sfx}_{field}")

    for i in range(1, max_insp_cnt + 1):
        grp_insp = f"점검결과 {i}"
        p = f"점검결과{i}"
        add_col(grp_insp, "제목(지적사항)", f"{p}_지적_제목")
        add_col(grp_insp, "조치계획", f"{p}_지적_조치계획")
        add_col(grp_insp, "지적대상", f"{p}_지적_지적대상")
        add_col(grp_insp, "점검항목", f"{p}_지적_점검항목")
        add_col(grp_insp, "지적 통보일", f"{p}_지적_통보일")
        add_col(grp_insp, "지적 작성자", f"{p}_지적_작성자")
        add_col(grp_insp, "지적 사진", f"{p}_지적_사진")
        add_col(grp_insp, "조치사항(조치결과)", f"{p}_조치_조치사항")
        add_col(grp_insp, "조치일", f"{p}_조치_조치일")
        add_col(grp_insp, "조치 통보일", f"{p}_조치_통보일")
        add_col(grp_insp, "조치 작성자", f"{p}_조치_작성자")
        add_col(grp_insp, "조치 사진(첨부파일)", f"{p}_조치_사진")
        add_col(grp_insp, "조치이행결과", f"{p}_확인_조치이행결과")
        add_col(grp_insp, "확인일", f"{p}_확인_확인일")
        add_col(grp_insp, "확인 통보일", f"{p}_확인_통보일")
        add_col(grp_insp, "확인 작성자", f"{p}_확인_작성자")
        add_col(grp_insp, "확인 사진", f"{p}_확인_사진")

    rows_data = []
    for site in collected_sites:
        list_info = site.get("list_info") or {}
        parsed = site.get("parsed") or {}
        ov = parsed.get("overview") or {}

        row = {}

        if has_list_info:
            row["목록_No"] = list_info.get("No", "")
            row["목록_공사기간"] = list_info.get("공사기간", "")
            row["목록_현장점검예정일"] = list_info.get("현장점검예정일", "")
            row["목록_현장점검종류"] = list_info.get("현장점검종류", "")
            row["목록_지적사항_조치완료"] = list_info.get("지적사항_조치완료", "")

        row["현장점검확인서_점검명"] = parsed.get("check_confirm_title") or list_info.get("점검명", "")
        row["현황_공사명"] = ov.get("공사명") or list_info.get("공사명", "")
        row["현황_현장주소"] = ov.get("현장주소") or list_info.get("현장주소", "")
        row["현황_발주자"] = ov.get("발주자") or list_info.get("발주자", "")
        row["현황_발주자_전화번호"] = ov.get("전화번호", "")
        row["현황_사업자등록번호"] = ov.get("사업자등록번호", "")
        row["현황_법인번호"] = ov.get("법인번호", "")
        row["현황_공공민간구분"] = ov.get("공공/민간 구분", "")
        row["현황_공동도급"] = ov.get("공동도급", "")

        prog_con = next((p for p in parsed.get("progress", []) if "시공" in p.get("부문", "")), {})
        prog_spv = next((p for p in parsed.get("progress", []) if "건설사업관리" in p.get("부문", "")), {})

        row["공사현황_시공_구분"] = prog_con.get("구분", "")
        row["공사현황_시공_계약일"] = prog_con.get("계약일", "")
        row["공사현황_시공_도급액"] = prog_con.get("도급액", "")
        row["공사현황_시공_공사기간"] = prog_con.get("공사기간", "")
        row["공사현황_시공_현재공정"] = prog_con.get("현재공정(%)", "")

        row["공사현황_건설사업관리_구분"] = prog_spv.get("구분", "")
        row["공사현황_건설사업관리_계약일"] = prog_spv.get("계약일", "")
        row["공사현황_건설사업관리_도급액"] = prog_spv.get("도급액", "")
        row["공사현황_건설사업관리_공사기간"] = prog_spv.get("공사기간", "")
        row["공사현황_건설사업관리_현재공정"] = prog_spv.get("현재공정(%)", "")

        permit_0 = parsed.get("permit", [{}])[0] if parsed.get("permit") else {}
        row["사업인허가_승인일자"] = permit_0.get("승인일자", "")
        row["사업인허가_승인기관"] = permit_0.get("승인기관", "")
        row["사업인허가_담당부서"] = permit_0.get("담당부서", "")
        row["사업인허가_전화번호"] = permit_0.get("전화번호", "")

        constructors = parsed.get("constructors", [])
        for idx in range(1, max_con_cnt + 1):
            k_sfx = f"_{idx}" if max_con_cnt > 1 else ""
            c_item = constructors[idx - 1] if idx <= len(constructors) else {}
            for field in ("업체명", "대표자", "업종", "면허번호", "주소", "전화번호", "팩스번호", "분담공종", "지분(%)"):
                val = c_item.get(field, "")
                if field == "업체명" and not val and idx == 1:
                    val = list_info.get("목록_시공자", "")
                row[f"시공업체{k_sfx}_{field}"] = val

        supervisors = parsed.get("supervisors", [])
        for idx in range(1, max_spv_cnt + 1):
            k_sfx = f"_{idx}" if max_spv_cnt > 1 else ""
            s_item = supervisors[idx - 1] if idx <= len(supervisors) else {}
            for field in ("업체명", "대표자", "업종", "면허번호", "주소", "전화번호", "팩스번호", "분담공종", "지분(%)"):
                row[f"건설사업관리{k_sfx}_{field}"] = s_item.get(field, "")

        eng_con_list = parsed.get("engineers_con", [])
        for idx in range(1, max_eng_con_cnt + 1):
            k_sfx = f"_{idx}" if max_eng_con_cnt > 1 else ""
            e_item = eng_con_list[idx - 1] if idx <= len(eng_con_list) else {}
            for field in ("기술인 구분", "성명", "소속회사", "자격종류", "자격증번호", "휴대전화번호"):
                row[f"건설기술인{k_sfx}_{field}"] = e_item.get(field, "")

        eng_spv_list = parsed.get("engineers_spv", [])
        for idx in range(1, max_eng_spv_cnt + 1):
            k_sfx = f"_{idx}" if max_eng_spv_cnt > 1 else ""
            e_item = eng_spv_list[idx - 1] if idx <= len(eng_spv_list) else {}
            for field in ("기술인 구분", "성명", "소속회사", "자격종류", "자격증번호", "휴대전화번호"):
                row[f"감독기술자{k_sfx}_{field}"] = e_item.get(field, "")

        row["하단_현장점검확인서"] = parsed.get("confirm_file", "")
        memo_list = parsed.get("memos", [])
        for idx in range(1, max_memo_cnt + 1):
            k_sfx = f"_{idx}" if max_memo_cnt > 1 else ""
            m_item = memo_list[idx - 1] if idx <= len(memo_list) else {}
            for field in ("내용", "점검일", "작성자", "작성일"):
                row[f"현장점검메모{k_sfx}_{field}"] = m_item.get(field, "")

        inspections = parsed.get("inspections", [])
        for i in range(1, max_insp_cnt + 1):
            p = f"점검결과{i}"
            if i <= len(inspections):
                insp = inspections[i - 1]
                cnfmn = insp.get("지적사항", {})
                managt = insp.get("조치결과", {})
                managt_rslt = insp.get("조치결과확인", {})

                row[f"{p}_지적_제목"] = cnfmn.get("제목", insp.get("패널제목", ""))
                row[f"{p}_지적_조치계획"] = cnfmn.get("조치계획", "")
                row[f"{p}_지적_지적대상"] = cnfmn.get("지적대상", "")
                row[f"{p}_지적_점검항목"] = cnfmn.get("점검항목", "")
                row[f"{p}_지적_통보일"] = cnfmn.get("통보일", "")
                row[f"{p}_지적_작성자"] = cnfmn.get("작성자", "")
                row[f"{p}_지적_사진"] = cnfmn.get("사진", "")

                row[f"{p}_조치_조치사항"] = managt.get("조치사항", "")
                row[f"{p}_조치_조치일"] = managt.get("조치일", "")
                row[f"{p}_조치_통보일"] = managt.get("통보일", "")
                row[f"{p}_조치_작성자"] = managt.get("작성자", "")
                row[f"{p}_조치_사진"] = managt.get("사진", "")

                row[f"{p}_확인_조치이행결과"] = managt_rslt.get("조치이행결과", "")
                row[f"{p}_확인_확인일"] = managt_rslt.get("확인일", "")
                row[f"{p}_확인_통보일"] = managt_rslt.get("통보일", "")
                row[f"{p}_확인_작성자"] = managt_rslt.get("작성자", "")
                row[f"{p}_확인_사진"] = managt_rslt.get("사진", "")
            else:
                for sub_key in (
                    "지적_제목",
                    "지적_조치계획",
                    "지적_지적대상",
                    "지적_점검항목",
                    "지적_통보일",
                    "지적_작성자",
                    "지적_사진",
                    "조치_조치사항",
                    "조치_조치일",
                    "조치_통보일",
                    "조치_작성자",
                    "조치_사진",
                    "확인_조치이행결과",
                    "확인_확인일",
                    "확인_통보일",
                    "확인_작성자",
                    "확인_사진",
                ):
                    row[f"{p}_{sub_key}"] = ""

        rows_data.append(row)

    return col_defs, rows_data


def generate_excel_bytes(collected_sites):
    col_defs, rows_data = build_column_schema_and_rows(collected_sites)

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "현장점검결과_전체"
    ws1.views.sheetView[0].showGridLines = True
    ws1.freeze_panes = "D3"

    basic_grp_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    basic_sub_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    insp_odd_grp_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    insp_odd_sub_fill = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")
    insp_even_grp_fill = PatternFill(start_color="385723", end_color="385723", fill_type="solid")
    insp_even_sub_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    grp_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    sub_font = Font(name="맑은 고딕", size=10, bold=True, color="1F2937")
    data_font = Font(name="맑은 고딕", size=10, color="111827")
    alt_row_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    for col_idx, (grp_name, sub_name, _) in enumerate(col_defs, start=1):
        c1 = ws1.cell(row=1, column=col_idx, value=grp_name)
        c2 = ws1.cell(row=2, column=col_idx, value=sub_name)

        if grp_name.startswith("점검결과"):
            try:
                insp_num = int(grp_name.replace("점검결과", "").strip())
            except ValueError:
                insp_num = 1
            g_fill = insp_odd_grp_fill if insp_num % 2 == 1 else insp_even_grp_fill
            s_fill = insp_odd_sub_fill if insp_num % 2 == 1 else insp_even_sub_fill
        else:
            g_fill = basic_grp_fill
            s_fill = basic_sub_fill

        c1.fill = g_fill
        c1.font = grp_font
        c1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c1.border = thin_border

        c2.fill = s_fill
        c2.font = sub_font
        c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c2.border = thin_border

    merge_start = 1
    for col_idx in range(2, len(col_defs) + 2):
        prev_grp = col_defs[col_idx - 2][0]
        curr_grp = col_defs[col_idx - 1][0] if col_idx <= len(col_defs) else None
        if curr_grp != prev_grp:
            if col_idx - 1 > merge_start:
                ws1.merge_cells(start_row=1, start_column=merge_start, end_row=1, end_column=col_idx - 1)
            merge_start = col_idx

    ws1.row_dimensions[1].height = 24
    ws1.row_dimensions[2].height = 26

    for r_offset, row_dict in enumerate(rows_data, start=3):
        for col_idx, (_, _, key) in enumerate(col_defs, start=1):
            val = row_dict.get(key, "")
            cell = ws1.cell(row=r_offset, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin_border
            if r_offset % 2 == 0:
                cell.fill = alt_row_fill
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
        ws1.row_dimensions[r_offset].height = 22

    for col_idx in range(1, len(col_defs) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(2, ws1.max_row + 1):
            val_str = str(ws1.cell(row=row_idx, column=col_idx).value or "")
            visual_len = sum(2 if ord(ch) > 127 else 1 for ch in val_str)
            max_len = max(max_len, visual_len)
        ws1.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 42)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), len(col_defs)


# ==========================================
# 세션 상태 초기화
# ==========================================
if "driver" not in st.session_state:
    st.session_state.driver = None
if "collected_sites" not in st.session_state:
    st.session_state.collected_sites = []
if "stop_requested" not in st.session_state:
    st.session_state.stop_requested = False


# ==========================================
# 화면 UI 구성
# ==========================================
st.markdown("<div class='main-title'>🏗️ CSI 현장점검 결과 자동 수집 & 엑셀 변환기</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='sub-title'>실시간 크롬 창에서 직접 로그인 및 검색 조건을 설정하고, [수집 시작]을 누르면 원본 파이썬 방식 그대로 자동 순회 수집합니다.</div>",
    unsafe_allow_html=True,
)

# ----------------------------------------------------
# 1단계: 브라우저 실행 및 로그인
# ----------------------------------------------------
with st.container():
    st.markdown("### 1️⃣ 브라우저 열기 및 로그인")
    st.caption("아래 [크롬 브라우저 열기]를 누르면 내 컴퓨터 화면에 실제 크롬 창이 뜹니다. 자동 로그인을 원하시면 아이디/비번을 입력하시고, 직접 로그인하셔도 됩니다.")

    col1, col2, col3 = st.columns([2, 2, 1.5])
    with col1:
        u_id = st.text_input("CSI 아이디 (선택)", placeholder="아이디 입력")
    with col2:
        u_pw = st.text_input("CSI 비밀번호 (선택)", type="password", placeholder="비밀번호 입력")
    with col3:
        st.write("")
        if st.button("🌐 CSI 크롬 브라우저 열기", type="primary", use_container_width=True):
            if st.session_state.driver:
                try:
                    st.session_state.driver.quit()
                except Exception:
                    pass
            with st.spinner("크롬 브라우저를 띄우는 중입니다..."):
                drv = create_driver()
                drv.get("https://www.csi.go.kr/spt/sptEndList.do")
                if u_id.strip() and u_pw.strip():
                    auto_login_csi(drv, u_id.strip(), u_pw.strip())
                st.session_state.driver = drv
                st.success("✅ 크롬 브라우저가 열렸습니다! 화면을 확인해주세요.")

# ----------------------------------------------------
# 2단계: 사용자 검색 설정 안내
# ----------------------------------------------------
with st.container():
    st.markdown("### 2️⃣ CSI 크롬 창에서 검색 조건 설정 및 검색 클릭")
    st.info(
        "💡 **진행 방법:**\n"
        "1. 방금 열린 크롬 브라우저에서 **로그인**을 완료하세요.\n"
        "2. 원하시는 **현장점검 기관, 지역, 기간 등 검색 조건**을 직접 선택하고 **[검색하기]** 버튼을 누르세요.\n"
        "3. 하단에 검색 결과 목록(10개씩)이 정상적으로 표출된 것을 확인하셨다면 아래 3단계로 이동하세요."
    )

# ----------------------------------------------------
# 3단계: 자동 수집 실행 및 엑셀 다운로드
# ----------------------------------------------------
with st.container():
    st.markdown("### 3️⃣ 자동 수집 실행 및 엑셀 저장")

    c_p1, c_p2, c_btn1, c_btn2 = st.columns([1, 1, 1.8, 1.2])
    with c_p1:
        start_p = st.number_input("수집 시작 페이지", min_value=1, value=1)
    with c_p2:
        end_p = st.number_input("수집 종료 페이지", min_value=1, value=10)
    with c_btn1:
        st.write("")
        start_collect = st.button("🚀 현재 화면부터 자동 수집 시작", type="primary", use_container_width=True)
    with c_btn2:
        st.write("")
        if st.button("🛑 즉시 정지 및 저장", use_container_width=True):
            st.session_state.stop_requested = True

    if start_collect:
        if not st.session_state.driver:
            st.error("먼저 1단계에서 [CSI 크롬 브라우저 열기]를 눌러주세요.")
        else:
            st.session_state.stop_requested = False
            driver = st.session_state.driver
            collected = list(st.session_state.collected_sites)
            progress_bar = st.progress(0)
            status_box = st.empty()
            log_expander = st.expander("📜 실시간 수집 콘솔", expanded=True)
            log_area = log_expander.empty()
            logs = []

            total_pages = max(1, int(end_p) - int(start_p) + 1)
            interrupted = False

            try:
                # 시작 페이지로 이동
                if int(start_p) > 1:
                    driver.execute_script(f"goPage('{int(start_p)}');")
                    time.sleep(2.5)

                for p_idx, curr_page in enumerate(range(int(start_p), int(end_p) + 1), start=1):
                    if st.session_state.stop_requested:
                        interrupted = True
                        break

                    status_box.info(f"🔄 **[페이지 {curr_page} / {int(end_p)}]** 수집 진행 중...")
                    time.sleep(1.8)

                    # 1. 목록 테이블의 실제 현장 행 인덱스 추출 (원본 방식)
                    raw_rows = driver.find_elements(By.XPATH, "//table//tbody//tr[td]")
                    filtered_indices = []
                    for idx, r in enumerate(raw_rows):
                        cols = r.find_elements(By.TAG_NAME, "td")
                        if len(cols) >= 7 and cols[0].text.strip().isdigit():
                            filtered_indices.append(idx)

                    row_count = len(filtered_indices)
                    if row_count == 0:
                        logs.append(f"[페이지 {curr_page}] 현장 목록이 없어 수집을 마칩니다.")
                        log_area.code("\n".join(logs[-20:]))
                        break

                    logs.append(f"📄 [페이지 {curr_page}] 총 {row_count}개 현장 순회 시작")
                    log_area.code("\n".join(logs[-20:]))

                    # 2. 순차적으로 클릭 → 파싱 → 복귀 (원본 collector 검증 방식)
                    for order_idx, orig_idx in enumerate(filtered_indices):
                        if st.session_state.stop_requested:
                            interrupted = True
                            break

                        try:
                            # 클릭 직전에 최신 행들을 다시 가져옴 (Stale 방지)
                            curr_rows = driver.find_elements(By.XPATH, "//table//tbody//tr[td]")
                            if orig_idx >= len(curr_rows):
                                break
                            target_row = curr_rows[orig_idx]
                            cols = target_row.find_elements(By.TAG_NAME, "td")

                            list_info = {
                                "No": cols[0].text.strip(),
                                "공사명": cols[1].text.strip() if len(cols) > 1 else "",
                                "점검명": cols[2].text.strip() if len(cols) > 2 else "",
                                "공사기간": cols[3].text.strip() if len(cols) > 3 else "",
                                "현장주소": cols[4].text.strip() if len(cols) > 4 else "",
                                "현장점검예정일": cols[5].text.strip() if len(cols) > 5 else "",
                                "현장점검종류": cols[6].text.strip() if len(cols) > 6 else "",
                                "발주자": cols[7].text.strip() if len(cols) > 7 else "",
                                "목록_시공자": cols[8].text.strip() if len(cols) > 8 else "",
                                "지적사항_조치완료": cols[9].text.strip() if len(cols) > 9 else "",
                            }

                            a_tags = target_row.find_elements(By.TAG_NAME, "a")
                            if not a_tags:
                                collected.append({"list_info": list_info, "parsed": {}})
                                continue

                            link_el = a_tags[0]
                            main_win = driver.current_window_handle
                            old_wins = driver.window_handles

                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link_el)
                            time.sleep(0.2)
                            driver.execute_script("arguments[0].click();", link_el)
                            time.sleep(2.0)

                            new_wins = driver.window_handles

                            # 새 창/탭으로 열린 경우
                            if len(new_wins) > len(old_wins):
                                new_tab = [w for w in new_wins if w not in old_wins][0]
                                driver.switch_to.window(new_tab)
                                parsed_detail = parse_detail_page(driver)
                                driver.close()
                                driver.switch_to.window(main_win)
                            # 동일 창에서 페이지 이동된 경우 (원본 방식 복귀)
                            else:
                                parsed_detail = parse_detail_page(driver)
                                driver.back()
                                time.sleep(2.0)

                            collected.append({"list_info": list_info, "parsed": parsed_detail})
                            st.session_state.collected_sites = collected

                            insp_cnt = len(parsed_detail.get("inspections", []))
                            logs.append(
                                f"  └ [{order_idx + 1}/{row_count}] No.{list_info['No']} '{list_info['공사명']}' 완료 (점검결과 {insp_cnt}건)"
                            )
                            log_area.code("\n".join(logs[-20:]))

                            done_ratio = ((p_idx - 1) + (order_idx + 1) / max(1, row_count)) / total_pages
                            progress_bar.progress(min(1.0, done_ratio))

                        except Exception as row_err:
                            logs.append(f"  └ [오류] {order_idx + 1}번째 현장 스킵: {row_err}")
                            log_area.code("\n".join(logs[-20:]))
                            try:
                                while len(driver.window_handles) > 1:
                                    driver.switch_to.window(driver.window_handles[-1])
                                    driver.close()
                                driver.switch_to.window(driver.window_handles[0])
                            except Exception:
                                pass

                    # 다음 페이지 이동
                    if not interrupted and curr_page < int(end_p):
                        next_p = curr_page + 1
                        try:
                            driver.execute_script(f"goPage('{next_p}');")
                            time.sleep(2.5)
                        except Exception:
                            page_btns = driver.find_elements(
                                By.XPATH, f"//a[contains(@href, \"goPage('{next_p}')\")]"
                            )
                            if page_btns:
                                driver.execute_script("arguments[0].click();", page_btns[0])
                                time.sleep(2.5)
                            else:
                                break

                if interrupted:
                    status_box.warning(
                        f"🛑 사용자에 의해 정지되었습니다. 현재까지 수집된 **{len(st.session_state.collected_sites)}건**을 바로 다운로드할 수 있습니다."
                    )
                else:
                    progress_bar.progress(1.0)
                    status_box.success(
                        f"🎉 총 **{len(collected)}개** 현장 수집 완료! 아래 버튼으로 엑셀을 저장하세요."
                    )

            except Exception as e:
                status_box.warning(
                    f"⚠️ 안내: {e} (현재까지 수집된 **{len(st.session_state.collected_sites)}건** 다운로드 가능)"
                )

    if st.session_state.collected_sites:
        excel_bytes, total_cols = generate_excel_bytes(st.session_state.collected_sites)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label=f"📥 엑셀 파일 다운로드 (총 {len(st.session_state.collected_sites)}개 현장 / {total_cols}개 열)",
            data=excel_bytes,
            file_name=f"CSI_현장점검결과_수집_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
