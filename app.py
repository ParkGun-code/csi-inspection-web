import io
import os
import re
import time
from datetime import datetime, date
import pandas as pd
from bs4 import BeautifulSoup

import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ==========================================
# 웹페이지 기본 설정
# ==========================================
st.set_page_config(
    page_title="CSI 현장점검 결과 자동 수집 & 엑셀 변환기",
    page_icon="🏗️",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 1400px; }
    .main-title { font-size: 26px; font-weight: 700; color: #0f172a; margin-bottom: 4px; }
    .sub-title { font-size: 14px; color: #475569; margin-bottom: 20px; border-bottom: 1px solid #e2e8f0; padding-bottom: 12px; }
    .step-box { background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; margin-bottom: 16px; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================
# 1. 크롬 브라우저 생성 (클라우드 서버 & 로컬 PC 자동 호환)
# ==========================================
def create_driver(headless=True):
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1600,1050")
    options.add_argument("--ignore-certificate-errors")

    if headless:
        options.add_argument("--headless=new")
    else:
        options.add_argument("--start-maximized")

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
# 2. CSI 자동 로그인 및 검색 제어 함수
# ==========================================
def login_to_csi(driver, user_id, user_pw):
    driver.get("https://www.csi.go.kr/main.do")
    time.sleep(2.0)
    dismiss_alerts(driver)

    id_selectors = [
        "input[name='user_id']",
        "input[name='userId']",
        "input[name='id']",
        "input#user_id",
        "input#userId",
        "input[placeholder*='아이디']",
        "input[title*='아이디']",
    ]
    pw_selectors = [
        "input[name='pswd']",
        "input[name='password']",
        "input[name='user_pw']",
        "input[name='userPw']",
        "input[type='password']",
    ]

    id_input = None
    for sel in id_selectors:
        elems = [e for e in driver.find_elements(By.CSS_SELECTOR, sel) if e.is_displayed()]
        if elems:
            id_input = elems[0]
            break

    if not id_input:
        login_open_btns = driver.find_elements(
            By.XPATH, "//a[contains(text(), '로그인')] | //button[contains(text(), '로그인')]"
        )
        for btn in login_open_btns:
            if btn.is_displayed():
                try:
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(1.0)
                except Exception:
                    pass
                break
        for sel in id_selectors:
            elems = [e for e in driver.find_elements(By.CSS_SELECTOR, sel) if e.is_displayed()]
            if elems:
                id_input = elems[0]
                break

    pw_input = None
    for sel in pw_selectors:
        elems = [e for e in driver.find_elements(By.CSS_SELECTOR, sel) if e.is_displayed()]
        if elems:
            pw_input = elems[0]
            break

    if not id_input or not pw_input:
        return False, "로그인 입력창을 찾지 못했습니다. CSI 사이트 접속 상태를 확인해주세요."

    id_input.clear()
    id_input.send_keys(user_id)
    pw_input.clear()
    pw_input.send_keys(user_pw)

    try:
        driver.execute_script("if(typeof go_login === 'function') { go_login(); }")
    except Exception:
        pass

    time.sleep(0.5)
    alert_msg = dismiss_alerts(driver)
    if alert_msg:
        return False, f"CSI 알림: {alert_msg}"

    if "님으로 로그인" not in driver.page_source and "로그아웃" not in driver.page_source:
        try:
            pw_input.send_keys(Keys.RETURN)
        except Exception:
            pass
        time.sleep(2.0)
        alert_msg = dismiss_alerts(driver)
        if alert_msg:
            return False, f"CSI 로그인 실패: {alert_msg}"

    time.sleep(1.5)
    html = driver.page_source
    if "님으로 로그인" in html or "로그아웃" in html or "go_logout" in html:
        soup = BeautifulSoup(html, "html.parser")
        user_span = soup.find("span", class_="username")
        welcome_txt = user_span.get_text(strip=True) if user_span else "로그인 성공"
        return True, welcome_txt

    return False, "로그인 후 사용자 정보를 확인할 수 없습니다. 아이디와 비밀번호를 다시 확인해주세요."


def get_available_inspection_orgs(driver):
    """
    CSI 목록 화면에 렌더링된 '현장점검 기관' 드롭다운(select) 옵션 목록을 가져옵니다.
    """
    org_list = []
    try:
        select_candidates = driver.find_elements(
            By.CSS_SELECTOR,
            "select[name*='inst'], select[name*='agency'], select[name*='org'], select#inst_cd, select#chck_inst",
        )
        for sel in select_candidates:
            options = sel.find_elements(By.TAG_NAME, "option")
            texts = [o.text.strip() for o in options if o.text.strip()]
            if any(k in " ".join(texts) for k in ["국토교통부", "관리원", "지방국토", "한국토지주택공사", "도로공사", "철도"]):
                for opt in options:
                    t = opt.text.strip()
                    v = opt.get_attribute("value") or ""
                    if t and t not in [x[0] for x in org_list]:
                        org_list.append((t, v))
                break
    except Exception:
        pass
    return org_list


def navigate_and_search(driver, target_list_url, search_params):
    driver.get(target_list_url)
    time.sleep(2.0)
    dismiss_alerts(driver)

    # 1. 셀렉트 박스(현장점검 기관 등) 또는 텍스트 입력창 자동 선택/주입 스크립트
    js_set_fields = """
    var params = arguments[0];
    var orgName = params['chck_inst_nm'] || '';

    // 현장점검 기관 드롭다운 처리
    if (orgName) {
        var selects = document.querySelectorAll('select');
        for (var s = 0; s < selects.length; s++) {
            var sel = selects[s];
            for (var o = 0; o < sel.options.length; o++) {
                if (sel.options[o].text.indexOf(orgName) >= 0) {
                    sel.selectedIndex = o;
                    sel.value = sel.options[o].value;
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                    break;
                }
            }
        }
    }

    // 일반 파라미터 폼 필드 입력
    for (var key in params) {
        var val = params[key];
        if (!val) continue;
        var elems = document.querySelectorAll('[name="' + key + '"], #' + key);
        for (var i = 0; i < elems.length; i++) {
            elems[i].value = val;
            elems[i].dispatchEvent(new Event('change', { bubbles: true }));
        }
    }
    """
    driver.execute_script(js_set_fields, search_params)
    time.sleep(0.3)

    # 2. 검색 버튼 클릭
    searched = False
    search_btns = driver.find_elements(
        By.XPATH,
        "//button[contains(., '검색')] | //a[contains(., '검색')] | //input[@type='button' and contains(@value, '검색')] | //input[@type='submit' and contains(@value, '검색')]",
    )
    for btn in search_btns:
        onclick_attr = btn.get_attribute("onclick") or ""
        title_attr = btn.get_attribute("title") or ""
        if "go_sch_search" in onclick_attr or "통합검색" in title_attr:
            continue
        if btn.is_displayed():
            try:
                driver.execute_script("arguments[0].click();", btn)
                searched = True
                break
            except Exception:
                pass

    if not searched:
        try:
            driver.execute_script(
                "if (typeof goSearch === 'function') { goSearch(); } "
                "else if (typeof fn_search === 'function') { fn_search(); } "
                "else if (typeof goPage === 'function') { goPage('1'); } "
                "else if (document.forms['main']) { document.forms['main'].submit(); }"
            )
        except Exception:
            pass

    time.sleep(2.5)
    dismiss_alerts(driver)
    total_pages, total_items = get_total_page_info(driver)
    return total_pages, total_items


# ==========================================
# 3. 전체 페이지 수 및 현재 목록 미리보기 추출
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


def get_current_page_list_preview(driver):
    rows_preview = []
    try:
        raw_rows = driver.find_elements(
            By.XPATH,
            "//table[contains(@class, 'tbl') or contains(@class, 'list') or contains(@class, 'board')]//tbody//tr[td]",
        )
        if not raw_rows:
            raw_rows = driver.find_elements(By.XPATH, "//table//tbody//tr[td]")

        for r in raw_rows:
            text = r.text.strip()
            if any(k in text for k in ["국토안전", "현장대리인", "현장점검 기관", "담당 현장점검"]):
                continue
            cols = r.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 5:
                if "조회된" in cols[0].text or "없습니다" in cols[0].text:
                    continue
                rows_preview.append(
                    {
                        "No": cols[0].text.strip(),
                        "공사명": cols[1].text.strip() if len(cols) > 1 else "",
                        "점검명": cols[2].text.strip() if len(cols) > 2 else "",
                        "공사기간": cols[3].text.strip() if len(cols) > 3 else "",
                        "현장주소": cols[4].text.strip() if len(cols) > 4 else "",
                        "현장점검예정일": cols[5].text.strip() if len(cols) > 5 else "",
                        "현장점검종류": cols[6].text.strip() if len(cols) > 6 else "",
                        "발주자": cols[7].text.strip() if len(cols) > 7 else "",
                        "시공자": cols[8].text.strip() if len(cols) > 8 else "",
                        "조치완료": cols[9].text.strip() if len(cols) > 9 else "",
                    }
                )
    except Exception:
        pass
    return rows_preview


# ==========================================
# 4. 텍스트 정리 및 테이블 파싱 유틸리티
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
# 5. 현장점검 결과(#cnfrmnView) HTML 전체 파싱
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
# 6. 각 항목을 개별 셀(Column)로 1:1 매핑하는 빌더
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
# 7. 상세 페이지 진입 및 파싱
# ==========================================
def parse_detail_page(driver):
    time.sleep(1.0)

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    in_iframe = False
    if iframes:
        try:
            driver.switch_to.frame(iframes[0])
            in_iframe = True
        except Exception:
            pass

    try:
        tab_links = driver.find_elements(By.CSS_SELECTOR, "a[href*='#cnfrmnView']")
        if tab_links:
            parent_li = tab_links[0].find_element(By.XPATH, "./..")
            li_class = parent_li.get_attribute("class") or ""
            if "active" not in li_class:
                driver.execute_script("arguments[0].click();", tab_links[0])
            else:
                tab_div = driver.find_elements(By.CSS_SELECTOR, "#tab table")
                if not tab_div:
                    driver.execute_script(
                        "window.location.hash = '#cnfrmnView'; window.onhashchange && window.onhashchange();"
                    )

        WebDriverWait(driver, 10).until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "#tab table")) >= 3)
        time.sleep(0.6)
    except Exception:
        time.sleep(1.2)

    html = driver.page_source

    if in_iframe:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass

    return parse_csi_detail_html(html)


# ==========================================
# 세션 상태 초기화
# ==========================================
if "driver" not in st.session_state:
    st.session_state.driver = None
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_welcome" not in st.session_state:
    st.session_state.user_welcome = ""
if "total_pages" not in st.session_state:
    st.session_state.total_pages = 0
if "total_items" not in st.session_state:
    st.session_state.total_items = 0
if "preview_rows" not in st.session_state:
    st.session_state.preview_rows = []
if "collected_sites" not in st.session_state:
    st.session_state.collected_sites = []
if "discovered_orgs" not in st.session_state:
    st.session_state.discovered_orgs = []


# ==========================================
# 화면 UI 구성
# ==========================================
st.markdown("<div class='main-title'>🏗️ CSI 현장점검 결과 자동 수집 & 엑셀 변환 웹앱</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='sub-title'>파이썬 설치 없이 웹에서 바로 CSI에 로그인하여 검색 조건 설정 후 전수 데이터를 엑셀(각 항목 개별 셀)로 추출합니다.</div>",
    unsafe_allow_html=True,
)

tab_auto, tab_html = st.tabs(["🌐 1. CSI 사이트 직접 연결 수집 (아이디/비밀번호)", "📂 2. 저장된 HTML 파일 즉시 변환"])

with tab_auto:
    # ----------------------------------------------------
    # STEP 1: CSI 로그인 및 브라우저 세션 시작
    # ----------------------------------------------------
    with st.expander("🔑 1단계: CSI 로그인", expanded=not st.session_state.logged_in):
        col_id, col_pw, col_opt, col_btn = st.columns([2, 2, 1.5, 1.5])
        with col_id:
            csi_id = st.text_input("CSI 아이디", placeholder="아이디 입력", key="csi_id_input")
        with col_pw:
            csi_pw = st.text_input("CSI 비밀번호", type="password", placeholder="비밀번호 입력", key="csi_pw_input")
        with col_opt:
            is_linux_server = os.path.exists("/usr/bin/chromium") or os.path.exists("/usr/bin/chromedriver")
            headless_mode = st.checkbox(
                "백그라운드 실행(서버 표준)",
                value=True,
                help="웹 서버 배포 시에는 항상 켜두세요. 본인 PC에서 크롬 창을 직접 띄워 보려면 체크를 해제하세요.",
                disabled=is_linux_server,
            )
        with col_btn:
            st.write("")
            if st.button("🚀 CSI 접속 및 로그인", type="primary", use_container_width=True):
                if not csi_id.strip() or not csi_pw.strip():
                    st.warning("CSI 아이디와 비밀번호를 입력해주세요.")
                else:
                    with st.spinner("CSI(csi.go.kr)에 접속하여 로그인 중입니다..."):
                        try:
                            if st.session_state.driver:
                                try:
                                    st.session_state.driver.quit()
                                except Exception:
                                    pass
                            drv = create_driver(headless=headless_mode)
                            ok, msg = login_to_csi(drv, csi_id.strip(), csi_pw.strip())
                            if ok:
                                st.session_state.driver = drv
                                st.session_state.logged_in = True
                                st.session_state.user_welcome = msg
                                # 로그인 직후 점검완료 페이지에서 '현장점검 기관' 드롭다운 옵션 자동 탐색
                                try:
                                    drv.get("https://www.csi.go.kr/spt/sptEndList.do")
                                    time.sleep(1.5)
                                    st.session_state.discovered_orgs = get_available_inspection_orgs(drv)
                                except Exception:
                                    pass
                                st.success(f"✅ 로그인 완료: {msg}")
                                st.rerun()
                            else:
                                drv.quit()
                                st.session_state.driver = None
                                st.session_state.logged_in = False
                                st.error(f"❌ {msg}")
                        except Exception as e:
                            st.error(f"❌ 브라우저 실행 오류: {e}")

        if st.session_state.logged_in:
            col_s1, col_s2 = st.columns([4, 1])
            with col_s1:
                st.success(f"🟢 현재 CSI 연결 상태: **{st.session_state.user_welcome}**")
            with col_s2:
                if st.button("🔌 세션 종료(로그아웃)", use_container_width=True):
                    if st.session_state.driver:
                        try:
                            st.session_state.driver.quit()
                        except Exception:
                            pass
                    st.session_state.driver = None
                    st.session_state.logged_in = False
                    st.session_state.total_pages = 0
                    st.session_state.total_items = 0
                    st.session_state.preview_rows = []
                    st.session_state.discovered_orgs = []
                    st.rerun()

    # ----------------------------------------------------
    # STEP 2: 검색 조건 설정 및 목록 조회 (현장점검 기관 항목 추가)
    # ----------------------------------------------------
    with st.expander("🔍 2단계: 검색 조건 설정 및 대상 조회", expanded=st.session_state.logged_in):
        # 1줄: 조회 메뉴, 현장점검 기관(주요 검색조건), 지역
        c1, c_org, c2 = st.columns([1.5, 2.5, 1.5])
        with c1:
            board_type = st.selectbox(
                "조회 메뉴 선택",
                ["점검완료 (/spt/sptEndList.do)", "점검대상 (/spt/sptTrgtList.do)"],
            )

        with c_org:
            # 기본 프리셋 기관 목록 + CSI 사이트에서 자동 읽어온 드롭다운 목록 통합
            preset_orgs = [
                "전체 (선택 안함)",
                "국토교통부",
                "국토안전관리원",
                "서울지방국토관리청",
                "원주지방국토관리청",
                "대전지방국토관리청",
                "익산지방국토관리청",
                "부산지방국토관리청",
                "한국토지주택공사(LH)",
                "한국도로공사",
                "국가철도공단",
                "한국철도공사",
                "한국수자원공사",
                "직접 입력",
            ]
            if st.session_state.discovered_orgs:
                dynamic_labels = [o[0] for o in st.session_state.discovered_orgs if o[0] not in preset_orgs]
                preset_orgs = ["전체 (선택 안함)"] + dynamic_labels + [p for p in preset_orgs[1:] if p not in dynamic_labels]

            selected_org_choice = st.selectbox("🏛️ 현장점검 기관", preset_orgs)
            if selected_org_choice == "직접 입력":
                custom_org_name = st.text_input("현장점검 기관명 직접 입력", placeholder="예: 대전지방국토관리청")
                chck_inst_val = custom_org_name.strip()
            elif selected_org_choice == "전체 (선택 안함)":
                chck_inst_val = ""
            else:
                chck_inst_val = selected_org_choice

        with c2:
            regions = [
                "전체",
                "서울",
                "부산",
                "대구",
                "인천",
                "광주",
                "대전",
                "울산",
                "세종",
                "경기",
                "강원",
                "충북",
                "충남",
                "전북",
                "전남",
                "경북",
                "경남",
                "제주",
            ]
            search_region = st.selectbox("지역", regions)

        # 2줄: 점검예정일 기간 및 추가 상세 검색어
        chck_col1, chck_col2, c4, c5 = st.columns(4)
        with chck_col1:
            chck_start = st.date_input("현장점검예정일 (시작)", value=date(date.today().year, 1, 1))
        with chck_col2:
            chck_end = st.date_input("현장점검예정일 (종료)", value=date.today())
        with c4:
            const_nm = st.text_input("공사명 (선택)", placeholder="검색할 공사명")
        with c5:
            chck_nm = st.text_input("점검명 (선택)", placeholder="예: 8월 건설공사")

        c6, c7 = st.columns(2)
        with c6:
            ord_corp_nm = st.text_input("발주자 (선택)", placeholder="예: 천안시청")
        with c7:
            con_corp = st.text_input("시공자 (선택)", placeholder="예: 롯데건설")

        if st.button("🔎 위 조건으로 CSI 목록 검색하기", type="primary", disabled=not st.session_state.logged_in):
            target_url = (
                "https://www.csi.go.kr/spt/sptEndList.do"
                if "sptEndList" in board_type
                else "https://www.csi.go.kr/spt/sptTrgtList.do"
            )
            params = {
                "chck_inst_nm": chck_inst_val,
                "inst_nm": chck_inst_val,
                "chck_inst": chck_inst_val,
                "search_region": "" if search_region == "전체" else search_region,
                "chck_start_ymd": chck_start.strftime("%Y%m%d") if chck_start else "",
                "chck_end_ymd": chck_end.strftime("%Y%m%d") if chck_end else "",
                "const_nm": const_nm.strip(),
                "chck_nm": chck_nm.strip(),
                "ord_corp_nm": ord_corp_nm.strip(),
                "con_corp": con_corp.strip(),
            }
            with st.spinner(f"'{chck_inst_val or '전체'}' 기관 및 설정하신 조건으로 CSI 목록을 검색 중입니다..."):
                try:
                    t_pages, t_items = navigate_and_search(st.session_state.driver, target_url, params)
                    st.session_state.total_pages = t_pages
                    st.session_state.total_items = t_items
                    st.session_state.preview_rows = get_current_page_list_preview(st.session_state.driver)
                except Exception as e:
                    st.error(f"검색 실행 중 오류 발생: {e}")

        if st.session_state.preview_rows or st.session_state.total_items > 0:
            st.info(
                f"📊 **검색 결과 감지**: 총 **{st.session_state.total_items}개** 현장 / 전체 **{st.session_state.total_pages}페이지**가 조회되었습니다."
            )
            st.dataframe(pd.DataFrame(st.session_state.preview_rows), use_container_width=True, hide_index=True)

    # ----------------------------------------------------
    # STEP 3: 전수 데이터 수집 및 엑셀 다운로드
    # ----------------------------------------------------
    with st.expander("📥 3단계: 현장점검 결과 수집 실행 및 엑셀 다운로드", expanded=True):
        max_p = max(1, st.session_state.total_pages)
        p_col1, p_col2, p_col3 = st.columns([1, 1, 2])
        with p_col1:
            start_page = st.number_input("수집 시작 페이지", min_value=1, max_value=max_p, value=1)
        with p_col2:
            end_page = st.number_input("수집 종료 페이지", min_value=1, max_value=max_p, value=max_p)
        with p_col3:
            st.write("")
            start_collect_btn = st.button(
                "⚡ 선택 페이지 현장점검 결과 전수 수집 시작",
                type="primary",
                disabled=not st.session_state.logged_in,
                use_container_width=True,
            )

        if start_collect_btn:
            driver = st.session_state.driver
            collected = []
            progress_bar = st.progress(0)
            status_box = st.empty()
            log_expander = st.expander("📜 실시간 수집 로그 보기", expanded=True)
            log_area = log_expander.empty()
            logs = []

            total_target_pages = max(1, int(end_page) - int(start_page) + 1)

            try:
                if int(start_page) > 1:
                    driver.execute_script(f"goPage('{int(start_page)}');")
                    time.sleep(2.5)

                for p_idx, curr_page in enumerate(range(int(start_page), int(end_page) + 1), start=1):
                    status_box.info(f"🔄 현재 **{curr_page} / {int(end_page)} 페이지** 수집 중...")
                    time.sleep(1.5)

                    raw_rows = driver.find_elements(
                        By.XPATH,
                        "//table[contains(@class, 'tbl') or contains(@class, 'list') or contains(@class, 'board')]//tbody//tr[td]",
                    )
                    if not raw_rows:
                        raw_rows = driver.find_elements(By.XPATH, "//table//tbody//tr[td]")

                    filtered_indices = []
                    for idx, r in enumerate(raw_rows):
                        text = r.text.strip()
                        if any(k in text for k in ["국토안전", "현장대리인", "현장점검 기관", "담당 현장점검"]):
                            continue
                        cols = r.find_elements(By.TAG_NAME, "td")
                        if len(cols) >= 5 and "조회된" not in cols[0].text and "없습니다" not in cols[0].text:
                            filtered_indices.append(idx)

                    row_count = len(filtered_indices)
                    if row_count == 0:
                        logs.append(f"[페이지 {curr_page}] 더 이상 수집할 목록이 없습니다.")
                        log_area.code("\n".join(logs[-20:]))
                        break

                    for order_idx, original_idx in enumerate(filtered_indices):
                        try:
                            current_rows = driver.find_elements(By.XPATH, "//table//tbody//tr[td]")
                            if original_idx >= len(current_rows):
                                break
                            target_row = current_rows[original_idx]
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
                                continue

                            link_element = a_tags[0]
                            main_window = driver.current_window_handle
                            old_windows = driver.window_handles

                            driver.execute_script("arguments[0].scrollIntoView(true);", link_element)
                            time.sleep(0.2)
                            driver.execute_script("arguments[0].click();", link_element)
                            time.sleep(1.5)

                            new_windows = driver.window_handles
                            if len(new_windows) > len(old_windows):
                                new_tab = [w for w in new_windows if w not in old_windows][0]
                                driver.switch_to.window(new_tab)
                                parsed_detail = parse_detail_page(driver)
                                driver.close()
                                driver.switch_to.window(main_window)
                            else:
                                parsed_detail = parse_detail_page(driver)
                                driver.back()
                                time.sleep(1.5)

                            collected.append({"list_info": list_info, "parsed": parsed_detail})
                            st.session_state.collected_sites = collected

                            insp_cnt = len(parsed_detail.get("inspections", []))
                            logs.append(
                                f"[페이지 {curr_page} - {order_idx + 1}/{row_count}] '{list_info['공사명']}' 수집 완료 (점검결과 {insp_cnt}건)"
                            )
                            log_area.code("\n".join(logs[-20:]))

                            done_ratio = ((p_idx - 1) + (order_idx + 1) / max(1, row_count)) / total_target_pages
                            progress_bar.progress(min(1.0, done_ratio))

                        except Exception as row_err:
                            logs.append(f"[오류] {order_idx + 1}번째 현장 스킵: {row_err}")
                            log_area.code("\n".join(logs[-20:]))
                            try:
                                driver.switch_to.window(driver.window_handles[0])
                            except Exception:
                                pass

                    if curr_page < int(end_page):
                        next_page_num = curr_page + 1
                        try:
                            driver.execute_script(f"goPage('{next_page_num}');")
                            time.sleep(2.2)
                        except Exception:
                            page_btn = driver.find_elements(
                                By.XPATH, f"//a[contains(@href, \"goPage('{next_page_num}')\")]"
                            )
                            if page_btn:
                                driver.execute_script("arguments[0].click();", page_btn[0])
                                time.sleep(2.2)
                            else:
                                break

                progress_bar.progress(1.0)
                status_box.success(f"🎉 총 **{len(collected)}개** 현장 데이터 수집이 완료되었습니다! 아래 버튼으로 엑셀을 다운로드하세요.")
            except Exception as e:
                status_box.warning(
                    f"⚠️ 수집 중 중단됨 ({e}). 현재까지 수집된 **{len(st.session_state.collected_sites)}건**의 데이터를 아래에서 바로 다운로드할 수 있습니다."
                )

        if st.session_state.collected_sites:
            excel_bytes, total_cols = generate_excel_bytes(st.session_state.collected_sites)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label=f"📥 수집된 CSI 현장점검 결과 엑셀 다운로드 (총 {len(st.session_state.collected_sites)}개 현장 / {total_cols}개 열)",
                data=excel_bytes,
                file_name=f"CSI_현장점검결과_전체수집_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

with tab_html:
    st.markdown("#### 📂 저장해 둔 현장점검 결과 HTML 파일들을 한 번에 엑셀로 변환")
    st.caption(
        "CSI 상세 페이지에서 `Ctrl + S`로 저장한 `.html` 파일을 1개 또는 여러 개 동시에 끌어다 놓으면, 각 현장을 한 행씩 배치하고 모든 항목을 개별 셀에 넣은 엑셀 파일을 즉시 생성합니다."
    )

    uploaded_files = st.file_uploader(
        "현장점검 결과 HTML 파일 선택 (다중 선택 가능)",
        type=["html", "htm"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        html_sites = []
        for uf in uploaded_files:
            raw_bytes = uf.read()
            try:
                html_str = raw_bytes.decode("utf-8")
            except UnicodeDecodeError:
                html_str = raw_bytes.decode("cp949", errors="replace")
            parsed_item = parse_csi_detail_html(html_str)
            html_sites.append({"list_info": {}, "parsed": parsed_item})

        excel_bytes_html, total_cols_html = generate_excel_bytes(html_sites)
        ts_html = datetime.now().strftime("%Y%m%d_%H%M%S")

        st.success(
            f"✅ 총 **{len(html_sites)}개** 현장 HTML 파일 분석 완료! (추출된 개별 셀 열 개수: **{total_cols_html}개**)"
        )
        st.download_button(
            label=f"📥 변환된 엑셀 파일(.xlsx) 즉시 다운로드 ({len(html_sites)}건)",
            data=excel_bytes_html,
            file_name=f"CSI_현장점검결과_HTML변환_{ts_html}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
