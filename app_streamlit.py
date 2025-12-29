# app_streamlit.py
import streamlit as st

APP_PASSWORD = "nk2026"

def require_password():
    if "authed" not in st.session_state:
        st.session_state.authed = False

    if not st.session_state.authed:
        pw = st.text_input("비밀번호", type="password")
        if st.button("로그인"):
            if pw == APP_PASSWORD:
                st.session_state.authed = True
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
        st.stop()

require_password()

import streamlit.components.v1 as components
import re

from proposal_core import (
    load_price_options,
    parse_data,
    render_html,
    generate_excel_bytes,
)

TEMPLATE_XLSX = "2025 건강검진 견적서_표준.xlsx"


def make_default_subplan(price_txt, defaults, i, count):
    """
    원본 update_sub_plans()의 기본 생성 규칙을 웹에서 재현
    """
    base_a, base_b, base_c = defaults["a"], defaults["b"], defaults["c"]

    p_name = f"{price_txt}"
    if i == 1:
        p_name += " (B형)"
    elif i == 2:
        p_name += " (C형)"
    elif count > 1:
        p_name += "-1"

    curr_a, curr_b, curr_c = base_a, base_b, base_c
    is_valid = True

    if i == 1:
        curr_a = base_a - 2
        curr_b = base_b + 1
        if curr_a < 0:
            is_valid = False
    elif i == 2:
        curr_a = base_a - 4
        curr_c = base_c + 1
        if curr_a < 0:
            is_valid = False

    str_a = f"선택 {curr_a}" if curr_a > 0 else "-"
    str_b = f"선택 {curr_b}" if curr_b > 0 else "-"
    str_c = f"선택 {curr_c}" if curr_c > 0 else "-"
    if not is_valid:
        str_a = "-"

    return p_name, str_a, str_b, str_c


st.set_page_config(page_title="2026 기업검진 제안서 생성기", layout="wide")
st.title("2026 기업 건강검진 제안서 (웹 버전)")

# 엑셀에서 가격 옵션 로드
try:
    header_row, options = load_price_options(TEMPLATE_XLSX)
except Exception as e:
    st.error(f"엑셀 템플릿을 읽을 수 없습니다: {e}")
    st.stop()

with st.sidebar:
    st.subheader("제안 정보 입력")
    company = st.text_input("기업명(고객사)", value="")
    mgr_name = st.text_input("담당자명", value="")
    mgr_phone = st.text_input("연락처", value="")
    mgr_email = st.text_input("이메일", value="")

st.subheader("1) 금액대 선택 및 플랜 구성")

plans = []

for opt in options:
    defaults = opt["defaults"]
    info = []
    if defaults["a"] > 0:
        info.append(f"A{defaults['a']}")
    if defaults["b"] > 0:
        info.append(f"B{defaults['b']}")
    if defaults["c"] > 0:
        info.append(f"C{defaults['c']}")
    label = f"{opt['price_txt']} ({'/'.join(info)})" if info else opt["price_txt"]

    with st.expander(label, expanded=False):
        use = st.checkbox("이 금액대 사용", key=f"use_{opt['col_idx']}")
        if not use:
            continue

        count = st.number_input(
            "개 플랜 생성(1~3)",
            min_value=1,
            max_value=3,
            value=1,
            step=1,
            key=f"count_{opt['col_idx']}",
        )

        st.caption("플랜명 및 A/B/C 선택 규칙을 필요에 따라 수정할 수 있습니다.")
        for i in range(int(count)):
            default_name, default_a, default_b, default_c = make_default_subplan(opt["price_txt"], defaults, i, int(count))

            c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
            name = c1.text_input("플랜명", value=default_name, key=f"name_{opt['col_idx']}_{i}")
            a_rule = c2.text_input("A선택", value=default_a, key=f"a_{opt['col_idx']}_{i}")
            b_rule = c3.text_input("B선택", value=default_b, key=f"b_{opt['col_idx']}_{i}")
            c_rule = c4.text_input("C선택", value=default_c, key=f"c_{opt['col_idx']}_{i}")

            plans.append({
                "col_idx": opt["col_idx"],
                "original_price": opt["price_txt"],
                "name": name.strip() or opt["price_txt"],
                "a_rule": a_rule.strip(),
                "b_rule": b_rule.strip(),
                "c_rule": c_rule.strip(),
            })

if st.button("2) HTML 미리보기/다운로드 생성"):
    if not plans:
        st.warning("최소 하나의 플랜을 선택해주세요.")
        st.stop()

    data, summary = parse_data(TEMPLATE_XLSX, header_row, plans)

    html = render_html(
        plans=plans,
        data=data,
        summary=summary,
        company=company,
        mgr_name=mgr_name,
        mgr_phone=mgr_phone,
        mgr_email=mgr_email,
    )

    st.subheader("미리보기")
    components.html(html, height=900, scrolling=True)

    safe_company = re.sub(r"[^0-9A-Za-z가-힣_-]+", "_", (company.strip() or "기업"))
    html_filename = f"2026_{safe_company}_건강검진_제안서.html"
    xlsx_filename = f"2026_{safe_company}_건강검진_제안서.xlsx"

    st.download_button(
        label="HTML 파일 다운로드",
        data=html.encode("utf-8"),
        file_name=html_filename,
        mime="text/html",
    )

    excel_bytes = generate_excel_bytes(
        plans=plans,
        data=data,
        summary=summary,
        company=company,
        mgr_name=mgr_name,
        mgr_phone=mgr_phone,
        mgr_email=mgr_email,
    )

    st.download_button(
        label="엑셀(.xlsx) 다운로드",
        data=excel_bytes,
        file_name=xlsx_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

