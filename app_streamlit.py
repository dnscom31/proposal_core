# app_streamlit.py
# -*- coding: utf-8 -*-

import io
import streamlit as st

from proposal_core import load_price_options, parse_data, render_html, generate_excel_bytes


def require_password():
    app_pw = st.secrets.get("APP_PASSWORD", None)
    if not app_pw:
        st.error("APP_PASSWORD가 설정되어 있지 않습니다. (Streamlit Cloud > App settings > Secrets)")
        st.stop()

    if st.session_state.get("authed", False):
        return

    st.subheader("접속 비밀번호")
    typed = st.text_input("비밀번호", type="password")
    if st.button("로그인"):
        if typed == app_pw:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
            st.stop()
    else:
        st.stop()


# -----------------------------
# 앱 시작
# -----------------------------
st.set_page_config(page_title="2026 건강검진 제안서 생성기", layout="wide")
require_password()

st.title("2026 건강검진 제안서 생성기 (Streamlit)")

TEMPLATE_XLSX = "2025 건강검진 견적서_표준.xlsx"

# 옵션 로드(캐시)
@st.cache_data(show_spinner=False)
def _load_options(path: str):
    return load_price_options(path)

try:
    header_row, options = _load_options(TEMPLATE_XLSX)
except Exception as e:
    st.error(f"템플릿 엑셀을 읽을 수 없습니다: {e}")
    st.stop()

# 회사/담당자
with st.sidebar:
    st.header("기본 정보")
    company = st.text_input("회사명", value="")
    mgr_name = st.text_input("담당자 이름", value="")
    mgr_phone = st.text_input("담당자 연락처", value="")
    mgr_email = st.text_input("담당자 이메일", value="")

st.divider()

st.subheader("플랜 구성")
st.caption("원하시는 금액대(컬럼)를 선택하고, 플랜명을 지정하고 A/B/C 규칙(예: '선택 5', '-' 등)을 입력합니다.")

# 최대 4개 플랜 UI
max_plans = 4
selected_plans = []

cols = st.columns(max_plans)
for i in range(max_plans):
    with cols[i]:
        st.markdown(f"### 플랜 {i+1}")
        use = st.checkbox("사용", value=(i == 0), key=f"use_{i}")

        price_labels = [o["price_txt"] for o in options]
        price = st.selectbox("금액대", price_labels, index=min(i, len(price_labels) - 1), key=f"price_{i}")
        plan_name = st.text_input("플랜명", value=price, key=f"name_{i}")

        # 선택한 가격 옵션 찾기
        opt = next(o for o in options if o["price_txt"] == price)
        defaults = opt["defaults"]
        a_rule = st.text_input("A그룹 규칙", value=f"선택 {defaults['a']}" if defaults["a"] else "-", key=f"a_{i}")
        b_rule = st.text_input("B그룹 규칙", value=f"선택 {defaults['b']}" if defaults["b"] else "-", key=f"b_{i}")
        c_rule = st.text_input("C그룹 규칙", value=f"선택 {defaults['c']}" if defaults["c"] else "-", key=f"c_{i}")

        if use:
            selected_plans.append({
                "name": plan_name.strip() or price,
                "col_idx": opt["col_idx"],
                "a_rule": a_rule.strip(),
                "b_rule": b_rule.strip(),
                "c_rule": c_rule.strip(),
            })

if not selected_plans:
    st.warning("최소 1개 플랜을 선택하세요.")
    st.stop()

# 데이터 파싱(캐시)
@st.cache_data(show_spinner=True)
def _parse(path: str, header_row_: int, plans_):
    return parse_data(path, header_row_, plans_)

data, summary = _parse(TEMPLATE_XLSX, header_row, selected_plans)

st.divider()
st.subheader("미리보기 / 다운로드")

left, right = st.columns([1, 1])

with left:
    st.markdown("#### HTML 제안서 미리보기")
    html = render_html(
        plans=selected_plans,
        data=data,
        summary=summary,
        company=company,
        mgr_name=mgr_name,
        mgr_phone=mgr_phone,
        mgr_email=mgr_email,
    )
    st.components.v1.html(html, height=720, scrolling=True)

    st.download_button(
        "HTML 다운로드",
        data=html.encode("utf-8"),
        file_name=f"2026_{company or '기업'}_건강검진_제안서.html",
        mime="text/html",
    )

with right:
    st.markdown("#### Excel 제안서 다운로드")
    xlsx_bytes = generate_excel_bytes(
        plans=selected_plans,
        data=data,
        summary=summary,
        company=company,
        mgr_name=mgr_name,
        mgr_phone=mgr_phone,
        mgr_email=mgr_email,
    )

    st.download_button(
        "Excel 다운로드",
        data=xlsx_bytes,
        file_name=f"2026_{company or '기업'}_건강검진_제안서.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption("참고: Streamlit Community Cloud에서는 파일 경로가 저장소 루트 기준 상대경로여야 합니다.")

