# app_streamlit.py
import json
import streamlit as st
from pathlib import Path

# (주의) proposal_core 모듈이 같은 폴더에 있어야 합니다.
from proposal_core import load_price_options, parse_data_from_excel, render_html_string, generate_excel_bytes
from flyer_data import (
    base_flyer_data,
    build_flyer_data_from_quote,
    normalize_flyer_data,
)
from flyer_ui import (
    embed_quote_payload,
    extract_quote_payload,
    embed_flyer_in_excel,
    extract_flyer_from_excel,
    render_flyer_editor,
)

EXCEL_FILENAME = "2026 건강검진 견적서_표준_수정.xlsx"

# 1. 페이지 설정 (가장 먼저 실행되어야 함)
st.set_page_config(page_title="2026 기업건강검진 견적서 생성기", layout="wide")


# ==========================================
# 비밀번호 확인 함수
# ==========================================
def check_password():
    """비밀번호가 맞으면 True, 아니면 False를 반환하고 입력창을 띄움"""

    def password_entered():
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input(
            "비밀번호를 입력하세요",
            type="password",
            on_change=password_entered,
            key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        st.text_input(
            "비밀번호를 입력하세요",
            type="password",
            on_change=password_entered,
            key="password"
        )
        st.error("😕 비밀번호가 틀렸습니다. 다시 입력해주세요.")
        return False
    else:
        return True


@st.cache_data
def load_excel_options():
    excel_path = Path(EXCEL_FILENAME)
    if not excel_path.exists():
        return None, None
    return load_price_options(str(excel_path))


def _clear_widget_prefix(prefix: str):
    """새 견적 생성 시 이전 안내문 편집 위젯 값이 남지 않도록 초기화."""
    for key in list(st.session_state.keys()):
        if key.startswith(prefix + "_"):
            del st.session_state[key]


def _render_standalone_flyer():
    st.title("📰 건강검진 안내문 제작기")
    st.caption("안내문만 직접 만들거나, 이 견적서 페이지에서 생성한 HTML/XLSX를 업로드해 플랜명·금액을 자동 복원할 수 있습니다.")

    start_mode = st.radio(
        "안내문 시작 방식",
        ["새 안내문", "견적서 HTML 불러오기", "견적서 엑셀 불러오기", "안내문 JSON 불러오기"],
        horizontal=True,
        key="standalone_mode",
    )
    data = base_flyer_data()

    if start_mode == "견적서 HTML 불러오기":
        uploaded = st.file_uploader(
            "이 견적서 페이지에서 생성한 HTML 파일",
            type=["html", "htm"],
            key="standalone_html",
        )
        if uploaded:
            raw = uploaded.getvalue().decode("utf-8", errors="replace")
            loaded = extract_quote_payload(raw)
            if loaded:
                data = loaded
                st.success("견적서의 플랜명·금액·A/B/C 구성을 자동으로 불러왔습니다.")
            else:
                st.error("안내문 연동 데이터가 없는 예전 HTML입니다. 현재 견적서 생성기에서 새로 생성한 HTML을 사용해주세요.")

    elif start_mode == "견적서 엑셀 불러오기":
        uploaded = st.file_uploader(
            "이 견적서 페이지에서 생성한 XLSX 파일",
            type=["xlsx"],
            key="standalone_xlsx",
        )
        if uploaded:
            loaded = extract_flyer_from_excel(uploaded.getvalue())
            if loaded:
                data = loaded
                st.success("견적서 엑셀의 플랜명·금액·A/B/C 구성을 자동으로 불러왔습니다.")
            else:
                st.error("안내문 연동 데이터가 없는 예전 엑셀입니다. 현재 견적서 생성기에서 새로 생성한 엑셀을 사용해주세요.")

    elif start_mode == "안내문 JSON 불러오기":
        uploaded = st.file_uploader("flyer_data.json", type=["json"], key="standalone_json")
        if uploaded:
            try:
                data = normalize_flyer_data(json.loads(uploaded.getvalue().decode("utf-8-sig")))
                st.success("안내문 데이터를 불러왔습니다.")
            except Exception as e:
                st.error(f"JSON 읽기 실패: {e}")

    render_flyer_editor(data, prefix="standalone")


def main():
    # 병원소개서 링크는 기존 기능 유지
    st.markdown("""
        <a href="https://26nkrproposal.streamlit.app/" target="_blank" style="text-decoration: none;">
            <button style="
                background-color: #8A2BE2;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 8px;
                font-size: 16px;
                font-weight: bold;
                cursor: pointer;
                margin-bottom: 10px;">
                병원소개서 생성 링크 버튼
            </button>
        </a>
    """, unsafe_allow_html=True)

    mode = st.radio(
        "작업 모드",
        ["견적서 생성 + 안내문 자동 생성", "안내문만 제작"],
        horizontal=True,
        key="work_mode",
    )

    if mode == "안내문만 제작":
        _render_standalone_flyer()
        return

    st.title("🏥 2026 기업 건강검진 견적서 생성기")
    st.caption("견적서에서 설정한 플랜명·금액·A/B/C 선택 규칙을 그대로 사용해 건강검진 안내문을 자동 생성합니다.")

    # 1. 엑셀 로드
    header_row, options = load_excel_options()
    if not header_row:
        st.error(f"'{EXCEL_FILENAME}' 파일을 찾을 수 없거나 헤더를 읽을 수 없습니다.")
        st.stop()

    # 2. 사이드바: 입력 및 선택
    with st.sidebar:
        st.header("1. 기본 정보 입력")
        company = st.text_input("기업명 (고객사)", placeholder="예: (주)테슬라")
        mgr_name = st.text_input("담당자명", value="담당자")
        mgr_phone = st.text_input("연락처", placeholder="010-0000-0000")
        mgr_email = st.text_input("이메일")

        st.divider()
        st.header("2. 금액대 선택")
        selected_prices = []
        if options:
            for opt in options:
                if st.checkbox(f"{opt['price_txt']}", key=f"chk_{opt['price_txt']}"):
                    selected_prices.append(opt)

    # 3. 메인 영역: 플랜 상세 설정
    if not selected_prices:
        st.info("👈 왼쪽 사이드바에서 제안할 금액대를 선택해주세요.")
        return

    st.subheader("3. 세부 플랜 설정")
    st.info("여기에서 입력한 플랜명과 선택한 금액이 안내문의 기존 '건강형/소망형/믿음형/행복형/사랑형' 자리를 자동으로 대체합니다.")

    final_plans = []

    for opt in selected_prices:
        price_txt = opt['price_txt']
        defaults = opt['defaults']
        base_a, base_b, base_c = defaults['a'], defaults['b'], defaults['c']

        with st.expander(f"{price_txt} 플랜 설정", expanded=True):
            cols = st.columns([1, 4])
            with cols[0]:
                cnt = st.number_input(f"{price_txt} 개수", min_value=1, max_value=3, value=1, key=f"cnt_{price_txt}")

            for i in range(int(cnt)):
                st.markdown(f"**Option {i+1}**")
                c1, c2, c3, c4 = st.columns([2, 1, 1, 1])

                def_name = f"{price_txt}"
                def_a, def_b, def_c = base_a, base_b, base_c

                if i == 1:
                    def_name += " (B형)"
                    def_a = max(0, base_a - 2)
                    def_b = base_b + 1
                elif i == 2:
                    def_name += " (C형)"
                    def_a = max(0, base_a - 4)
                    def_c = base_c + 1

                str_a = f"선택 {def_a}" if def_a > 0 else "-"
                str_b = f"선택 {def_b}" if def_b > 0 else "-"
                str_c = f"선택 {def_c}" if def_c > 0 else "-"

                with c1:
                    p_name = st.text_input("플랜명", value=def_name, key=f"name_{price_txt}_{i}")
                with c2:
                    p_a = st.text_input("A선택", value=str_a, key=f"a_{price_txt}_{i}")
                with c3:
                    p_b = st.text_input("B선택", value=str_b, key=f"b_{price_txt}_{i}")
                with c4:
                    p_c = st.text_input("C선택", value=str_c, key=f"c_{price_txt}_{i}")

                final_plans.append({
                    "name": p_name,
                    "col_idx": opt['col_idx'],
                    "sort_key": opt.get('sort_key'),
                    "a_rule": p_a,
                    "b_rule": p_b,
                    "c_rule": p_c,
                    "price_txt": opt['price_txt']
                })

    st.divider()

    # 4. 견적서 생성
    if st.button("견적서 생성하기 (HTML 미리보기 & 엑셀 & 안내문)", type="primary"):
        with st.spinner("데이터 처리 중..."):
            try:
                info = {"company": company, "name": mgr_name, "phone": mgr_phone, "email": mgr_email}
                data, summary = parse_data_from_excel(str(Path(EXCEL_FILENAME).resolve()), header_row, final_plans)

                # 견적서와 안내문이 같은 final_plans / data를 사용합니다.
                flyer_data = build_flyer_data_from_quote(final_plans, data, summary, info)
                html_str = render_html_string(final_plans, data, summary, info)
                html_str = embed_quote_payload(html_str, flyer_data)
                excel_bytes = generate_excel_bytes(final_plans, data, summary, info)
                excel_bytes = embed_flyer_in_excel(excel_bytes, flyer_data)

                _clear_widget_prefix("quote_flyer")
                st.session_state["quote_result"] = {
                    "html": html_str,
                    "excel": excel_bytes,
                    "company": company,
                    "flyer_data": flyer_data,
                }
                st.success("견적서와 안내문 연동 데이터 생성이 완료되었습니다.")
            except Exception as e:
                st.error(f"생성 중 오류가 발생했습니다: {e}")

    # 버튼 이후 rerun에도 결과 유지
    result = st.session_state.get("quote_result")
    if result:
        tab1, tab2, tab3 = st.tabs(["📄 HTML 미리보기", "💾 견적서 다운로드", "📰 안내문 자동 생성"])

        with tab1:
            st.components.v1.html(result["html"], height=1000, scrolling=True)

        with tab2:
            st.success("견적서 생성 완료")
            col1, col2 = st.columns(2)
            filename_xls = f"2026_{result['company']}_건강검진_견적서.xlsx"
            filename_html = f"2026_{result['company']}_건강검진_견적서.html"
            with col1:
                st.download_button(
                    "📥 엑셀 파일 다운로드 (.xlsx)",
                    result["excel"],
                    filename_xls,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with col2:
                st.download_button(
                    "📥 HTML 파일 다운로드 (.html)",
                    result["html"],
                    filename_html,
                    "text/html",
                )
            st.caption("다운로드한 HTML과 XLSX에는 안내문용 플랜명·금액·A/B/C 데이터가 함께 저장됩니다. '안내문만 제작' 모드에서 다시 업로드할 수 있습니다.")

        with tab3:
            render_flyer_editor(result["flyer_data"], prefix="quote_flyer")


if __name__ == "__main__":
    if check_password():
        main()
