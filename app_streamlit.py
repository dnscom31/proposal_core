# app_streamlit.py
import streamlit as st
import os
from pathlib import Path
# (주의) proposal_core.py 파일이 같은 폴더에 있어야 합니다.
from proposal_core import load_price_options, parse_data_from_excel, render_html_string, generate_excel_bytes

# 기본으로 사용할 파일명 (업로드 안 했을 때 사용)
DEFAULT_EXCEL_FILENAME = "2025 건강검진 견적서_표준.xlsx"

# 1. 페이지 설정 (가장 먼저 실행)
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
        st.text_input("비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        st.error("😕 비밀번호가 틀렸습니다. 다시 입력해주세요.")
        return False
    else:
        return True

# ==========================================
# 데이터 로드 함수 (캐시 적용)
# ==========================================
@st.cache_data
def load_excel_options(file_path_str):
    """경로를 인자로 받아 데이터를 로드 (파일이 바뀌면 캐시 갱신)"""
    excel_path = Path(file_path_str)
    if not excel_path.exists():
        return None, None
    return load_price_options(str(excel_path))

# ==========================================
# 메인 함수
# ==========================================
def main():
    # [병원소개서 링크 버튼]
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

    st.title("🏥 2026 기업 건강검진 견적서 생성기")

    # -----------------------------------------------------------
    # [기능 추가] 사이드바: 파일 업로드 및 설정
    # -----------------------------------------------------------
    with st.sidebar:
        st.header("📂 엑셀 파일 설정")
        
        # 1. 파일 업로더
        uploaded_file = st.file_uploader("수정된 견적서 엑셀 파일 업로드", type=['xlsx'])
        
        # 파일 경로 결정 로직
        if uploaded_file is not None:
            # 업로드된 파일을 임시 파일로 저장
            target_file_path = "temp_uploaded_excel.xlsx"
            with open(target_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success("✅ 업로드된 파일이 적용되었습니다.")
        else:
            # 업로드가 없으면 기본 파일 사용
            target_file_path = DEFAULT_EXCEL_FILENAME
            st.info(f"기본 파일 사용 중: {DEFAULT_EXCEL_FILENAME}")

        st.divider()

        st.header("1. 기본 정보 입력")
        company = st.text_input("기업명 (고객사)", placeholder="예: (주)테슬라")
        mgr_name = st.text_input("담당자명", value="담당자")
        mgr_phone = st.text_input("연락처", placeholder="010-0000-0000")
        mgr_email = st.text_input("이메일")
        
        # -------------------------------------------------------
        # 엑셀 로드 (위에서 결정된 target_file_path 사용)
        # -------------------------------------------------------
        header_row, options = load_excel_options(target_file_path)
        
        if not header_row:
            st.error(f"❌ 파일을 읽을 수 없습니다. 경로: {target_file_path}")
            # 업로드된 파일이 없고 기본 파일도 없는 경우 중단
            if uploaded_file is None and not Path(DEFAULT_EXCEL_FILENAME).exists():
                st.warning("기본 엑셀 파일이 없습니다. 파일을 업로드해주세요.")
                st.stop()
        
        st.divider()
        st.header("2. 금액대 선택")
        selected_prices = []
        
        # 로드된 옵션으로 체크박스 생성
        if options:
            for opt in options:
                if st.checkbox(f"{opt['price_txt']}", key=f"chk_{opt['price_txt']}"):
                    selected_prices.append(opt)
        else:
            st.warning("엑셀에서 금액 정보를 찾을 수 없습니다.")

    # 3. 메인 영역: 플랜 상세 설정
    if not selected_prices:
        st.info("👈 왼쪽 사이드바에서 제안할 금액대를 선택해주세요.")
        return

    st.subheader("3. 세부 플랜 설정")
    
    final_plans = []
    
    # 선택된 금액대별 설정 카드
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
                
                # 기본값 계산
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
                    "a_rule": p_a, "b_rule": p_b, "c_rule": p_c,
                    "price_txt": opt['price_txt']
                })

    st.divider()

    # 4. 생성 및 다운로드
    if st.button("견적서 생성하기 (HTML 미리보기 & 엑셀 생성)", type="primary"):
        with st.spinner("데이터 처리 중..."):
            info = {"company": company, "name": mgr_name, "phone": mgr_phone, "email": mgr_email}
            
            # [중요] 결정된 경로(target_file_path)를 사용하여 데이터 파싱
            data, summary = parse_data_from_excel(str(Path(target_file_path).resolve()), header_row, final_plans)
            
            html_str = render_html_string(final_plans, data, summary, info)
            excel_bytes = generate_excel_bytes(final_plans, data, summary, info)
            
            tab1, tab2 = st.tabs(["📄 HTML 미리보기", "💾 다운로드"])
            
            with tab1:
                st.components.v1.html(html_str, height=1000, scrolling=True)
            
            with tab2:
                st.success("생성이 완료되었습니다!")
                col1, col2 = st.columns(2)
                with col1:
                    filename_xls = f"2026_{company}_건강검진_견적서.xlsx"
                    st.download_button("📥 엑셀 파일 다운로드 (.xlsx)", excel_bytes, filename_xls, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with col2:
                    filename_html = f"2026_{company}_건강검진_견적서.html"
                    st.download_button("📥 HTML 파일 다운로드 (.html)", html_str, filename_html, "text/html")

if __name__ == "__main__":
    if check_password():
        main()
