# app_streamlit.py
import streamlit as st
from pathlib import Path
from proposal_core import load_price_options, parse_data_from_excel, render_html_string, generate_excel_bytes

EXCEL_FILENAME = "2025 건강검진 견적서_표준.xlsx"

st.set_page_config(page_title="2026 기업건강검진 견적서 생성기", layout="wide")

@st.cache_data
def load_excel_options():
    excel_path = Path(EXCEL_FILENAME)
    if not excel_path.exists():
        return None, None
    return load_price_options(str(excel_path))

def main():
    st.title("🏥 2026 기업 건강검진 제안서 생성기")

    # 1. 엑셀 로드
    header_row, options = load_excel_options()
    if not header_row:
        st.error(f"'{EXCEL_FILENAME}' 파일을 찾을 수 없거나 헤더를 읽을 수 없습니다.")
        st.stop()

    # 2. 사이드바: 입력 및 선택
    with st.sidebar:
        st.header("1. 기본 정보 입력")
        company = st.text_input("기업명 (고객사)", placeholder="예: 삼성전자")
        mgr_name = st.text_input("담당자명", value="담당자")
        mgr_phone = st.text_input("연락처", placeholder="010-0000-0000")
        mgr_email = st.text_input("이메일")
        
        st.divider()
        st.header("2. 금액대 선택")
        selected_prices = []
        for opt in options:
            if st.checkbox(f"{opt['price_txt']}", key=f"chk_{opt['price_txt']}"):
                selected_prices.append(opt)

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
            
            # N개의 플랜 입력 폼 생성
            for i in range(int(cnt)):
                st.markdown(f"**Option {i+1}**")
                c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
                
                # 기본값 계산 로직 (Tkinter와 동일)
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
                    "a_rule": p_a, "b_rule": p_b, "c_rule": p_c
                })

    st.divider()

    # 4. 생성 및 다운로드
    if st.button("제안서 생성하기 (HTML 미리보기 & 엑셀 생성)", type="primary"):
        with st.spinner("데이터 처리 중..."):
            # 데이터 파싱
            info = {"company": company, "name": mgr_name, "phone": mgr_phone, "email": mgr_email}
            data, summary = parse_data_from_excel(str(Path(EXCEL_FILENAME).resolve()), header_row, final_plans)
            
            # HTML 생성
            html_str = render_html_string(final_plans, data, summary, info)
            
            # 엑셀 생성
            excel_bytes = generate_excel_bytes(final_plans, data, summary, info)
            
            # 탭으로 보기 분리
            tab1, tab2 = st.tabs(["📄 HTML 미리보기", "💾 다운로드"])
            
            with tab1:
                st.components.v1.html(html_str, height=1000, scrolling=True)
            
            with tab2:
                st.success("생성이 완료되었습니다!")
                
                col1, col2 = st.columns(2)
                with col1:
                    filename_xls = f"2026_{company}_건강검진_제안서.xlsx"
                    st.download_button(
                        label="📥 엑셀 파일 다운로드 (.xlsx)",
                        data=excel_bytes,
                        file_name=filename_xls,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with col2:
                    filename_html = f"2026_{company}_건강검진_제안서.html"
                    st.download_button(
                        label="📥 HTML 파일 다운로드 (.html)",
                        data=html_str,
                        file_name=filename_html,
                        mime="text/html"
                    )

if __name__ == "__main__":
    main()

