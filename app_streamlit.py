# app_streamlit.py
import json
from pathlib import Path

import streamlit as st
from streamlit_local_storage import LocalStorage

from proposal_core import (
    load_price_options,
    parse_data,
    create_summary_table,
    render_html,
    generate_excel_bytes,
)

EXCEL_FILENAME = "2025 건강검진 견적서_표준.xlsx"
STATE_KEY = "proposal.state.v1"


def _safe_json_loads(v):
    if v is None:
        return None
    if isinstance(v, (dict, list)):
        return v
    try:
        return json.loads(v)
    except Exception:
        return None


def load_state(localS: LocalStorage):
    if st.session_state.get("_state_loaded"):
        return

    raw = localS.getItem(STATE_KEY)
    state = _safe_json_loads(raw)
    if not isinstance(state, dict):
        return

    for k in ("company", "mgr_name", "mgr_phone", "mgr_email"):
        if k in state and not st.session_state.get(k):
            st.session_state[k] = state[k]

    saved_prices = state.get("prices", {})
    if isinstance(saved_prices, dict):
        st.session_state["prices"] = saved_prices

    st.session_state["_state_loaded"] = True


def save_state(localS: LocalStorage):
    prices = st.session_state.get("prices", {})
    state = {
        "company": st.session_state.get("company", ""),
        "mgr_name": st.session_state.get("mgr_name", ""),
        "mgr_phone": st.session_state.get("mgr_phone", ""),
        "mgr_email": st.session_state.get("mgr_email", ""),
        "prices": prices,
    }
    localS.setItem(STATE_KEY, json.dumps(state, ensure_ascii=False))


def require_password():
    app_pw = st.secrets.get("APP_PASSWORD")
    if not app_pw:
        st.error("Streamlit Secrets에 APP_PASSWORD를 설정해야 합니다.")
        st.stop()

    if st.session_state.get("authed"):
        return

    pw = st.text_input("비밀번호", type="password")
    if pw and pw == app_pw:
        st.session_state["authed"] = True
        st.rerun()
    elif pw:
        st.error("비밀번호가 올바르지 않습니다.")
        st.stop()
    else:
        st.stop()


def main():
    st.set_page_config(page_title="건강검진 제안서 생성기", layout="wide")
    require_password()

    localS = LocalStorage()
    load_state(localS)

    st.title("건강검진 제안서 생성기")

    excel_path = Path(__file__).with_name(EXCEL_FILENAME)
    if not excel_path.exists():
        st.error(f"엑셀 파일을 찾지 못했습니다: {EXCEL_FILENAME}")
        st.stop()

    try:
        header_row, options = load_price_options(str(excel_path))
    except Exception as e:
        st.error(str(e))
        st.stop()

    with st.sidebar:
        st.header("제안 정보")
        st.text_input("회사명", key="company")
        st.text_input("담당자명", key="mgr_name")
        st.text_input("담당자 연락처", key="mgr_phone")
        st.text_input("담당자 이메일", key="mgr_email")

        st.divider()
        st.header("금액 선택")

    if "prices" not in st.session_state or not isinstance(st.session_state["prices"], dict):
        st.session_state["prices"] = {}

    def ensure_price_state(price_txt: str, defaults: dict):
        if price_txt not in st.session_state["prices"]:
            base_a = defaults.get("a", 0)
            base_b = defaults.get("b", 0)
            base_c = defaults.get("c", 0)

            plans_default = [
                {
                    "name": price_txt,
                    "a_rule": f"선택 {base_a}" if base_a > 0 else "-",
                    "b_rule": f"선택 {base_b}" if base_b > 0 else "-",
                    "c_rule": f"선택 {base_c}" if base_c > 0 else "-",
                },
                {
                    "name": f"{price_txt} (B형)",
                    "a_rule": f"선택 {base_a - 2}" if base_a - 2 > 0 else "-",
                    "b_rule": f"선택 {base_b + 1}" if base_b + 1 > 0 else "-",
                    "c_rule": "-",
                },
                {
                    "name": f"{price_txt} (C형)",
                    "a_rule": f"선택 {base_a - 4}" if base_a - 4 > 0 else "-",
                    "b_rule": "-",
                    "c_rule": f"선택 {base_c + 1}" if base_c + 1 > 0 else "-",
                },
            ]

            st.session_state["prices"][price_txt] = {
                "selected": False,
                "plan_count": 1,
                "plans": plans_default,
            }

    with st.sidebar:
        for opt in options:
            price_txt = opt["price_txt"]
            ensure_price_state(price_txt, opt["defaults"])

            key_sel = f"sel::{price_txt}"
            if key_sel not in st.session_state:
                st.session_state[key_sel] = bool(st.session_state["prices"][price_txt].get("selected", False))

            selected = st.checkbox(price_txt, key=key_sel)
            st.session_state["prices"][price_txt]["selected"] = bool(selected)

    st.subheader("플랜 구성 (유동적 그룹 선택 시스템)")
    selected_options = [opt for opt in options if st.session_state["prices"].get(opt["price_txt"], {}).get("selected")]
    if not selected_options:
        st.info("왼쪽에서 금액을 선택하세요.")
        save_state(localS)
        return

    plans: list[dict] = []
    for opt in selected_options:
        price_txt = opt["price_txt"]
        col_idx = opt["col_idx"]
        price_state = st.session_state["prices"][price_txt]

        with st.expander(price_txt, expanded=True):
            plan_count = st.number_input(
                "이 금액의 플랜 개수",
                min_value=1,
                max_value=3,
                value=int(price_state.get("plan_count", 1)),
                step=1,
                key=f"cnt::{price_txt}",
            )
            price_state["plan_count"] = int(plan_count)

            if not isinstance(price_state.get("plans"), list) or len(price_state["plans"]) < 3:
                ensure_price_state(price_txt, opt["defaults"])
                price_state = st.session_state["prices"][price_txt]

            for i in range(int(plan_count)):
                p = price_state["plans"][i]
                st.markdown(f"**플랜 {i+1}**")
                c1, c2, c3, c4 = st.columns([2, 1, 1, 1])

                with c1:
                    p["name"] = st.text_input("플랜명", value=p.get("name", ""), key=f"name::{price_txt}::{i}")
                with c2:
                    p["a_rule"] = st.text_input("A 규칙", value=p.get("a_rule", "-"), key=f"a::{price_txt}::{i}")
                with c3:
                    p["b_rule"] = st.text_input("B 규칙", value=p.get("b_rule", "-"), key=f"b::{price_txt}::{i}")
                with c4:
                    p["c_rule"] = st.text_input("C 규칙", value=p.get("c_rule", "-"), key=f"c::{price_txt}::{i}")

                if p["name"].strip():
                    plans.append(
                        {
                            "name": p["name"].strip(),
                            "price_txt": price_txt,
                            "col_idx": col_idx,
                            "a_rule": p.get("a_rule", "-").strip(),
                            "b_rule": p.get("b_rule", "-").strip(),
                            "c_rule": p.get("c_rule", "-").strip(),
                        }
                    )

    if not plans:
        st.warning("플랜명이 비어 있어 생성할 플랜이 없습니다. 플랜명을 입력하세요.")
        save_state(localS)
        return

    data, summary_info = parse_data(str(excel_path), header_row, plans)
    summary_table = create_summary_table(plans)

    st.divider()
    st.subheader("미리보기")
    html = render_html(
        plans=plans,
        data=data,
        summary=summary_info,
        company=st.session_state.get("company", ""),
        mgr_name=st.session_state.get("mgr_name", ""),
        mgr_phone=st.session_state.get("mgr_phone", ""),
        mgr_email=st.session_state.get("mgr_email", ""),
    )
    st.components.v1.html(html, height=900, scrolling=True)

    st.divider()
    st.subheader("다운로드")

    xlsx_bytes = generate_excel_bytes(
        plans=plans,
        data=data,
        summary=summary_table,
        company=st.session_state.get("company", ""),
        mgr_name=st.session_state.get("mgr_name", ""),
        mgr_phone=st.session_state.get("mgr_phone", ""),
        mgr_email=st.session_state.get("mgr_email", ""),
    )

    st.download_button(
        "엑셀(.xlsx) 다운로드",
        data=xlsx_bytes,
        file_name=f"2026_{st.session_state.get('company','기업')}_제안서.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    save_state(localS)


if __name__ == "__main__":
    main()
