# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import streamlit as st

from flyer_data import base_flyer_data, normalize_flyer_data
from flyer_ui import extract_quote_payload, render_flyer_editor

st.set_page_config(page_title="건강검진 안내문 제작기", layout="wide")
st.title("건강검진 안내문 제작기")
st.caption("견적서 HTML을 업로드하면 플랜명·금액·A/B/C 구성값을 자동 불러오거나, 안내문만 단독으로 제작할 수 있습니다.")

mode = st.radio("시작 방식", ["안내문 새로 만들기", "견적서 HTML 불러오기", "안내문 JSON 불러오기"], horizontal=True)
data = base_flyer_data()

if mode == "견적서 HTML 불러오기":
    f = st.file_uploader("proposal_core에서 생성한 견적서 HTML", type=["html", "htm"])
    if f:
        raw = f.getvalue().decode("utf-8", errors="replace")
        loaded = extract_quote_payload(raw)
        if loaded:
            data = loaded
            st.success("견적서의 플랜명·금액·검진 구성을 자동으로 불러왔습니다.")
        else:
            st.error("이 HTML에는 안내문 연동 데이터가 없습니다. 새 버전의 견적서 생성기에서 다시 생성한 HTML을 사용해주세요.")
elif mode == "안내문 JSON 불러오기":
    f = st.file_uploader("flyer_data.json", type=["json"])
    if f:
        try:
            data = normalize_flyer_data(json.loads(f.getvalue().decode("utf-8-sig")))
        except Exception as e:
            st.error(f"JSON 읽기 실패: {e}")

render_flyer_editor(data, prefix="standalone")
