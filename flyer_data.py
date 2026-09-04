# -*- coding: utf-8 -*-
from __future__ import annotations

from copy import deepcopy
from typing import Any, Dict, List
import json
import re

OFFICIAL_2026: Dict[str, Any] = {
    "schema_version": 2,
    "source": "proposal_core",
    "organization": "뉴고려병원",
    "company": "",
    "title": "2026 건강검진 안내문",
    "target": "임직원 및 가족",
    "period": "",
    "application_period": "",
    "phone": "1833-9988",
    "qr_url": "",
    "event_title": "EVENT",
    "event_lines": ["종합검진 진행 시 A그룹 추가 혜택"],
    "common_items": (
        "간기능 | 간염 | 순환기계 | 당뇨 | 췌장기능 | 철결핍성 | 빈혈 | 혈액질환 | 전해질 | "
        "신장기능 | 골격계질환 | 감염성 | 갑상선기능 | 부갑상선기능 | 종양표지자 | 소변 등 "
        "800여종 혈액(소변) 검사 | 심전도 | 신장 | 체중 | 혈압 | 시력 | 청력 | 체성분 | "
        "건강유형분석 | 폐기능 | 안저 | 안압 | 혈액점도검사 | 유전자20종 | 흉부X-ray | "
        "복부초음파 | 위수면내시경 | (여)자궁경부세포진 | (여)유방촬영 - #30세이상 권장#"
    ),
    "groups": {"A": [], "B": [], "C": []},
    "packages": [],
    "notes": [
        "공단검진 대상자는 종합검진 진행 시 공단청구 금액을 차감해드립니다.",
        "연속검진 중복 할인 적용 불가합니다.",
    ],
    "theme": "기업형",
}


def base_flyer_data() -> Dict[str, Any]:
    return deepcopy(OFFICIAL_2026)


def _selected_count(rule: str) -> int:
    m = re.search(r"(\d+)", str(rule or ""))
    return int(m.group(1)) if m else 0


def _plan_detail(plan: Dict[str, Any]) -> str:
    chunks = ["공통항목"]
    for group, key in (("A그룹", "a_rule"), ("B그룹", "b_rule"), ("C그룹", "c_rule")):
        count = _selected_count(plan.get(key, ""))
        if count:
            chunks.append(f"{group} {count}가지")
    return " + ".join(chunks)


def _group_items(parsed_data: Dict[str, Any], group: str) -> List[str]:
    out: List[str] = []
    rows = parsed_data.get(group, []) if isinstance(parsed_data, dict) else []
    for i, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            continue
        name = str(row.get("name", "")).strip()
        if not name:
            continue
        prefix = f"[{i:02d}]" if group == "A" else f"[{chr(44032 + (i-1)*588)}]" if i <= 14 else f"[{i}]"
        if re.match(r"^\[[^\]]+\]", name):
            out.append(name)
        else:
            out.append(f"{prefix} {name}")
    return out


def build_flyer_data_from_quote(
    plans: List[Dict[str, Any]],
    parsed_data: Dict[str, Any],
    summary: List[Dict[str, Any]],
    info: Dict[str, Any],
) -> Dict[str, Any]:
    """견적서 생성 결과를 안내문용 구조로 변환합니다."""
    d = base_flyer_data()
    company = str(info.get("company", "")).strip()
    d["company"] = company
    d["title"] = f"{company} 건강검진 안내문" if company else "2026 건강검진 안내문"
    d["phone"] = "1833-9988"

    d["groups"]["A"] = _group_items(parsed_data, "A")
    d["groups"]["B"] = _group_items(parsed_data, "B")
    d["groups"]["C"] = _group_items(parsed_data, "C")

    packages: List[Dict[str, str]] = []
    for plan in plans:
        price = str(plan.get("price_txt", "")).strip()
        packages.append({
            "name": str(plan.get("name", price or "검진형")).strip(),
            "detail": _plan_detail(plan),
            "price": price,
            "male_price": price,
            "female_price": price,
        })
    if packages:
        d["packages"] = packages

    d["event_title"] = "기업 건강검진 혜택"
    d["event_lines"] = [
        "선택한 견적 플랜을 기준으로 안내문이 자동 생성되었습니다.",
        "배경·기간·대상·이벤트 문구는 안내문 편집에서 자유롭게 수정할 수 있습니다.",
    ]
    return d


def normalize_flyer_data(data: Dict[str, Any] | None) -> Dict[str, Any]:
    base = base_flyer_data()
    if not isinstance(data, dict):
        return base
    for key in [
        "organization", "company", "title", "target", "period", "application_period", "phone",
        "qr_url", "event_title", "common_items", "theme"
    ]:
        if key in data and data[key] is not None:
            base[key] = str(data[key])
    for key in ["event_lines", "notes"]:
        if isinstance(data.get(key), list):
            base[key] = [str(x) for x in data[key] if str(x).strip()]
    if isinstance(data.get("groups"), dict):
        for g in ("A", "B", "C"):
            if isinstance(data["groups"].get(g), list):
                base["groups"][g] = [str(x) for x in data["groups"][g] if str(x).strip()]
    if isinstance(data.get("packages"), list):
        rows = []
        for row in data["packages"]:
            if not isinstance(row, dict):
                continue
            price = str(row.get("price", "")).strip()
            rows.append({
                "name": str(row.get("name", "")).strip(),
                "detail": str(row.get("detail", "")).strip(),
                "price": price,
                "male_price": str(row.get("male_price", price)).strip(),
                "female_price": str(row.get("female_price", price)).strip(),
            })
        if rows:
            base["packages"] = rows
    return base


def dumps_flyer_json(data: Dict[str, Any]) -> bytes:
    return json.dumps(normalize_flyer_data(data), ensure_ascii=False, indent=2).encode("utf-8")
