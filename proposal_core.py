# proposal_core.py
# -*- coding: utf-8 -*-
"""
뉴고려병원 2026 기업검진 제안서 생성 코어(웹/서버용)

- 엑셀 템플릿에서 가격 옵션을 읽고(load_price_options)
- 플랜 구성(plans)을 받아 항목 데이터를 파싱(parse_data)
- HTML 제안서 생성(render_html)
- 엑셀 제안서 생성(generate_excel_bytes)
"""

from __future__ import annotations

import io
import re
from datetime import datetime
from typing import Dict, List, Tuple, Any

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break


# -----------------------------
# 1) 템플릿 스캔 / 플랜 옵션 로드
# -----------------------------
def scan_default_counts(sheet, col_idx: int, start_row: int) -> Dict[str, int]:
    """
    특정 금액 컬럼(col_idx)에서 A/B/C 그룹별 기본 선택(선택 N) 최대값을 스캔해 추정합니다.
    """
    counts = {"a": 0, "b": 0, "c": 0}
    max_scan = min(start_row + 150, sheet.max_row)
    current_cat = ""

    for r in range(start_row + 1, max_scan + 1):
        cell_group = str(sheet.cell(row=r, column=1).value).strip() if sheet.cell(row=r, column=1).value else ""
        cell_val = str(sheet.cell(row=r, column=col_idx).value).strip() if sheet.cell(row=r, column=col_idx).value else ""

        if "A그룹" in cell_group:
            current_cat = "a"
        elif "B그룹" in cell_group:
            current_cat = "b"
        elif "C그룹" in cell_group:
            current_cat = "c"

        if current_cat in ["a", "b", "c"] and "선택" in cell_val:
            nums = re.findall(r"\d+", cell_val)
            if nums:
                n = int(nums[0])
                if n > counts[current_cat]:
                    counts[current_cat] = n
    return counts


def load_price_options(excel_filename: str) -> Tuple[int, List[Dict[str, Any]]]:
    """
    엑셀에서 '만원' 헤더 행을 찾고, 각 금액대 컬럼 옵션을 반환합니다.

    Returns:
        header_row_idx: 금액 헤더 행 번호
        options: [{price_txt, col_idx, defaults}, ...]
    """
    wb = openpyxl.load_workbook(excel_filename, data_only=True)
    sheet = wb.active

    header_row_idx = None
    for row in sheet.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and "만원" in str(cell.value):
                header_row_idx = cell.row
                break
        if header_row_idx:
            break

    if not header_row_idx:
        wb.close()
        raise ValueError("금액 헤더('만원')를 찾을 수 없습니다.")

    excluded = ["10만원", "15만원"]
    row_cells = list(sheet.rows)[header_row_idx - 1]

    # 원본(데스크톱) 기준 수동 기본값 테이블
    manual_defaults = {
        25: {"a": 3, "b": 0, "c": 0}, 30: {"a": 3, "b": 0, "c": 0},
        35: {"a": 4, "b": 0, "c": 0}, 40: {"a": 5, "b": 0, "c": 0},
        45: {"a": 4, "b": 1, "c": 0}, 50: {"a": 5, "b": 1, "c": 0},
        60: {"a": 3, "b": 1, "c": 1}, 70: {"a": 5, "b": 1, "c": 1},
        80: {"a": 5, "b": 2, "c": 1}, 90: {"a": 5, "b": 3, "c": 1},
        100: {"a": 3, "b": 3, "c": 2},
    }

    options: List[Dict[str, Any]] = []
    for idx, cell in enumerate(row_cells):
        val = str(cell.value).strip() if cell.value else ""
        if "만원" in val and not any(e in val for e in excluded):
            col_idx = idx + 1
            scanned = scan_default_counts(sheet, col_idx, header_row_idx)
            try:
                price_num = int(re.sub(r"[^0-9]", "", val))
            except Exception:
                price_num = 0
            defaults = manual_defaults.get(price_num, scanned)

            options.append({
                "price_txt": val,
                "col_idx": col_idx,
                "defaults": defaults,
            })

    wb.close()
    options.sort(key=lambda x: int(re.sub(r"[^0-9]", "", x["price_txt"]) or "999"))
    return header_row_idx, options


# -----------------------------
# 2) 데이터 파싱
# -----------------------------
def parse_data(excel_filename: str, header_row: int, plans: List[Dict[str, Any]]) -> Tuple[Dict[str, List[Dict[str, Any]]], List[Dict[str, str]]]:
    """
    템플릿 엑셀에서 A/B/C/EQUIP/COMMON 항목을 읽어,
    plans(플랜 구성)에 맞는 값을 매핑합니다.
    """
    wb = openpyxl.load_workbook(excel_filename, data_only=True)
    sheet = wb.active

    parsed_data: Dict[str, List[Dict[str, Any]]] = {"A": [], "B": [], "C": [], "EQUIP": [], "COMMON_BLOOD": []}
    summary_info = [{"name": p["name"], "a": p["a_rule"], "b": p["b_rule"], "c": p["c_rule"]} for p in plans]

    # "선택 N"이 비어있는 행에서 위 값이 이어지는 경우를 위해 캐시
    fill_cache = {i: {"A": None, "B": None, "C": None} for i in range(len(plans))}
    current_main_cat = ""

    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        if not row or len(row) < 2:
            continue

        col0 = str(row[0]).strip() if row[0] else ""
        col1 = str(row[1]).strip() if row[1] else ""

        if "A그룹" in col0:
            current_main_cat = "A"
        elif "B그룹" in col0:
            current_main_cat = "B"
        elif "C그룹" in col0:
            current_main_cat = "C"
        elif "장비검사" in col0 or "소화기검사" in col0:
            current_main_cat = "EQUIP"
        elif "혈액" in col0 and "소변" in col0:
            current_main_cat = "COMMON"

        if not col1 or col1 in ["검진항목", "내용"]:
            continue

        item_name = col1
        item_desc = str(row[2]).strip() if len(row) > 2 and row[2] else ""
        sub_cat = col0 if current_main_cat == "EQUIP" and col0 else ""

        row_vals: List[str] = []
        for idx, plan in enumerate(plans):
            col_idx0 = plan["col_idx"] - 1
            val = str(row[col_idx0]).strip() if col_idx0 < len(row) and row[col_idx0] else ""

            # 캐시 적용 (A/B/C 그룹)
            if current_main_cat in ["A", "B", "C"]:
                cache = fill_cache[idx]
                if "선택" in val:
                    cache[current_main_cat] = val
                elif val == "" and cache[current_main_cat]:
                    val = cache[current_main_cat]
                elif val != "":
                    cache[current_main_cat] = None

            # 사용자 입력 a_rule/b_rule/c_rule로 override
            if current_main_cat in ["A", "B", "C"] and "선택" in val:
                custom_rule = ""
                if current_main_cat == "A":
                    custom_rule = plan["a_rule"]
                elif current_main_cat == "B":
                    custom_rule = plan["b_rule"]
                elif current_main_cat == "C":
                    custom_rule = plan["c_rule"]

                if custom_rule:
                    if custom_rule == "-":
                        val = ""
                    else:
                        val = custom_rule

            if "미선택" in val:
                val = ""

            row_vals.append(val)

        entry = {"category": sub_cat, "name": item_name, "desc": item_desc, "values": row_vals}

        if current_main_cat == "A":
            parsed_data["A"].append(entry)
        elif current_main_cat == "B":
            parsed_data["B"].append(entry)
        elif current_main_cat == "C":
            parsed_data["C"].append(entry)
        elif current_main_cat == "EQUIP":
            parsed_data["EQUIP"].append(entry)
        elif current_main_cat == "COMMON":
            parsed_data["COMMON_BLOOD"].append(entry)

    wb.close()
    return parsed_data, summary_info


# -----------------------------
# 3) HTML 생성 (propsal2026.py의 render_html을 함수화)
# -----------------------------
def render_html(
    plans: List[Dict[str, Any]],
    data: Dict[str, List[Dict[str, Any]]],
    summary: List[Dict[str, str]],
    company: str,
    mgr_name: str,
    mgr_phone: str,
    mgr_email: str,
) -> str:
    today_date = datetime.now().strftime("%Y년 %m월 %d일")
    mgr_name = mgr_name or "담당자"
    mgr_phone = mgr_phone or ""
    mgr_email = mgr_email or ""
    company = (company or "").strip()
    proposal_title = f"2026 {company} 임직원 건강검진 제안서" if company else "2026 기업 임직원 건강검진 제안서"

    def normalize_text(text):
        return re.sub(r"(선택)\s*(\d+)", r"\1 \2", str(text))

    def get_val_display(val):
        if not val or val in ["X", "x", "-", "미선택"]:
            return ""
        if val in ["O", "o", "○"] or "기본" in str(val):
            return "O"
        if "선택" in str(val):
            return normalize_text(val)
        return str(val)

    def render_table(title, item_list, show_sub=False, footer=None, merge=True):
        if not item_list:
            return ""
        grid = []
        for item in item_list:
            row = [get_val_display(v) for v in item["values"]]
            grid.append(row)

        rows_cnt = len(grid)
        cols_cnt = len(plans)
        rowspan_map = [[1] * cols_cnt for _ in range(rows_cnt)]
        skip_map = [[False] * cols_cnt for _ in range(rows_cnt)]

        if merge:
            for c in range(cols_cnt):
                for r in range(rows_cnt):
                    if skip_map[r][c]:
                        continue
                    val = grid[r][c]
                    if val != "":
                        span = 1
                        for k in range(r + 1, rows_cnt):
                            if grid[k][c] == val:
                                span += 1
                                skip_map[k][c] = True
                            else:
                                break
                        rowspan_map[r][c] = span

        html_rows = ""
        for r in range(rows_cnt):
            item = item_list[r]
            sub_tag = f"<span class='cat-tag'>[{item['category']}]</span> " if show_sub and item.get("category") else ""
            row_str = f"<tr><td class='item-name-cell'>{sub_tag}{item['name']}</td>"
            for c in range(cols_cnt):
                if skip_map[r][c]:
                    continue
                val = grid[r][c]
                span = rowspan_map[r][c]
                cls = "text-center"
                if val == "O":
                    cls += " text-bold"
                elif "선택" in str(val):
                    cls += " text-navy text-bold"
                attr = f' rowspan="{span}"' if span > 1 else ""
                row_str += f'<td{attr} class="{cls}">{val}</td>'
            row_str += "</tr>"
            html_rows += row_str

        header_cols = "".join([f"<th>{p['name']}</th>" for p in plans])
        footer_div = f"<div class='table-footer'>{footer}</div>" if footer else ""
        return f"""<div class="section"><div class="sec-title">{title}</div>
        <table><thead><tr><th style="width:28%">검사 항목</th>{header_cols}</tr></thead>
        <tbody>{html_rows}</tbody></table>{footer_div}</div>"""

    a_vals = [s["a"] for s in summary]
    b_vals = [s["b"] for s in summary]
    c_vals = [s["c"] for s in summary]

    def make_sum_row(title, vals):
        tds = "".join([f"<td class='text-center'>{v}</td>" for v in vals])
        return f"<tr><td class='summary-header'>{title}</td>{tds}</tr>"

    sum_rows_html = make_sum_row("A그룹", a_vals) + make_sum_row("B그룹", b_vals) + make_sum_row("C그룹", c_vals)
    sum_headers = "".join([f"<th>{p['name']}</th>" for p in plans])

    # propsal2026.py의 1~3 페이지 구성을 그대로 반영
    return f"""
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <title>{proposal_title}</title>
  <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    @page {{ size: A4; margin: 10mm; }}
    body {{ font-family: 'Pretendard', sans-serif; background: #fff; margin: 0; padding: 20px; color: #333; font-size: 11px; }}
    .page {{ width: 210mm; min-height: 297mm; margin: 0 auto; background: white; padding: 15px 40px; box-sizing: border-box; }}
    .hospital-brand {{ font-size: 26px; font-weight: 900; color: #1a253a; letter-spacing: -1px; }}
    .hospital-sub {{ font-size: 16px; color: #555; margin-top: 5px; font-weight: bold; }}
    .contact-card {{ background-color: #f8f9fa; border: 2px solid #2c3e50; border-radius: 8px; padding: 10px 15px; text-align: right; box-shadow: 2px 2px 8px rgba(0,0,0,0.05); min-width: 200px; }}
    .contact-title {{ font-size: 10px; color: #7f8c8d; font-weight: bold; margin-bottom: 2px; }}
    .contact-name {{ font-size: 14px; font-weight: 800; color: #2c3e50; margin-bottom: 1px; }}
    .contact-info {{ font-size: 11px; color: #333; font-weight: 600; line-height: 1.3; }}
    header {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 15px; }}
    .header-divider {{ border-bottom: 2px solid #2c3e50; margin-bottom: 15px; }}
    .section {{ margin-bottom: 25px; page-break-inside: avoid; }}
    .sec-title {{ font-size: 15px; font-weight: 800; color: #2c3e50; margin-bottom: 8px; padding-left: 8px; border-left: 4px solid #2c3e50; }}
    table {{ width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 11px; border-top: 2px solid #2c3e50; }}
    th {{ background: #f0f2f5; color: #2c3e50; padding: 8px; border: 1px solid #bdc3c7; font-weight: bold; }}
    td {{ padding: 6px; border: 1px solid #bdc3c7; vertical-align: middle; word-break: keep-all; height: 24px; }}
    .summary-table th {{ background: #34495e; color: white; border-color: #2c3e50; }}
    .summary-header {{ background: #f8f9fa; font-weight: bold; color: #2c3e50; padding-left: 15px; text-align: left; }}
    .text-center {{ text-align: center; }}
    .text-bold {{ font-weight: bold; }}
    .text-navy {{ color: #2c3e50; }}
    .item-name-cell {{ text-align:left; padding-left:10px; width: 28%; font-weight: 600; }}
    .cat-tag {{ color: #7f8c8d; font-size: 10px; margin-right:3px; }}
    .table-footer {{ font-size: 11px; color: #2c3e50; text-align: right; margin-top: 5px; font-weight: bold; }}

    /* 1~2 섹션용 */
    .guide-box {{ background-color: #fff; border: 2px solid #2c3e50; padding: 15px; margin-bottom: 15px; font-size: 11px; line-height: 1.6; color: #333; }}
    .guide-title {{ font-weight: 800; font-size: 14px; margin-bottom: 10px; display:block; color: #2c3e50; border-bottom: 1px solid #ddd; padding-bottom: 5px; }}
    .highlight-text {{ font-weight: bold; color: #1a253a; }}
    .important-note {{ color: #c0392b; font-weight: bold; }}
    .program-grid {{ display: flex; flex-direction: column; gap: 6px; margin-bottom: 20px; border: 1px solid #ccc; padding: 6px; background: #fff; }}
    .grid-row {{ display: flex; gap: 6px; }}
    .grid-col {{ display: flex; flex-direction: column; gap: 6px; }}
    .grid-box {{ border: 1px solid #95a5a6; background: white; }}
    .grid-header {{ background: #34495e; color: white; padding: 6px 10px; font-weight: bold; font-size: 12px; text-align: center; }}
    .grid-content {{ padding: 10px; font-size: 11px; line-height: 1.5; color: #333; }}
    .grid-content-list {{ display: grid; grid-template-columns: 1fr 1fr; gap: 2px 10px; padding: 8px 10px; font-size: 11px; font-weight: 500; color: #444; }}
    .grid-sub-header {{ background: #ecf0f1; color: #2c3e50; padding: 4px 10px; font-weight: bold; font-size: 11px; border-bottom: 1px solid #ddd; }}
    .header-common {{ background: #2c3e50; font-size: 13px; text-align: left; padding-left: 15px; }}
    .header-a {{ background: #566573; }}
    .header-b {{ background: #7f8c8d; }}
    .header-c {{ background: #2c3e50; }}
    .page-break {{ page-break-after: always; }}

    @media print {{
      body {{ padding: 0; }}
      .page {{ width: 100%; padding: 0; border: none; }}
      td, th {{ -webkit-print-color-adjust: exact; vertical-align: middle !important; }}
      .summary-table th {{ background-color: #34495e !important; color: white !important; }}
      .guide-box, .contact-card {{ border: 2px solid #2c3e50 !important; }}
      .header-a, .header-b, .header-c, .header-common {{ color: white !important; }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <header>
      <div>
        <div class="hospital-brand">뉴고려병원</div>
        <div class="hospital-sub">{proposal_title}</div>
        <div style="font-size:11px; color:#666; margin-top:4px;">제안일자: {today_date}</div>
      </div>
      <div class="contact-card">
        <div class="contact-title">PROPOSAL CONTACT</div>
        <div class="contact-name">{mgr_name} 팀장</div>
        <div class="contact-info">📞 {mgr_phone}</div>
        <div class="contact-info">✉️ {mgr_email}</div>
      </div>
    </header>

    <div class="header-divider"></div>

    <div class="guide-box">
      <span class="guide-title">1. 유동적 그룹 선택 시스템 (Flexible Option)</span>
      <div style="display:flex; justify-content:space-between; align-items: flex-start; gap: 20px;">
        <div style="flex: 1;">
          <div style="margin-bottom: 6px; background-color:#ffebee; padding:4px 8px; border-radius:4px; border-left:3px solid #e57373;">
            • <b>A그룹 2개</b> <span style="color:#aaa">⇄</span> <span class="highlight-text">B그룹 1개</span> 로 변경 선택 가능
          </div>
          <div style="margin-bottom: 6px; background-color:#ffebee; padding:4px 8px; border-radius:4px; border-left:3px solid #e57373;">
            • <b>A그룹 4개</b> <span style="color:#aaa">⇄</span> <span class="highlight-text">C그룹 1개</span> 로 변경 선택 가능
          </div>
          <div style="margin-bottom: 6px; padding:2px 5px;">• <span class="highlight-text">유전자검사 20종</span> (기본제공) <span style="color:#aaa">⇄</span> <b>A그룹 1개</b> 로 변경 가능</div>
          <div style="padding:2px 5px;">• <span class="important-note">공단 위암 대상자</span> 위내시경 진행 시 <span class="highlight-text">A그룹 추가 1가지</span> 선택 가능</div>
        </div>
        <div style="flex: 0.8; border-left:3px solid #ddd; padding-left:20px; color:#2c3e50;">
          <span style="font-weight:bold; display:block; margin-bottom:8px; font-size:13px; color:#c0392b;">[비고: MRI 정밀 장비 안내]</span>
          <span style="font-weight:bold; font-size:14px; color:#000;">Full Protocol Scan 시행</span><br>
          <span style="color:#666; font-size:11px;">(Spot protocol 아님)</span><br>
          <span class="highlight-text" style="font-size:14px;">최신 3.0T MRI 장비 보유</span>
        </div>
      </div>
      <div style="margin-top:12px; font-style:italic; color:#666; font-size: 11px; padding-left:5px;">
        (예시: 70만원형 기본 [A5, B1, C1] → 변경 [A1, B3, C1] 또는 [A1, B2, C2] 등 자유롭게 조합 가능)
      </div>
    </div>

    <div class="program-grid">
      <div class="grid-box common-box">
        <div class="grid-header header-common">2. 상세 검진 항목 및 그룹 구성</div>
        <div class="grid-sub-header">공통 항목 <span style="font-weight:normal;">(위내시경 포함)</span></div>
        <div class="grid-content">
          간기능 | 간염 | 순환기계 | 당뇨 | 췌장기능 | 철결핍성 | 빈혈 | 혈액질환 | 전해질 | 신장기능 | 골격계질환<br>
          감염성 | 갑상선기능 | 부갑상선기능 | 종양표지자 | 소변 등 80여종 혈액(소변)검사<br>
          심전도 | 신장 | 체중 | 혈압 | 시력 | 청력 | 체성분 | 건강유형분석 | 폐기능 | 안저 | 안압<br>
          혈액점도검사 | 유전자20종 | 흉부X-ray | 복부초음파 | 위수면내시경<br>
          (여)자궁경부세포진 | (여)유방촬영 - #30세이상 권장#
        </div>
      </div>
      <div class="grid-row">
        <div class="grid-col" style="flex: 1.2;">
          <div class="grid-box">
            <div class="grid-header header-a">A 그룹 (정밀)</div>
            <div class="grid-content-list">
              <div>[01] 갑상선초음파</div> <div>[10] 골다공증QCT+비타민D</div>
              <div>[02] 경동맥초음파</div> <div>[11] 혈관협착도ABI</div>
              <div>[03] (여)경질초음파</div> <div>[12] (여)액상 자궁경부세포진</div>
              <div>[04] 뇌CT</div> <div>[13] (여) HPV바이러스</div>
              <div>[05] 폐CT</div> <div>[14] (여)(혈액)마스토체크:유방암</div>
              <div>[06] 요추CT</div> <div>[15] (혈액)NK뷰키트</div>
              <div>[07] 경추CT</div> <div>[16] (여)(혈액)여성호르몬</div>
              <div>[08] 심장MDCT</div> <div>[17] (남)(혈액)남성호르몬</div>
              <div>[09] 복부비만CT</div>
            </div>
          </div>
        </div>
        <div class="grid-col" style="flex: 1;">
          <div class="grid-box">
            <div class="grid-header header-b">B 그룹 (특화)</div>
            <div class="grid-content-list">
              <div>[가] 대장수면내시경</div> <div>[마] 부정맥검사S-PATCH</div>
              <div>[나] 심장초음파</div> <div>[바] [혈액]알레르기검사</div>
              <div>[다] (여)유방초음파</div> <div>[사] [혈액]알츠온:치매위험도</div>
              <div>[라] [분변]대장암_얼리텍</div> <div>[아] [혈액]간섬유화검사</div>
              <div></div> <div>[자] 폐렴예방접종:15가</div>
            </div>
          </div>
          <div class="grid-box" style="margin-top:5px; flex-grow:1;">
            <div class="grid-header header-c">C 그룹 (VIP)</div>
            <div class="grid-content-list">
              <div>[A] 뇌MRI+MRA</div>
              <div style="letter-spacing:-1.5px; white-space:nowrap;">[D] [혈액]스마트암검사(남6/여7종)</div>
              <div>[B] 경추MRI</div> <div>[E] [혈액]선천적 유전자검사</div>
              <div>[C] 요추MRI</div> <div>[F] [혈액]에피클락 (생체나이)</div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <div class="section">
      <div class="sec-title">3. 검진 프로그램 요약</div>
      <table class="summary-table">
        <thead>
          <tr>
            <th style="width:25%">구분</th>
            {sum_headers}
          </tr>
        </thead>
        <tbody>
          {sum_rows_html}
        </tbody>
      </table>
    </div>

    <div class="page-break"></div>

    {render_table("4. A 그룹 (정밀검사)", data['A'])}
    {render_table("5. B 그룹 (특화검사)", data['B'], footer="* A그룹 2개를 제외하고 B그룹 1개 선택 가능")}
    {render_table("6. C 그룹 (VIP검사)", data['C'], footer="* A그룹 4개를 제외하고 C그룹 1개 선택 가능")}

    <div class="page-break"></div>

    {render_table("7. 기초 장비 및 혈액 검사", data['EQUIP'] + data['COMMON_BLOOD'], show_sub=True, merge=False)}

    <div style="text-align:center; font-size:11px; color:#7f8c8d; margin-top:30px; padding-top:20px; border-top:1px solid #eee;">
      본 제안서는 귀사의 임직원 건강 증진을 위해 작성되었으며, 세부 검진 항목 및 일정은 협의에 따라 조정될 수 있습니다.
    </div>
  </div>
</body>
</html>
"""


# -----------------------------
# 4) 엑셀 생성 (propsal2026_patched.py의 generate_report_excel을 bytes로 변환)
# -----------------------------
def generate_excel_bytes(
    plans: List[Dict[str, Any]],
    data: Dict[str, List[Dict[str, Any]]],
    summary: List[Dict[str, str]],
    company: str,
    mgr_name: str,
    mgr_phone: str,
    mgr_email: str,
) -> bytes:
    company = (company or "").strip() or "기업"
    title_text = f"2026 {company} 임직원 건강검진 제안서"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "제안서"

    # A4 설정(숫자 9) 및 여백
    ws.page_setup.paperSize = 9
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    # Styles
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    box_side = Side(style="medium", color="2C3E50")

    header_fill = PatternFill(start_color="F0F2F5", end_color="F0F2F5", fill_type="solid")
    sum_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 섹션 외곽 테두리 헬퍼
    def draw_box_border(ws_, min_r: int, max_r: int, min_c: int, max_c: int):
        # Top
        for c in range(min_c, max_c + 1):
            cell = ws_.cell(row=min_r, column=c)
            old = cell.border
            cell.border = Border(left=old.left, right=old.right, top=box_side, bottom=old.bottom)
        # Bottom
        for c in range(min_c, max_c + 1):
            cell = ws_.cell(row=max_r, column=c)
            old = cell.border
            cell.border = Border(left=old.left, right=old.right, top=old.top, bottom=box_side)
        # Left
        for r in range(min_r, max_r + 1):
            cell = ws_.cell(row=r, column=min_c)
            old = cell.border
            cell.border = Border(left=box_side, right=old.right, top=old.top, bottom=old.bottom)
        # Right
        for r in range(min_r, max_r + 1):
            cell = ws_.cell(row=r, column=max_c)
            old = cell.border
            cell.border = Border(left=old.left, right=box_side, top=old.top, bottom=old.bottom)

    # 1. Header
    ws["A1"] = "뉴고려병원"
    ws["A1"].font = Font(size=16, bold=True, color="1A253A")
    ws["A2"] = title_text
    ws["A2"].font = Font(size=14, bold=True)
    ws["A3"] = f"제안일자: {datetime.now().strftime('%Y-%m-%d')}"

    # last_col: 실제 데이터 열(플랜 수 + 항목열 1개)
    last_col = len(plans) + 1
    if last_col < 3:
        last_col = 3

    # 담당자 정보 (오른쪽 상단)
    ws.merge_cells(start_row=1, start_column=last_col - 1, end_row=1, end_column=last_col)
    ws.cell(row=1, column=last_col - 1, value="담당자").font = Font(bold=True, color="7F8C8D")
    ws.cell(row=1, column=last_col - 1).alignment = Alignment(horizontal="right")

    ws.merge_cells(start_row=2, start_column=last_col - 1, end_row=2, end_column=last_col)
    ws.cell(row=2, column=last_col - 1, value=f"{mgr_name} 팀장").font = Font(bold=True, size=12)
    ws.cell(row=2, column=last_col - 1).alignment = Alignment(horizontal="right")

    ws.merge_cells(start_row=3, start_column=last_col - 1, end_row=3, end_column=last_col)
    ws.cell(row=3, column=last_col - 1, value=mgr_phone).alignment = Alignment(horizontal="right")

    ws.merge_cells(start_row=4, start_column=last_col - 1, end_row=4, end_column=last_col)
    ws.cell(row=4, column=last_col - 1, value=mgr_email).alignment = Alignment(horizontal="right")

    current_row = 6

    # -------------------------
    # 1. 유동적 그룹 선택 시스템
    # -------------------------
    section1_title_row = current_row
    ws.cell(row=current_row, column=1, value="1. 유동적 그룹 선택 시스템 (Flexible Option)").font = Font(
        bold=True, size=12, color="2C3E50"
    )
    ws.merge_cells(start_row=section1_title_row, start_column=1, end_row=section1_title_row, end_column=last_col)
    ws.cell(row=section1_title_row, column=1).alignment = left_align
    current_row += 1

    guide_text = (
        "• A그룹 2개 ⇄ B그룹 1개 로 변경 선택 가능\n"
        "• A그룹 4개 ⇄ C그룹 1개 로 변경 선택 가능\n"
        "• 유전자검사 20종 (기본제공) ⇄ A그룹 1개 로 변경 가능\n"
        "• 공단 위암 대상자 위내시경 진행 시 A그룹 추가 1가지 선택 가능\n\n"
        "[비고: MRI 정밀 장비 안내]\n"
        "Full Protocol Scan 시행 (Spot protocol 아님) / 최신 3.0T MRI 장비 보유\n"
        "(예시: 70만원형 기본 [A5, B1, C1] → 변경 [A1, B3, C1] 또는 [A1, B2, C2] 등 자유롭게 조합 가능)"
    )
    start_r = current_row
    end_r = current_row + 6

    ws.merge_cells(start_row=start_r, start_column=1, end_row=end_r, end_column=last_col)
    cell = ws.cell(row=start_r, column=1, value=guide_text)
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left", indent=1)

    # 외곽 테두리(제목행 포함)
    draw_box_border(ws, section1_title_row, end_r, 1, last_col)

    # 행 높이(요청값: 25)
    for r in range(start_r, end_r + 1):
        ws.row_dimensions[r].height = 25

    current_row += 8

    # -------------------------
    # 2. 상세 검진 항목 및 그룹 구성
    # -------------------------
    section2_title_row = current_row
    ws.cell(row=current_row, column=1, value="2. 상세 검진 항목 및 그룹 구성").font = Font(
        bold=True, size=12, color="2C3E50"
    )
    ws.merge_cells(start_row=section2_title_row, start_column=1, end_row=section2_title_row, end_column=last_col)
    ws.cell(row=section2_title_row, column=1).alignment = left_align
    current_row += 1

    text_common = (
        "간기능 | 간염 | 순환기계 | 당뇨 | 췌장기능 | 철결핍성 | 빈혈 | 혈액질환 | 전해질 | 신장기능 | 골격계질환\n"
        "감염성 | 갑상선기능 | 부갑상선기능 | 종양표지자 | 소변 등 80여종 혈액(소변)검사\n"
        "심전도 | 신장 | 체중 | 혈압 | 시력 | 청력 | 체성분 | 건강유형분석 | 폐기능 | 안저 | 안압\n"
        "혈액점도검사 | 유전자20종 | 흉부X-ray | 복부초음파 | 위수면내시경\n"
        "(여)자궁경부세포진 | (여)유방촬영 - #30세이상 권장#"
    )

    text_a = (
        "[01] 갑상선초음파  [10] 골다공증QCT+비타민D\n"
        "[02] 경동맥초음파  [11] 혈관협착도ABI\n"
        "[03] (여)경질초음파  [12] (여)액상 자궁경부세포진\n"
        "[04] 뇌CT  [13] (여) HPV바이러스\n"
        "[05] 폐CT  [14] (여)(혈액)마스토체크:유방암\n"
        "[06] 요추CT  [15] (혈액)NK뷰키트\n"
        "[07] 경추CT  [16] (여)(혈액)여성호르몬\n"
        "[08] 심장MDCT  [17] (남)(혈액)남성호르몬\n"
        "[09] 복부비만CT"
    )

    text_b = (
        "[가] 대장수면내시경  [마] 부정맥검사S-PATCH\n"
        "[나] 심장초음파  [바] [혈액]알레르기검사\n"
        "[다] (여)유방초음파  [사] [혈액]알츠온:치매위험도\n"
        "[라] [분변]대장암_얼리텍  [아] [혈액]간섬유화검사\n"
        "[자] 폐렴예방접종:15가"
    )

    text_c = (
        "[A] 뇌MRI+MRA  [D] [혈액]스마트암검사(남6/여7종)\n"
        "[B] 경추MRI  [E] [혈액]선천적 유전자검사\n"
        "[C] 요추MRI  [F] [혈액]에피클락 (생체나이)"
    )

    # 공통 항목 박스
    box_start_row = current_row

    ws.cell(row=current_row, column=1, value="공통 항목 (위내시경 포함)").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=current_row, column=1).fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=last_col)
    ws.cell(row=current_row, column=1).alignment = center_align
    current_row += 1

    content_start = current_row
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 4, end_column=last_col)
    c = ws.cell(row=current_row, column=1, value=text_common)
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left", indent=1)
    c.border = thin_border

    # 공통 항목 내용 행 높이 20(요청값)
    for r in range(content_start, current_row + 5):
        ws.row_dimensions[r].height = 20

    box_end_row = current_row + 4
    draw_box_border(ws, box_start_row, box_end_row, 1, last_col)

    current_row += 5

    # A/B/C 그룹 박스
    def write_group_box(title: str, text: str, color_hex: str, row_h: int):
        nonlocal current_row
        b_start = current_row

        # Header merge (col 1, 4 rows)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 3, end_column=1)
        cell_h = ws.cell(row=current_row, column=1, value=title)
        cell_h.font = Font(bold=True, color="FFFFFF")
        cell_h.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
        cell_h.alignment = center_align

        # Content merge (col 2~last, 4 rows)
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row + 3, end_column=last_col)
        cell_c = ws.cell(row=current_row, column=2, value=text)
        cell_c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left", indent=1)
        cell_c.border = thin_border

        for r in range(current_row, current_row + 4):
            ws.row_dimensions[r].height = row_h

        b_end = current_row + 3
        draw_box_border(ws, b_start, b_end, 1, last_col)
        current_row += 4

    # 요청값: A 40 / B 25 / C 21
    write_group_box("A 그룹\n(정밀)", text_a, "566573", 40)
    write_group_box("B 그룹\n(특화)", text_b, "7F8C8D", 25)
    write_group_box("C 그룹\n(VIP)", text_c, "2C3E50", 21)

    # 2번 섹션 전체 외곽 테두리(제목행 포함)
    section2_end_row = current_row - 1
    draw_box_border(ws, section2_title_row, section2_end_row, 1, last_col)

    current_row += 1

    # -------------------------
    # 3. 요약
    # -------------------------
    ws.cell(row=current_row, column=1, value="3. 검진 프로그램 요약").font = Font(bold=True, size=12)
    current_row += 1

    ws.cell(row=current_row, column=1, value="구분").fill = sum_fill
    ws.cell(row=current_row, column=1).font = white_font
    ws.cell(row=current_row, column=1).alignment = center_align

    for i, p in enumerate(plans):
        cell_ = ws.cell(row=current_row, column=i + 2, value=p["name"])
        cell_.fill = sum_fill
        cell_.font = white_font
        cell_.alignment = center_align

    current_row += 1

    def write_sum_row(title: str, vals: List[str]):
        nonlocal current_row
        ws.cell(row=current_row, column=1, value=title).font = Font(bold=True)
        ws.cell(row=current_row, column=1).border = thin_border
        ws.cell(row=current_row, column=1).alignment = left_align
        for i, v in enumerate(vals):
            cell_ = ws.cell(row=current_row, column=i + 2, value=v)
            cell_.alignment = center_align
            cell_.border = thin_border
        current_row += 1

    write_sum_row("A그룹", [s["a"] for s in summary])
    write_sum_row("B그룹", [s["b"] for s in summary])
    write_sum_row("C그룹", [s["c"] for s in summary])

    current_row += 1

    # 1페이지 종료(페이지 나누기) — 1~3이 1페이지에 들어가도록
    ws.row_breaks.append(Break(id=current_row))
    current_row += 1

    # -------------------------
    # 4~7 상세 표 (템플릿 데이터)
    # -------------------------
    def write_section(title: str, items: List[Dict[str, Any]], merge: bool = True):
        nonlocal current_row
        if not items:
            return

        ws.cell(row=current_row, column=1, value=title).font = Font(bold=True, size=12, color="2C3E50")
        current_row += 1

        # Header
        ws.cell(row=current_row, column=1, value="검사 항목").fill = header_fill
        ws.cell(row=current_row, column=1).border = thin_border
        ws.cell(row=current_row, column=1).alignment = center_align

        for i, p in enumerate(plans):
            h = ws.cell(row=current_row, column=i + 2, value=p["name"])
            h.fill = header_fill
            h.border = thin_border
            h.alignment = center_align

        current_row += 1
        start_row = current_row

        def norm(v):
            if not v or v in ["-", "미선택", "X"]:
                return ""
            if "선택" in str(v):
                return re.sub(r"(선택)\s*(\d+)", r"\1 \2", str(v))
            if "O" in str(v) or "기본" in str(v):
                return "O"
            return str(v)

        grid: List[List[str]] = []
        for item in items:
            row_vals = [norm(v) for v in item["values"]]
            grid.append(row_vals)

            name_val = f"[{item['category']}] {item['name']}" if item.get("category") else item["name"]
            c0 = ws.cell(row=current_row, column=1, value=name_val)
            c0.border = thin_border
            c0.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            for i, v in enumerate(row_vals):
                c1 = ws.cell(row=current_row, column=i + 2, value=v)
                c1.border = thin_border
                c1.alignment = center_align
                if v == "O":
                    c1.font = Font(bold=True)

            current_row += 1

        # 동일 값 세로 병합
        if merge:
            for c_idx in range(len(plans)):
                r = 0
                while r < len(grid):
                    val = grid[r][c_idx]
                    if val:
                        span = 1
                        for k in range(r + 1, len(grid)):
                            if grid[k][c_idx] == val:
                                span += 1
                            else:
                                break
                        if span > 1:
                            ws.merge_cells(
                                start_row=start_row + r,
                                start_column=c_idx + 2,
                                end_row=start_row + r + span - 1,
                                end_column=c_idx + 2,
                            )
                            ws.cell(row=start_row + r, column=c_idx + 2).alignment = center_align
                        r += span
                    else:
                        r += 1

        current_row += 2

    write_section("4. A 그룹 (정밀검사)", data.get("A", []))
    write_section("5. B 그룹 (특화검사)", data.get("B", []))
    write_section("6. C 그룹 (VIP검사)", data.get("C", []))

    ws.row_breaks.append(Break(id=current_row))
    current_row += 1

    write_section("7. 기초 장비 및 혈액 검사", (data.get("EQUIP", []) + data.get("COMMON_BLOOD", [])), merge=False)

    # Column widths
    ws.column_dimensions["A"].width = 32
    for i in range(len(plans)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 20

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
