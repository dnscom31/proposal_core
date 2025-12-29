# proposal_core.py
# 웹(Streamlit)에서 건강검진 제안서 HTML/엑셀을 생성하기 위한 코어 모듈

import io
import re
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break


# -------------------------
# Excel 템플릿 스캔/파싱
# -------------------------
def scan_default_counts(sheet, col_idx, start_row):
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
                val = int(nums[0])
                if val > counts[current_cat]:
                    counts[current_cat] = val
    return counts


def load_price_options(excel_filename):
    """
    엑셀 템플릿에서 '만원' 헤더행을 찾고,
    각 금액 컬럼(col_idx)과 기본 선택(A/B/C) 값을 옵션으로 반환
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

    # (propsal2026.py에서 사용하던 수동 기본값을 동일하게 반영)
    manual_defaults = {
        25: {"a": 3, "b": 0, "c": 0}, 30: {"a": 3, "b": 0, "c": 0},
        35: {"a": 4, "b": 0, "c": 0}, 40: {"a": 5, "b": 0, "c": 0},
        45: {"a": 4, "b": 1, "c": 0}, 50: {"a": 5, "b": 1, "c": 0},
        60: {"a": 3, "b": 1, "c": 1}, 70: {"a": 5, "b": 1, "c": 1},
        80: {"a": 5, "b": 2, "c": 1}, 90: {"a": 5, "b": 3, "c": 1},
        100: {"a": 3, "b": 3, "c": 2},
    }

    options = []
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
                "defaults": defaults
            })

    wb.close()
    options.sort(key=lambda x: int(re.sub(r"[^0-9]", "", x["price_txt"]) or "999"))
    return header_row_idx, options


def parse_data(excel_filename, header_row, plans):
    """
    템플릿 엑셀을 읽어서 A/B/C/EQUIP/COMMON_BLOOD 항목 테이블 데이터를 생성
    """
    wb = openpyxl.load_workbook(excel_filename, data_only=True)
    sheet = wb.active

    parsed_data = {"A": [], "B": [], "C": [], "EQUIP": [], "COMMON_BLOOD": []}
    summary_info = [{"name": p["name"], "a": p["a_rule"], "b": p["b_rule"], "c": p["c_rule"]} for p in plans]

    # propsal2026.py의 “선택” 캐시 로직을 동일 반영
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

        row_vals = []
        for idx, plan in enumerate(plans):
            col_idx0 = plan["col_idx"] - 1
            val = str(row[col_idx0]).strip() if col_idx0 < len(row) and row[col_idx0] else ""

            if current_main_cat in ["A", "B", "C"]:
                cache = fill_cache[idx]
                if "선택" in val:
                    cache[current_main_cat] = val
                elif val == "" and cache[current_main_cat]:
                    val = cache[current_main_cat]
                elif val != "":
                    cache[current_main_cat] = None

            # 웹 입력(a_rule/b_rule/c_rule)로 선택 규칙 override
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


# -------------------------
# HTML 생성
# -------------------------
def render_html(plans, data, summary, company, mgr_name, mgr_phone, mgr_email):
    """
    propsal2026.py의 HTML 구조를 웹용으로 동일하게 구성:
      - 1. 유동적 그룹 선택 시스템 (guide-box)
      - 2. 상세 검진 항목 및 그룹 구성 (program-grid)
      - 3. 요약
      - 4~7 표
    """
    today_date = datetime.now().strftime("%Y년 %m월 %d일")
    mgr_name = mgr_name or "담당자"
    mgr_phone = mgr_phone or ""
    mgr_email = mgr_email or ""
    company = (company or "").strip()
    proposal_title = f"2026 {company} 임직원 건강검진 제안서" if company else "2026 기업 임직원 건강검진 제안서"

    # propsal2026.py에서 쓰던 고정 텍스트(2. 상세 구성)
    text_common = (
        "간기능 | 간염 | 순환기계 | 당뇨 | 췌장기능 | 철결핍성 | 빈혈 | 혈액질환 | 전해질 | 신장기능 | 골격계질환<br>"
        "감염성 | 갑상선기능 | 부갑상선기능 | 종양표지자 | 소변 등 80여종 혈액(소변)검사<br>"
        "심전도 | 신장 | 체중 | 혈압 | 시력 | 청력 | 체성분 | 건강유형분석 | 폐기능 | 안저 | 안압<br>"
        "혈액점도검사 | 유전자20종 | 흉부X-ray | 복부초음파 | 위수면내시경<br>"
        "(여)자궁경부세포진 | (여)유방촬영 - #30세이상 권장#"
    )
    text_a = (
        "[01] 갑상선초음파  [10] 골다공증QCT+비타민D<br>"
        "[02] 경동맥초음파  [11] 혈관협착도ABI<br>"
        "[03] (여)경질초음파  [12] (여)액상 자궁경부세포진<br>"
        "[04] 뇌CT  [13] (여) HPV바이러스<br>"
        "[05] 폐CT  [14] (여)(혈액)마스토체크:유방암<br>"
        "[06] 요추CT  [15] (혈액)NK뷰키트<br>"
        "[07] 경추CT  [16] (여)(혈액)여성호르몬<br>"
        "[08] 심장MDCT  [17] (남)(혈액)남성호르몬<br>"
        "[09] 복부비만CT"
    )
    text_b = (
        "[가] 대장수면내시경  [마] 부정맥검사S-PATCH<br>"
        "[나] 심장초음파  [바] [혈액]알레르기검사<br>"
        "[다] (여)유방초음파 [사] [혈액]알츠온:치매위험도<br>"
        "[라] [분변]대장암_얼리텍 [아][혈액]간섬유화<br>"        
        "A그룹 2개 ⇄ B그룹 1개 변경 가능"
    )
    text_c = (
        "[A] 뇌MRI+A  [D][혈액]스마트암(6/7종)<br>"
        "[B] 경추MRI [E][혈액]선천적유전자34종 (3.0T)<br>"
        "[C] 요추MRI [F][혈액]에피클락(생체나이)  "
        "A그룹 4개 ⇄ C그룹 1개 변경 가능"
    )

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
            sub_tag = f"<span class='cat-tag'>[{item['category']}]</span> " if show_sub and item["category"] else ""
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

    # 요약 표
    a_vals = [s["a"] for s in summary]
    b_vals = [s["b"] for s in summary]
    c_vals = [s["c"] for s in summary]

    def make_sum_row(title, vals):
        tds = "".join([f"<td class='text-center'>{v}</td>" for v in vals])
        return f"<tr><td class='summary-header'>{title}</td>{tds}</tr>"

    sum_rows_html = make_sum_row("A그룹", a_vals) + make_sum_row("B그룹", b_vals) + make_sum_row("C그룹", c_vals)
    sum_headers = "".join([f"<th>{p['name']}</th>" for p in plans])

    # C그룹 D항목 줄바꿈 방지(요청 반영)
    text_c_html = text_c.replace(
        "[D] (여)(혈액)스마트암검사(유방) - #60만원 상당#",
        '<span style="letter-spacing:-1.5px; white-space:nowrap;">[D] (여)(혈액)스마트암검사(유방) - #60만원 상당#</span>'
    )

    guide_html = """
    <div class="guide-box">
      <span class="guide-title">1. 유동적 그룹 선택 시스템 (Flexible Option)</span>
      <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:20px;">
        <div style="flex:1;">
          <div style="margin-bottom:6px; background-color:#ffebee; padding:4px 8px; border-radius:4px; border-left:3px solid #e57373;">
            • <b>A그룹 2개</b> <span style="color:#aaa">⇄</span> <span class="highlight-text">B그룹 1개</span> 로 변경 선택 가능
          </div>
          <div style="margin-bottom:6px; padding:2px 5px;">
            • <b>A그룹 4개</b> <span style="color:#aaa">⇄</span> <span class="highlight-text">C그룹 1개</span> 로 변경 선택 가능
          </div>
          <div style="margin-bottom:6px; padding:2px 5px;">
            • <span class="highlight-text">유전자검사 20종</span> (기본제공) <span style="color:#aaa">⇄</span> <b>A그룹 1개</b> 로 변경 가능
          </div>
          <div style="padding:2px 5px;">
            • <span class="important-note">공단 위암 대상자</span> 위내시경 진행 시 <span class="highlight-text">A그룹 추가 1가지</span> 선택 가능
          </div>
        </div>
        <div style="flex:0.8; border-left:3px solid #ddd; padding-left:20px; color:#2c3e50;">
          <span style="font-weight:bold; display:block; margin-bottom:8px; font-size:13px; color:#c0392b;">[비고: MRI 정밀 장비 안내]</span>
          <span style="font-weight:bold; font-size:14px; color:#000;">Full Protocol Scan 시행</span><br>
          <span style="color:#666; font-size:11px;">(Spot protocol 아님)</span><br>
          <span class="highlight-text" style="font-size:14px;">최신 3.0T MRI 장비 보유</span>
        </div>
      </div>
      <div style="margin-top:12px; font-style:italic; color:#666; font-size:11px; padding-left:5px;">
        (예시: 70만원형 기본 [A5, B1, C1] → 변경 [A1, B3, C1] 또는 [A1, B2, C2] 등 자유롭게 조합 가능)
      </div>
    </div>
    """

    program_grid_html = f"""
    <div class="program-grid">
      <div class="grid-box common-box">
        <div class="grid-header header-common">2. 상세 검진 항목 및 그룹 구성</div>
        <div class="grid-sub-header">공통 항목 <span style="font-weight:normal;">(위내시경 포함)</span></div>
        <div class="grid-content">{text_common}</div>
      </div>

      <div class="grid-row">
        <div class="grid-col" style="flex:1.2;">
          <div class="grid-box">
            <div class="grid-header header-a">A 그룹 (정밀)</div>
            <div class="grid-content-list">{text_a}</div>
          </div>
        </div>

        <div class="grid-col" style="flex:1;">
          <div class="grid-box">
            <div class="grid-header header-b">B 그룹 (특화)</div>
            <div class="grid-content">{text_b}</div>
          </div>
        </div>

        <div class="grid-col" style="flex:1;">
          <div class="grid-box">
            <div class="grid-header header-c">C 그룹 (VIP)</div>
            <div class="grid-content">{text_c_html}</div>
          </div>
        </div>
      </div>
    </div>
    """

    return f"""
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <title>{proposal_title}</title>
  <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    @page {{ size: A4; margin: 10mm; }}
    body {{ font-family: 'Pretendard', sans-serif; background:#fff; margin:0; padding:18px; color:#333; font-size:11px; }}
    .page {{ width:210mm; min-height:297mm; margin:0 auto; background:white; padding:14px 34px; box-sizing:border-box; }}
    .hospital-brand {{ font-size:26px; font-weight:900; color:#1a253a; letter-spacing:-1px; }}
    .hospital-sub {{ font-size:16px; color:#555; margin-top:5px; font-weight:bold; }}
    .contact-card {{ background:#f8f9fa; border:2px solid #2c3e50; border-radius:8px; padding:10px 15px; text-align:right;
                    box-shadow:2px 2px 8px rgba(0,0,0,0.05); min-width:200px; }}
    .contact-title {{ font-size:10px; color:#7f8c8d; font-weight:bold; margin-bottom:2px; }}
    .contact-name {{ font-size:14px; font-weight:800; color:#2c3e50; margin-bottom:1px; }}
    .contact-info {{ font-size:11px; color:#333; font-weight:600; line-height:1.3; }}

    header {{ display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:10px; }}
    .header-divider {{ border-bottom:2px solid #2c3e50; margin-bottom:10px; }}

    /* 1. 유동적 그룹 */
    .guide-box {{ border:2px solid #2c3e50; border-radius:8px; padding:10px 12px; margin-bottom:10px; }}
    .guide-title {{ display:block; font-size:14px; font-weight:800; color:#2c3e50; margin-bottom:6px; }}
    .highlight-text {{ color:#c0392b; font-weight:800; }}
    .important-note {{ color:#2c3e50; font-weight:800; }}

    /* 2. 상세 구성 */
    .program-grid {{ margin-bottom:10px; }}
    .grid-box {{ border:1px solid #bdc3c7; border-radius:8px; overflow:hidden; background:#fff; }}
    .grid-header {{ color:white; padding:6px 10px; font-weight:bold; font-size:12px; text-align:center; }}
    .header-common {{ background:#2c3e50; font-size:13px; text-align:left; padding-left:15px; }}
    .header-a {{ background:#566573; }}
    .header-b {{ background:#7f8c8d; }}
    .header-c {{ background:#2c3e50; }}
    .grid-sub-header {{ background:#ecf0f1; color:#2c3e50; padding:4px 10px; font-weight:bold; font-size:11px; border-bottom:1px solid #ddd; }}
    .grid-content {{ padding:8px 10px; font-size:11px; line-height:1.45; color:#333; }}
    .grid-content-list {{ padding:8px 10px; font-size:11px; line-height:1.45; color:#333; }}
    .grid-row {{ display:flex; gap:8px; margin-top:8px; }}
    .grid-col {{ display:flex; flex-direction:column; gap:8px; }}

    /* 표(요약/상세) */
    .section {{ margin-bottom:12px; page-break-inside: avoid; }}
    .sec-title {{ font-size:14px; font-weight:800; color:#2c3e50; margin-bottom:6px; padding-left:8px; border-left:4px solid #2c3e50; }}
    table {{ width:100%; border-collapse:collapse; table-layout:fixed; font-size:11px; border-top:2px solid #2c3e50; }}
    th {{ background:#f0f2f5; color:#2c3e50; padding:7px; border:1px solid #bdc3c7; font-weight:bold; }}
    td {{ padding:6px; border:1px solid #bdc3c7; vertical-align:middle; word-break:keep-all; height:22px; }}
    .summary-table th {{ background:#34495e; color:white; border-color:#2c3e50; }}
    .summary-header {{ background:#f8f9fa; font-weight:bold; color:#2c3e50; padding-left:12px; text-align:left; }}
    .text-center {{ text-align:center; }}
    .text-bold {{ font-weight:bold; }}
    .text-navy {{ color:#2c3e50; }}
    .item-name-cell {{ text-align:left; padding-left:10px; width:28%; font-weight:600; }}
    .cat-tag {{ color:#7f8c8d; font-size:10px; margin-right:3px; }}
    .table-footer {{ font-size:11px; color:#2c3e50; text-align:right; margin-top:4px; font-weight:bold; }}
    .page-break {{ page-break-after: always; }}

    @media print {{
      body {{ padding:0; }}
      .page {{ width:100%; padding:0; border:none; }}
      td, th {{ -webkit-print-color-adjust: exact; vertical-align: middle !important; }}
      .summary-table th {{ background-color:#34495e !important; color:white !important; }}
      .guide-box, .contact-card {{ border:2px solid #2c3e50 !important; }}
      .header-a, .header-b, .header-c, .header-common {{ color:white !important; }}
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

    {guide_html}
    {program_grid_html}

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
  </div>
</body>
</html>
"""


# -------------------------
# Excel 생성
# -------------------------
def generate_excel_bytes(plans, data, summary, company, mgr_name, mgr_phone, mgr_email):
    company = (company or "").strip() or "기업"
    mgr_name = mgr_name or "담당자"
    mgr_phone = mgr_phone or ""
    mgr_email = mgr_email or ""
    title_text = f"2026 {company} 임직원 건강검진 제안서"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "제안서"

    # 인쇄/레이아웃
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    box_side = Side(style="medium", color="2C3E50")

    header_fill = PatternFill(start_color="F0F2F5", end_color="F0F2F5", fill_type="solid")
    sum_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
    title_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    white_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    left_wrap_align = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)

    def draw_box_border(min_r, max_r, min_c, max_c):
        # 위/아래
        for c in range(min_c, max_c + 1):
            cell = ws.cell(row=min_r, column=c)
            old = cell.border
            cell.border = Border(left=old.left, right=old.right, top=box_side, bottom=old.bottom)
        for c in range(min_c, max_c + 1):
            cell = ws.cell(row=max_r, column=c)
            old = cell.border
            cell.border = Border(left=old.left, right=old.right, top=old.top, bottom=box_side)
        # 좌/우
        for r in range(min_r, max_r + 1):
            cell = ws.cell(row=r, column=min_c)
            old = cell.border
            cell.border = Border(left=box_side, right=old.right, top=old.top, bottom=old.bottom)
        for r in range(min_r, max_r + 1):
            cell = ws.cell(row=r, column=max_c)
            old = cell.border
            cell.border = Border(left=old.left, right=box_side, top=old.top, bottom=old.bottom)

    # 열 계산: 실제 데이터 마지막 열 = (A열=1) + 플랜 수
    last_col = len(plans) + 1

    # 헤더(병원/제안서/담당자)
    ws["A1"] = "뉴고려병원"
    ws["A1"].font = Font(size=16, bold=True, color="1A253A")
    ws["A2"] = title_text
    ws["A2"].font = Font(size=14, bold=True)
    ws["A3"] = f"제안일자: {datetime.now().strftime('%Y-%m-%d')}"
    ws["A3"].font = Font(size=10)

    # 담당자 영역(우측 2칸: last_col-1 ~ last_col)
    # last_col이 2인 경우(플랜 1개)에도 동작하도록 보호
    contact_start = max(2, last_col - 1)
    contact_end = max(2, last_col)

    ws.merge_cells(start_row=1, start_column=contact_start, end_row=1, end_column=contact_end)
    ws.cell(row=1, column=contact_start, value="담당자").font = Font(bold=True, color="7F8C8D")
    ws.cell(row=1, column=contact_start).alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells(start_row=2, start_column=contact_start, end_row=2, end_column=contact_end)
    ws.cell(row=2, column=contact_start, value=f"{mgr_name} 팀장").font = Font(bold=True, size=12)
    ws.cell(row=2, column=contact_start).alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells(start_row=3, start_column=contact_start, end_row=3, end_column=contact_end)
    ws.cell(row=3, column=contact_start, value=mgr_phone).alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells(start_row=4, start_column=contact_start, end_row=4, end_column=contact_end)
    ws.cell(row=4, column=contact_start, value=mgr_email).alignment = Alignment(horizontal="right", vertical="center")

    current_row = 6

    # --- 1. 유동적 그룹 선택 시스템 ---
    section1_title_row = current_row
    ws.cell(row=current_row, column=1, value="1. 유동적 그룹 선택 시스템 (Flexible Option)").font = Font(bold=True, size=12, color="2C3E50")
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=last_col)
    ws.cell(row=current_row, column=1).alignment = left_align
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

    # 외곽 테두리(제목 행 포함)
    draw_box_border(section1_title_row, end_r, 1, last_col)

    # 행 높이 25
    for r in range(start_r, end_r + 1):
        ws.row_dimensions[r].height = 25

    current_row = end_r + 2

    # --- 2. 상세 검진 항목 및 그룹 구성 ---
    section2_title_row = current_row
    ws.cell(row=current_row, column=1, value="2. 상세 검진 항목 및 그룹 구성").font = Font(bold=True, size=12, color="2C3E50")
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=last_col)
    ws.cell(row=current_row, column=1).alignment = left_align
    current_row += 1

    # 공통/A/B/C 박스를 “텍스트 박스” 형태로 엑셀에 구성(요청 반영)
    def write_group_box(title, body_text, header_color, content_rows, row_height):
        nonlocal current_row

        # 헤더
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=last_col)
        h = ws.cell(row=current_row, column=1, value=title)
        h.fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
        h.font = white_font
        h.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        # 테두리
        for c in range(1, last_col + 1):
            ws.cell(row=current_row, column=c).border = thin_border

        current_row += 1

        # 내용(여러 행으로 나눠 병합)
        start_body = current_row
        end_body = current_row + content_rows - 1
        ws.merge_cells(start_row=start_body, start_column=1, end_row=end_body, end_column=last_col)
        b = ws.cell(row=start_body, column=1, value=body_text)
        b.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        b.border = thin_border

        for r in range(start_body, end_body + 1):
            ws.row_dimensions[r].height = row_height
            for c in range(1, last_col + 1):
                ws.cell(row=r, column=c).border = thin_border

        current_row = end_body + 1

    # 공통(5행, 높이 20)
    common_body = (
        "간기능 | 간염 | 순환기계 | 당뇨 | 췌장기능 | 철결핍성 | 빈혈 | 혈액질환 | 전해질 | 신장기능 | 골격계질환\n"
        "감염성 | 갑상선기능 | 부갑상선기능 | 종양표지자 | 소변 등 80여종 혈액(소변)검사\n"
        "심전도 | 신장 | 체중 | 혈압 | 시력 | 청력 | 체성분 | 건강유형분석 | 폐기능 | 안저 | 안압\n"
        "혈액점도검사 | 유전자20종 | 흉부X-ray | 복부초음파 | 위수면내시경\n"
        "(여)자궁경부세포진 | (여)유방촬영 - #30세이상 권장#"
    )
    write_group_box("공통 항목 (위내시경 포함)", common_body, "2C3E50", content_rows=5, row_height=20)

    # A/B/C
    a_body = (
        "[01] 갑상선초음파  [10] 골다공증QCT+비타민D\n"
        "[02] 경동맥초음파  [11] 혈관협착도ABI\n"
        "[03] (여)경질초음파  [12] (여)액상 자궁경부세포진\n"
        "[04] 뇌CT  [13] (여) HPV바이러스\n"
        "[05] 폐CT  [14] (여)(혈액)마스토체크:유방암\n"
        "[06] 요추CT  [15] (혈액)NK뷰키트\n"
        "[07] 경추CT  [16] NK면역검사\n"
        "[08] (혈액)알츠온(치매)  [17] (혈액)피검사(간염)\n"
        "[09] (혈액)암 6종  [18] (혈액)암 8종"
    )
    write_group_box("A 그룹 (정밀)", a_body, "566573", content_rows=4, row_height=40)

    b_body = (
        "[A] A그룹 2개 ⇄ B그룹 1개 변경 가능\n"
        "[01] 전립선초음파  [07] MRA(뇌혈관) (3.0T)\n"
        "[02] 심장초음파  [08] 뇌MRI (3.0T)\n"
        "[03] MRI(요추) (3.0T)  [09] MRI(경추) (3.0T)\n"
        "[04] MRI(뇌) (3.0T)  [10] (여)유방초음파\n"
        "[05] CT(대장)  [11] (여)인유두종 바이러스 검사\n"
        "[06] (혈액)유전자 30종"
    )
    write_group_box("B 그룹 (특화)", b_body, "7F8C8D", content_rows=4, row_height=25)

    c_body = (
        "[A] A그룹 4개 ⇄ C그룹 1개 변경 가능\n"
        "[B] A그룹 2개 ⇄ B그룹 1개로 변경 가능\n"
        "[01] PET-CT  [04] (여)유방MRI\n"
        "[02] MRI(뇌+혈관) (3.0T)  [05] MRI(복부) (3.0T)\n"
        "[03] MRI(심장) (3.0T)  [D] (여)(혈액)스마트암검사(유방) - #60만원 상당#"
    )
    write_group_box("C 그룹 (VIP)", c_body, "2C3E50", content_rows=4, row_height=21)

    section2_end_row = current_row - 1
    draw_box_border(section2_title_row, section2_end_row, 1, last_col)
    current_row += 1

    # --- 3. 검진 프로그램 요약 ---
    ws.cell(row=current_row, column=1, value="3. 검진 프로그램 요약").font = Font(bold=True, size=12, color="2C3E50")
    current_row += 1

    ws.cell(row=current_row, column=1, value="구분").fill = sum_fill
    ws.cell(row=current_row, column=1).font = white_font
    ws.cell(row=current_row, column=1).alignment = center_align
    ws.cell(row=current_row, column=1).border = thin_border

    for i, p in enumerate(plans):
        c = ws.cell(row=current_row, column=i + 2, value=p["name"])
        c.fill = sum_fill
        c.font = white_font
        c.alignment = center_align
        c.border = thin_border

    current_row += 1

    def write_sum_row(title, vals):
        nonlocal current_row
        c0 = ws.cell(row=current_row, column=1, value=title)
        c0.font = Font(bold=True)
        c0.border = thin_border
        c0.alignment = left_align
        for i, v in enumerate(vals):
            cc = ws.cell(row=current_row, column=i + 2, value=v)
            cc.alignment = center_align
            cc.border = thin_border
        current_row += 1

    write_sum_row("A그룹", [s["a"] for s in summary])
    write_sum_row("B그룹", [s["b"] for s in summary])
    write_sum_row("C그룹", [s["c"] for s in summary])

    # 1페이지에 1/2/3 섹션을 모으고 싶으면, 여기서 페이지 브레이크
    current_row += 1
    ws.row_breaks.append(Break(id=current_row))
    current_row += 1

    # --- 4~7 표 출력 (기존 로직 유지) ---
    def norm(v):
        if not v or v in ["-", "미선택", "X", "x"]:
            return ""
        if "선택" in str(v):
            return re.sub(r"(선택)\s*(\d+)", r"선택 \2", str(v))
        if "O" in str(v) or "기본" in str(v):
            return "O"
        return str(v)

    def write_section(title, items, merge=True, footer=None):
        nonlocal current_row
        if not items:
            return

        ws.cell(row=current_row, column=1, value=title).font = Font(bold=True, size=12, color="2C3E50")
        current_row += 1

        # 헤더
        h0 = ws.cell(row=current_row, column=1, value="검사 항목")
        h0.fill = header_fill
        h0.border = thin_border
        h0.alignment = center_align

        for i, p in enumerate(plans):
            hc = ws.cell(row=current_row, column=i + 2, value=p["name"])
            hc.fill = header_fill
            hc.border = thin_border
            hc.alignment = center_align

        current_row += 1
        start_row = current_row

        grid = []
        for item in items:
            row_vals = [norm(v) for v in item["values"]]
            grid.append(row_vals)

            name_val = f"[{item['category']}] {item['name']}" if item.get("category") else item["name"]
            c = ws.cell(row=current_row, column=1, value=name_val)
            c.border = thin_border
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            for i, v in enumerate(row_vals):
                cc = ws.cell(row=current_row, column=i + 2, value=v)
                cc.border = thin_border
                cc.alignment = center_align
                if v == "O":
                    cc.font = Font(bold=True)

            current_row += 1

        # 동일값 세로 병합
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

        # footer(엑셀은 텍스트로 한 줄 추가)
        if footer:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=last_col)
            f = ws.cell(row=current_row, column=1, value=footer)
            f.alignment = Alignment(horizontal="right", vertical="center")
            f.font = Font(bold=True, color="2C3E50", size=10)
            current_row += 1

        current_row += 1

    write_section("4. A 그룹 (정밀검사)", data["A"], merge=True)
    write_section("5. B 그룹 (특화검사)", data["B"], merge=True, footer="* A그룹 2개를 제외하고 B그룹 1개 선택 가능")
    write_section("6. C 그룹 (VIP검사)", data["C"], merge=True, footer="* A그룹 4개를 제외하고 C그룹 1개 선택 가능")

    ws.row_breaks.append(Break(id=current_row))
    current_row += 1

    write_section("7. 기초 장비 및 혈액 검사", data["EQUIP"] + data["COMMON_BLOOD"], merge=False)

    # 열 너비
    ws.column_dimensions["A"].width = 32
    for i in range(len(plans)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 20

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

