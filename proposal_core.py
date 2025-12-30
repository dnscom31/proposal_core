# proposal_core.py
# Core logic for proposal generation (HTML + Excel)

from __future__ import annotations

import io
import re
from datetime import datetime
from typing import Any, Dict, List, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break


def scan_default_counts(ws, col_idx: int, start_row: int) -> Dict[str, int]:
    """
    Scan the sheet for rows that contain '선택 N' in the given column, and return max counts per group.
    """
    defaults = {"A": 0, "B": 0, "C": 0}
    current = None

    for r in range(start_row, ws.max_row + 1):
        main_cat = ws.cell(row=r, column=2).value
        if main_cat:
            main_cat = str(main_cat).strip()
            if main_cat in ("A", "B", "C"):
                current = main_cat
            elif main_cat in ("D", "E", "F", "G", "COMMON", "EQUIP"):
                current = None

        val = ws.cell(row=r, column=col_idx).value
        if current and isinstance(val, str) and val.startswith("선택"):
            try:
                n = int(val.replace("선택", "").strip())
                defaults[current] = max(defaults[current], n)
            except Exception:
                pass

    return {"a": defaults["A"], "b": defaults["B"], "c": defaults["C"]}


def load_price_options(excel_filename: str) -> Tuple[int, List[Dict[str, Any]]]:
    """
    Return (header_row_index, options)
    option: {"price_txt": "...", "col_idx": int, "defaults": {"a":int,"b":int,"c":int}}
    """
    wb = openpyxl.load_workbook(excel_filename, data_only=True)
    ws = wb.active

    header_row_idx = None
    for r in range(1, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(isinstance(v, str) and "만원" in v for v in row_vals):
            header_row_idx = r
            break

    if header_row_idx is None:
        wb.close()
        raise ValueError("엑셀에서 '만원'이 포함된 헤더 행을 찾지 못했습니다.")

    header_cells = list(ws.iter_rows(min_row=header_row_idx, max_row=header_row_idx, values_only=False))[0]

    manual_defaults = {
        65: {"a": 4, "b": 1, "c": 1},
        70: {"a": 5, "b": 1, "c": 1},
        80: {"a": 7, "b": 1, "c": 1},
        90: {"a": 8, "b": 1, "c": 1},
    }

    options: List[Dict[str, Any]] = []
    for cell in header_cells:
        v = cell.value
        if not (isinstance(v, str) and "만원" in v):
            continue
        if "10만원" in v or "15만원" in v:
            continue

        col_idx = cell.column
        price_txt = str(v).strip()
        m = re.search(r"(\d+)", price_txt)
        price_number = int(m.group(1)) if m else None

        if price_number in manual_defaults:
            defaults = dict(manual_defaults[price_number])
        else:
            defaults = scan_default_counts(ws, col_idx=col_idx, start_row=header_row_idx + 1)

        options.append({"price_txt": price_txt, "col_idx": col_idx, "defaults": defaults})

    def _sort_key(opt):
        m = re.search(r"(\d+)", opt["price_txt"])
        return int(m.group(1)) if m else 999999

    options.sort(key=_sort_key)

    wb.close()
    return header_row_idx, options


def parse_data(excel_filename: str, header_row: int, plans):
    wb = openpyxl.load_workbook(excel_filename)
    ws = wb.active

    data = {}
    current_main_cat = None
    current_sub_cat = None

    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(row):
            continue

        main_cat, sub_cat, item_name, description = row[:4]

        if main_cat:
            current_main_cat = str(main_cat).strip()
            current_sub_cat = None

        if sub_cat:
            current_sub_cat = str(sub_cat).strip()

        if not item_name:
            continue

        item_name_str = str(item_name).strip()
        description_str = str(description).strip() if description else ""

        values = []
        for p in plans:
            col_idx = p["col_idx"] - 1
            cell_value = row[col_idx] if col_idx < len(row) else None
            values.append(cell_value if cell_value else "")

        cat_key = current_main_cat if current_main_cat else "UNCLASSIFIED"
        if cat_key not in data:
            data[cat_key] = []

        data[cat_key].append({
            "sub_cat": current_sub_cat,
            "name": item_name_str,
            "desc": description_str,
            "values": values
        })

    # Summary info for HTML (matches original render_html usage)
    summary_info = []
    for p in plans:
        summary_info.append({
            "name": p["name"],
            "a": p.get("a_rule", "-"),
            "b": p.get("b_rule", "-"),
            "c": p.get("c_rule", "-")
        })

    wb.close()
    return data, summary_info


def create_summary_table(plans: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """
    Summary rows for Excel '3. 검진 프로그램 요약'.
    """
    rows = [
        {"label": "공통 항목", "key": "common", "fixed": "O"},
        {"label": "A 그룹 (정밀)", "key": "a_rule"},
        {"label": "B 그룹 (특화)", "key": "b_rule"},
        {"label": "C 그룹 (VIP)", "key": "c_rule"},
    ]
    summary: List[Dict[str, str]] = []
    for r in rows:
        row = {"label": r["label"]}
        for p in plans:
            if r.get("fixed"):
                row[p["name"]] = r["fixed"]
            else:
                row[p["name"]] = p.get(r["key"], "-")
        summary.append(row)
    return summary


def render_html(plans, data, summary, company="", mgr_name="담당자", mgr_phone="", mgr_email=""):

    today_date = datetime.now().strftime("%Y년 %m월 %d일")
    mgr_name = mgr_name or "담당자"
    mgr_phone = mgr_phone or ""
    mgr_email = mgr_email or ""
    company = (company or "").strip()
    proposal_title = f"2026 {company} 임직원 건강검진 제안서" if company else "2026 기업 임직원 건강검진 제안서"

    # Summary table HTML
    sum_headers = "<th>구분</th>" + "".join([f"<th>{p['name']}</th>" for p in plans])

    sum_rows_html = ""
    labels = ["A 그룹 (정밀)", "B 그룹 (특화)", "C 그룹 (VIP)"]
    keys = ["a", "b", "c"]
    for label, k in zip(labels, keys):
        row = f"<tr><td class='row-label'>{label}</td>"
        for s in summary:
            row += f"<td>{s.get(k, '')}</td>"
        row += "</tr>"
        sum_rows_html += row

    # Helpers for group tables
    def make_table_rows(items: List[Dict[str, Any]]):
        rows = ""
        for it in items:
            name = it.get("name", "")
            desc = it.get("desc", "")
            vals = it.get("values", [])
            rows += "<tr>"
            rows += f"<td class='item-name'>{name}<div class='item-desc'>{desc}</div></td>"
            for v in vals:
                rows += f"<td>{v if v is not None else ''}</td>"
            rows += "</tr>"
        return rows

    # Build HTML
    html = f"""
    <!DOCTYPE html>
    <html lang="ko">
        <head>
            <meta charset="UTF-8">
            <title>{proposal_title}</title>
            <style>
                body {{
                    font-family: 'Malgun Gothic', sans-serif;
                    background: #f4f6f9;
                    margin: 0;
                    padding: 20px;
                }}
                .page {{
                    max-width: 940px;
                    margin: 0 auto;
                    background: white;
                    padding: 30px;
                    box-shadow: 0 0 8px rgba(0,0,0,0.1);
                    border-radius: 8px;
                }}
                header {{
                    display: flex;
                    justify-content: space-between;
                    align-items: flex-start;
                }}
                .hospital-brand {{
                    font-size: 24px;
                    font-weight: bold;
                    color: #1a253a;
                }}
                .hospital-sub {{
                    font-size: 18px;
                    font-weight: bold;
                    margin-top: 5px;
                    color: #2c3e50;
                }}
                .contact-card {{
                    border: 2px solid #2c3e50;
                    border-radius: 8px;
                    padding: 10px 14px;
                    text-align: right;
                    min-width: 230px;
                }}
                .contact-title {{
                    font-size: 11px;
                    font-weight: bold;
                    color: #7f8c8d;
                }}
                .contact-name {{
                    font-size: 14px;
                    font-weight: bold;
                    margin-top: 4px;
                }}
                .contact-info {{
                    font-size: 12px;
                    color: #34495e;
                    margin-top: 2px;
                }}
                .header-divider {{
                    border-bottom: 2px solid #2c3e50;
                    margin: 15px 0 20px;
                }}
                .guide-box {{
                    border: 2px solid #2c3e50;
                    border-radius: 8px;
                    padding: 12px 14px;
                    background: #fdfdfd;
                }}
                .guide-title {{
                    font-weight: bold;
                    color: #2c3e50;
                    font-size: 13px;
                    display: block;
                    margin-bottom: 8px;
                }}
                .highlight-text {{
                    font-weight: bold;
                    color: #c0392b;
                }}
                .important-note {{
                    font-weight: bold;
                    color: #2c3e50;
                }}
                .program-grid {{
                    margin-top: 14px;
                    display: flex;
                    flex-direction: column;
                    gap: 8px;
                }}
                .grid-box {{
                    border: 1px solid #ccc;
                    border-radius: 6px;
                    overflow: hidden;
                    background: #fff;
                }}
                .grid-header {{
                    color: white;
                    padding: 6px 10px;
                    font-weight: bold;
                    font-size: 12px;
                    text-align: center;
                }}
                .grid-content {{
                    padding: 10px;
                    font-size: 11px;
                    line-height: 1.5;
                    color: #333;
                }}
                .grid-content-list {{
                    display: grid;
                    grid-template-columns: 1fr 1fr;
                    gap: 2px 10px;
                    padding: 8px 10px;
                    font-size: 11px;
                    font-weight: 500;
                    color: #444;
                }}
                .grid-sub-header {{
                    background: #ecf0f1;
                    color: #2c3e50;
                    padding: 4px 10px;
                    font-weight: bold;
                    font-size: 11px;
                    border-bottom: 1px solid #ddd;
                }}
                .header-common {{ background: #2c3e50; font-size: 13px; text-align: left; padding-left: 15px; }}
                .header-a {{ background: #566573; }}
                .header-b {{ background: #7f8c8d; }}
                .header-c {{ background: #2c3e50; }}
                .page-break {{ page-break-after: always; }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 12px;
                    font-size: 11px;
                }}
                th, td {{
                    border: 1px solid #ccc;
                    padding: 6px;
                    text-align: center;
                    vertical-align: middle;
                }}
                th {{
                    background: #34495e;
                    color: white;
                    font-weight: bold;
                }}
                td.row-label {{
                    font-weight: bold;
                    background: #f0f2f5;
                    text-align: left;
                }}
                td.item-name {{
                    text-align: left;
                    font-weight: bold;
                }}
                .item-desc {{
                    font-weight: normal;
                    color: #666;
                    font-size: 10px;
                    margin-top: 2px;
                }}
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
                        </div>
                        <div class="grid-col" style="flex: 0.9;">
                            <div class="grid-box">
                                <div class="grid-header header-c">C 그룹 (VIP)</div>
                                <div class="grid-content-list" style="grid-template-columns: 1fr; font-size:11px;">
                                    <div>[A] 뇌MRI+MRA</div>
                                    <div>[B] 경추MRI</div>
                                    <div>[C] 요추MRI</div>
                                    <div style="white-space:nowrap; letter-spacing:-0.3px;">[D] [혈액]스마트암검사(남6/여7종)</div>
                                    <div>[E] [혈액]선천적 유전자검사</div>
                                    <div>[F] [혈액]에피클락 (생체나이)</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <h3 style="margin-top:18px;">3. 검진 프로그램 요약</h3>
                <table class="summary-table">
                    <tr>{sum_headers}</tr>
                    {sum_rows_html}
                </table>

                <div class="page-break"></div>

                <h3>4. A 그룹 (정밀검사)</h3>
                <table>
                    <tr><th>검사 항목</th>{"".join([f"<th>{p['name']}</th>" for p in plans])}</tr>
                    {make_table_rows(data.get("A", []))}
                </table>

                <h3>5. B 그룹 (특화검사)</h3>
                <table>
                    <tr><th>검사 항목</th>{"".join([f"<th>{p['name']}</th>" for p in plans])}</tr>
                    {make_table_rows(data.get("B", []))}
                </table>

                <h3>6. C 그룹 (VIP검사)</h3>
                <table>
                    <tr><th>검사 항목</th>{"".join([f"<th>{p['name']}</th>" for p in plans])}</tr>
                    {make_table_rows(data.get("C", []))}
                </table>

                <div class="page-break"></div>

                <h3>7. 기초 장비 및 혈액 검사</h3>
                <table>
                    <tr><th>검사 항목</th>{"".join([f"<th>{p['name']}</th>" for p in plans])}</tr>
                    {make_table_rows(data.get("EQUIP", []) + data.get("COMMON_BLOOD", []))}
                </table>

            </div>
        </body>
    </html>
    """
    return html


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

    # A4 설정 (숫자 9) 및 여백
    ws.page_setup.paperSize = 9
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
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
    white_font = Font(color="FFFFFF", bold=True)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def draw_box_border(ws, min_r, max_r, min_c, max_c):
        for c in range(min_c, max_c + 1):
            cell = ws.cell(row=min_r, column=c)
            old = cell.border
            cell.border = Border(left=old.left, right=old.right, top=box_side, bottom=old.bottom)
        for c in range(min_c, max_c + 1):
            cell = ws.cell(row=max_r, column=c)
            old = cell.border
            cell.border = Border(left=old.left, right=old.right, top=old.top, bottom=box_side)
        for r in range(min_r, max_r + 1):
            cell = ws.cell(row=r, column=min_c)
            old = cell.border
            cell.border = Border(left=box_side, right=old.right, top=old.top, bottom=old.bottom)
        for r in range(min_r, max_r + 1):
            cell = ws.cell(row=r, column=max_c)
            old = cell.border
            cell.border = Border(left=old.left, right=box_side, top=old.top, bottom=old.bottom)

    ws["A1"] = "뉴고려병원"
    ws["A1"].font = Font(size=16, bold=True, color="1A253A")
    ws["A2"] = title_text
    ws["A2"].font = Font(size=14, bold=True)
    ws["A3"] = f"제안일자: {datetime.now().strftime('%Y-%m-%d')}"

    # last_col: 실제 출력 컬럼(구분 1열 + 플랜 수)
    last_col = 1 + len(plans)
    if last_col < 2:
        last_col = 2

    # 담당자 정보: 플랜이 1개면 마지막 컬럼 1칸만 사용, 2개 이상이면 마지막 2칸 merge
    mgr_start_col = last_col if last_col < 3 else (last_col - 1)

    ws.merge_cells(start_row=1, start_column=mgr_start_col, end_row=1, end_column=last_col)
    ws.cell(row=1, column=mgr_start_col, value="담당자").font = Font(bold=True, color="7F8C8D")
    ws.cell(row=1, column=mgr_start_col).alignment = Alignment(horizontal="right")

    ws.merge_cells(start_row=2, start_column=mgr_start_col, end_row=2, end_column=last_col)
    ws.cell(row=2, column=mgr_start_col, value=f"{mgr_name} 팀장").font = Font(bold=True, size=12)
    ws.cell(row=2, column=mgr_start_col).alignment = Alignment(horizontal="right")

    ws.merge_cells(start_row=3, start_column=mgr_start_col, end_row=3, end_column=last_col)
    ws.cell(row=3, column=mgr_start_col, value=mgr_phone).alignment = Alignment(horizontal="right")

    ws.merge_cells(start_row=4, start_column=mgr_start_col, end_row=4, end_column=last_col)
    ws.cell(row=4, column=mgr_start_col, value=mgr_email).alignment = Alignment(horizontal="right")

    current_row = 6

    ws.cell(row=current_row, column=1, value="1. 유동적 그룹 선택 시스템 (Flexible Option)").font = Font(
        bold=True, size=12, color="2C3E50"
    )
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=last_col)
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

    draw_box_border(ws, start_r, end_r, 1, last_col)
    for r in range(start_r, end_r + 1):
        ws.row_dimensions[r].height = 25

    current_row = end_r + 2

    ws.cell(row=current_row, column=1, value="2. 상세 검진 항목 및 그룹 구성").font = Font(bold=True, size=12, color="2C3E50")
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=last_col)
    current_row += 1

    text_common = (
        "간기능 | 간염 | 순환기계 | 당뇨 | 췌장기능 |\n"
        "철결핍성 | 빈혈 | 내분비 | 신장기능 |\n"
        "전립선 | 갑상선 | 염증 | 통풍 |\n"
        "골격계질환\n"
        "감염성 | 위장질환 | B형간염 | 에이즈 |\n"
        "류마티스 | 매독 | 성병 | 소변정밀 | 혈액(소변)검사\n"
        "심전도 | 폐기능 | 청력 | 눈(시력) |\n"
        "동맥경화 | 체성분검사 | 안저검사 | 안압\n"
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

    for r in range(content_start, content_start + 5):
        ws.row_dimensions[r].height = 20

    draw_box_border(ws, box_start_row, content_start + 4, 1, last_col)
    current_row = content_start + 5

    def write_group_box(title, text, color_hex, row_h):
        nonlocal current_row
        b_start = current_row

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 3, end_column=1)
        cell_h = ws.cell(row=current_row, column=1, value=title)
        cell_h.font = Font(bold=True, color="FFFFFF")
        cell_h.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
        cell_h.alignment = center_align

        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row + 3, end_column=last_col)
        cell_c = ws.cell(row=current_row, column=2, value=text)
        cell_c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left", indent=1)
        cell_c.border = thin_border

        for r in range(current_row, current_row + 4):
            ws.row_dimensions[r].height = row_h

        draw_box_border(ws, b_start, current_row + 3, 1, last_col)
        current_row += 4

    write_group_box("A 그룹\n(정밀)", text_a, "566573", 40)
    write_group_box("B 그룹\n(특화)", text_b, "7F8C8D", 25)
    write_group_box("C 그룹\n(VIP)", text_c, "2C3E50", 21)

    current_row += 1

    ws.cell(row=current_row, column=1, value="3. 검진 프로그램 요약").font = Font(bold=True, size=12, color="2C3E50")
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=last_col)
    current_row += 1

    ws.cell(row=current_row, column=1, value="구분").font = white_font
    ws.cell(row=current_row, column=1).fill = sum_fill
    ws.cell(row=current_row, column=1).alignment = center_align
    ws.cell(row=current_row, column=1).border = thin_border

    for i, p in enumerate(plans, start=2):
        ws.cell(row=current_row, column=i, value=p["name"]).font = white_font
        ws.cell(row=current_row, column=i).fill = sum_fill
        ws.cell(row=current_row, column=i).alignment = center_align
        ws.cell(row=current_row, column=i).border = thin_border

    current_row += 1

    for s in summary:
        ws.cell(row=current_row, column=1, value=s["label"]).fill = header_fill
        ws.cell(row=current_row, column=1).alignment = center_align
        ws.cell(row=current_row, column=1).border = thin_border
        for i, p in enumerate(plans, start=2):
            ws.cell(row=current_row, column=i, value=s.get(p["name"], "")).alignment = center_align
            ws.cell(row=current_row, column=i).border = thin_border
        ws.row_dimensions[current_row].height = 18
        current_row += 1

    # Page breaks
    ws.row_breaks.append(Break(id=current_row + 1))

    def write_items_table(title: str, items: List[Dict[str, Any]], footer: str | None = None):
        nonlocal current_row

        ws.cell(row=current_row, column=1, value=title).font = Font(bold=True, size=12, color="2C3E50")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=last_col)
        current_row += 1

        ws.cell(row=current_row, column=1, value="검사 항목").font = white_font
        ws.cell(row=current_row, column=1).fill = sum_fill
        ws.cell(row=current_row, column=1).alignment = center_align
        ws.cell(row=current_row, column=1).border = thin_border
        for i, p in enumerate(plans, start=2):
            ws.cell(row=current_row, column=i, value=p["name"]).font = white_font
            ws.cell(row=current_row, column=i).fill = sum_fill
            ws.cell(row=current_row, column=i).alignment = center_align
            ws.cell(row=current_row, column=i).border = thin_border
        current_row += 1

        for it in items:
            ws.cell(row=current_row, column=1, value=it.get("name", "")).alignment = left_align
            ws.cell(row=current_row, column=1).border = thin_border
            for i, p in enumerate(plans, start=2):
                v = it.get("values", [""] * len(plans))[i - 2] if it.get("values") else ""
                ws.cell(row=current_row, column=i, value=v).alignment = center_align
                ws.cell(row=current_row, column=i).border = thin_border
            ws.row_dimensions[current_row].height = 18
            current_row += 1

        if footer:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=last_col)
            fcell = ws.cell(row=current_row, column=1, value=footer)
            fcell.alignment = Alignment(horizontal="left", vertical="center")
            current_row += 2
        else:
            current_row += 1

    write_items_table("4. A 그룹 (정밀검사)", data.get("A", []))
    write_items_table("5. B 그룹 (특화검사)", data.get("B", []), footer="* A그룹 2개를 제외하고 B그룹 1개 선택 가능")
    write_items_table("6. C 그룹 (VIP검사)", data.get("C", []), footer="* A그룹 4개를 제외하고 C그룹 1개 선택 가능")

    ws.row_breaks.append(Break(id=current_row + 1))

    equip_items = (data.get("EQUIP", []) or []) + (data.get("COMMON_BLOOD", []) or [])
    write_items_table("7. 기초 장비 및 혈액 검사", equip_items)

    ws.column_dimensions["A"].width = 34
    for col in range(2, last_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18

    ws.print_area = f"A1:{get_column_letter(last_col)}{current_row}"

    bio = io.BytesIO()
    wb.save(bio)
    wb.close()
    return bio.getvalue()
