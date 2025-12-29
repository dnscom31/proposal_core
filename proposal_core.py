# proposal_core.py
import re
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break


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
    엑셀에서 '만원' 헤더 행을 찾고, 각 금액 컬럼의 col_idx/price_txt/default_counts를 반환
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
    wb = openpyxl.load_workbook(excel_filename, data_only=True)
    sheet = wb.active

    parsed_data = {"A": [], "B": [], "C": [], "EQUIP": [], "COMMON_BLOOD": []}
    summary_info = [{"name": p["name"], "a": p["a_rule"], "b": p["b_rule"], "c": p["c_rule"]} for p in plans]

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

            # 원본과 동일한 "선택" 이어붙이기 캐시
            if current_main_cat in ["A", "B", "C"]:
                cache = fill_cache[idx]
                if "선택" in val:
                    cache[current_main_cat] = val
                elif val == "" and cache[current_main_cat]:
                    val = cache[current_main_cat]
                elif val != "":
                    cache[current_main_cat] = None

            # 웹에서 사용자가 입력한 a_rule/b_rule/c_rule로 override
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


def render_html(plans, data, summary, company, mgr_name, mgr_phone, mgr_email):
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
        if "선택" in val:
            return normalize_text(val)
        return val

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

    a_vals = [s["a"] for s in summary]
    b_vals = [s["b"] for s in summary]
    c_vals = [s["c"] for s in summary]

    def make_sum_row(title, vals):
        tds = "".join([f"<td class='text-center'>{v}</td>" for v in vals])
        return f"<tr><td class='summary-header'>{title}</td>{tds}</tr>"

    sum_rows_html = make_sum_row("A그룹", a_vals) + make_sum_row("B그룹", b_vals) + make_sum_row("C그룹", c_vals)
    sum_headers = "".join([f"<th>{p['name']}</th>" for p in plans])

    # 원본 HTML 템플릿을 그대로 사용(필요시 추후 분리 가능)
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
    .page-break {{ page-break-after: always; }}
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


def generate_excel_bytes(plans, data, summary, company, mgr_name, mgr_phone, mgr_email):
    company = (company or "").strip() or "기업"
    title_text = f"2026 {company} 임직원 건강검진 제안서"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "제안서"

    ws.page_setup.paperSize = 9
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
    white_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def draw_box_border(min_r, max_r, min_c, max_c):
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

    last_col = max(len(plans) + 1, 3)

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

    # Summary
    ws.cell(row=current_row, column=1, value="3. 검진 프로그램 요약").font = Font(bold=True, size=12)
    current_row += 1

    ws.cell(row=current_row, column=1, value="구분").fill = sum_fill
    ws.cell(row=current_row, column=1).font = white_font
    ws.cell(row=current_row, column=1).alignment = center_align

    for i, p in enumerate(plans):
        c = ws.cell(row=current_row, column=i + 2, value=p["name"])
        c.fill = sum_fill
        c.font = white_font
        c.alignment = center_align

    current_row += 1

    def write_sum_row(title, vals):
        nonlocal current_row
        ws.cell(row=current_row, column=1, value=title).font = Font(bold=True)
        ws.cell(row=current_row, column=1).border = thin_border
        ws.cell(row=current_row, column=1).alignment = left_align
        for i, v in enumerate(vals):
            c = ws.cell(row=current_row, column=i + 2, value=v)
            c.alignment = center_align
            c.border = thin_border
        current_row += 1

    write_sum_row("A그룹", [s["a"] for s in summary])
    write_sum_row("B그룹", [s["b"] for s in summary])
    write_sum_row("C그룹", [s["c"] for s in summary])
    current_row += 1

    ws.row_breaks.append(Break(id=current_row))
    current_row += 1

    def write_section(title, items, merge=True):
        nonlocal current_row
        if not items:
            return

        ws.cell(row=current_row, column=1, value=title).font = Font(bold=True, size=12, color="2C3E50")
        current_row += 1

        ws.cell(row=current_row, column=1, value="검사 항목").fill = header_fill
        ws.cell(row=current_row, column=1).border = thin_border
        ws.cell(row=current_row, column=1).alignment = center_align

        for i, p in enumerate(plans):
            c = ws.cell(row=current_row, column=i + 2, value=p["name"])
            c.fill = header_fill
            c.border = thin_border
            c.alignment = center_align

        current_row += 1
        start_row = current_row

        def norm(v):
            if not v or v in ["-", "미선택", "X"]:
                return ""
            if "선택" in str(v):
                return re.sub(r"(선택)\s*(\d+)", r"\1 \2", str(v))
            if "O" in str(v) or "기본" in str(v):
                return "O"
            return v

        grid = []
        for item in items:
            row_vals = [norm(v) for v in item["values"]]
            grid.append(row_vals)

            name_val = f"[{item['category']}] {item['name']}" if item.get("category") else item["name"]
            c = ws.cell(row=current_row, column=1, value=name_val)
            c.border = thin_border
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            for i, v in enumerate(row_vals):
                c = ws.cell(row=current_row, column=i + 2, value=v)
                c.border = thin_border
                c.alignment = center_align
                if v == "O":
                    c.font = Font(bold=True)

            current_row += 1

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

    write_section("4. A 그룹 (정밀검사)", data["A"])
    write_section("5. B 그룹 (특화검사)", data["B"])
    write_section("6. C 그룹 (VIP검사)", data["C"])

    ws.row_breaks.append(Break(id=current_row))
    current_row += 1

    write_section("7. 기초 장비 및 혈액 검사", data["EQUIP"] + data["COMMON_BLOOD"], merge=False)

    ws.column_dimensions["A"].width = 32
    for i in range(len(plans)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 20

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
