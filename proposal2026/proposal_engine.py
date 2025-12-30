# proposal_engine.py
# -*- coding: utf-8 -*-
"""
Proposal Maker Engine
- Loads and saves an external HTML template (proposal_template.html)
- Allows page/table/icon blocks to be edited (via markers)
- Stores uploaded images locally (copied into program folder) so paths do not break
- Resizes images to their target boxes (cover-crop) and embeds as Base64 in final output
- Adds "text-only blocks" per page: user edits plain text, engine converts to safe HTML
- Adds layout controls (spacing/sizes) via CSS variables
"""

from __future__ import annotations

import base64
import hashlib
import html
import json
import os
import re
import shutil
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from PIL import Image


# -----------------------------
# Marker patterns (template editing)
# -----------------------------
PAGE_BLOCK_RE = re.compile(r'<!--PAGE_START:(\d+)-->\s*(.*?)\s*<!--PAGE_END:\1-->', re.S)
TABLE_BLOCK_RE = re.compile(r'<!--TABLE_START:(\d+)-->\s*(.*?)\s*<!--TABLE_END:\1-->', re.S)
ICON_GROUP_RE = re.compile(r'<!--ICON_GROUP_START:([a-zA-Z0-9_\-]+)-->\s*(.*?)\s*<!--ICON_GROUP_END:\1-->', re.S)
TEXT_BLOCK_RE = re.compile(r'<!--TEXT_BLOCK_START:([a-zA-Z0-9_\-]+)-->\s*(.*?)\s*<!--TEXT_BLOCK_END:\1-->', re.S)

# Fallback (if template has no markers)
RAW_PAGE_START_RE = re.compile(r'<div class="page"\b', re.I)
RAW_TABLE_RE = re.compile(r'<table[^>]*>.*?</table>', re.S | re.I)


# -----------------------------
# Utilities
# -----------------------------
def _safe_read_text(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def _safe_write_text(path: str, content: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


def _hash_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def _remove_all_markers(html_text: str) -> str:
    # remove only our editor markers; keep other comments (if any)
    html_text = re.sub(r'<!--PAGE_START:\d+-->\s*', '', html_text)
    html_text = re.sub(r'\s*<!--PAGE_END:\d+-->', '', html_text)

    html_text = re.sub(r'<!--TABLE_START:\d+-->\s*', '', html_text)
    html_text = re.sub(r'\s*<!--TABLE_END:\d+-->', '', html_text)

    html_text = re.sub(r'<!--ICON_GROUP_START:[^>]+-->\s*', '', html_text)
    html_text = re.sub(r'\s*<!--ICON_GROUP_END:[^>]+-->', '', html_text)

    html_text = re.sub(r'<!--TEXT_BLOCK_START:[^>]+-->\s*', '', html_text)
    html_text = re.sub(r'\s*<!--TEXT_BLOCK_END:[^>]+-->', '', html_text)
    return html_text


def _find_matching_div_end(html_text: str, start_idx: int) -> int:
    """
    Finds the index (exclusive) of the matching closing </div> for the <div> starting at start_idx.
    Returns -1 if not found.
    """
    depth = 0
    pos = start_idx
    open_re = re.compile(r"<div\b", re.I)
    close_re = re.compile(r"</div\s*>", re.I)

    while pos < len(html_text):
        m_open = open_re.search(html_text, pos)
        m_close = close_re.search(html_text, pos)

        if not m_close:
            return -1

        if m_open and m_open.start() < m_close.start():
            depth += 1
            pos = m_open.end()
        else:
            depth -= 1
            pos = m_close.end()
            if depth == 0:
                return pos

    return -1


def _set_root_var(html_text: str, var_name: str, value_with_unit: str) -> str:
    """
    Updates (or inserts) a CSS variable inside :root { ... }.
    var_name must be like "page-padding" (without leading "--").
    """
    m = re.search(r":root\s*\{([^}]*)\}", html_text, re.S)
    if not m:
        return html_text

    block = m.group(1)
    # update if exists
    if re.search(rf"--{re.escape(var_name)}\s*:", block):
        def _repl(m2: re.Match) -> str:
            return m2.group(1) + value_with_unit + m2.group(3)
        block2 = re.sub(
            rf"(--{re.escape(var_name)}\s*:\s*)([^;]+)(;)",
            _repl,
            block,
        )
    else:
        block2 = block + f"\n        --{var_name}: {value_with_unit};"

    return html_text[: m.start(1)] + block2 + html_text[m.end(1) :]


def _ensure_layout_support(html_text: str) -> str:
    """
    One-time upgrade helper:
    - Adds CSS variables for spacing/sizes
    - Changes key CSS rules to reference variables
    - Converts inline img-box heights (style="height:XXXpx;") into classes (.img-h-XXX)
    """
    # 1) Ensure root vars exist
    defaults = {
        "page-padding": "20mm",
        "page-gap": "20px",
        "img-box-height": "220px",
        "img-box-margin-v": "10px",
        "highlight-margin-v": "15px",
        "table-margin-top": "10px",
        "table-cell-padding": "7px",
        "user-block-gap": "12px",
        "img-h-300": "300px",
        "img-h-250": "250px",
        "img-h-180": "180px",
        "img-h-150": "150px",
    }
    for k, v in defaults.items():
        html_text = _set_root_var(html_text, k, v)

    # 2) .page padding/margin-bottom
    html_text = re.sub(r"(\.page\s*\{[^}]*?)padding:\s*20mm\s*;", r"\1padding: var(--page-padding);", html_text, flags=re.S)
    html_text = re.sub(r"(\.page\s*\{[^}]*?)margin-bottom:\s*20px\s*;?", r"\1margin-bottom: var(--page-gap);", html_text, flags=re.S)

    # 3) .img-box height + margin
    html_text = re.sub(r"(\.img-box\s*\{[^}]*?)height:\s*220px\s*;", r"\1height: var(--img-box-height);", html_text, flags=re.S)
    # add margin if missing in first .img-box rule
    def _add_img_margin(match: re.Match) -> str:
        block = match.group(0)
        if "margin:" in block:
            return block
        return block[:-1] + " margin: var(--img-box-margin-v) 0; }"
    html_text = re.sub(r"\.img-box\s*\{[^}]*\}", _add_img_margin, html_text, count=1, flags=re.S)

    # 4) highlight-box margin
    html_text = re.sub(r"(\.highlight-box\s*\{[^}]*?)margin:\s*15px\s*0\s*;", r"\1margin: var(--highlight-margin-v) 0;", html_text, flags=re.S)

    # 5) table margins/padding
    html_text = re.sub(r"(proposal-table\s*\{[^}]*?)margin-top:\s*10px\s*;", r"\1margin-top: var(--table-margin-top);", html_text, flags=re.S)
    html_text = re.sub(r"(\.proposal-table\s+th\s*\{[^}]*?)padding:\s*7px\s*;", r"\1padding: var(--table-cell-padding);", html_text, flags=re.S)
    html_text = re.sub(r"(\.proposal-table\s+td\s*\{[^}]*?)padding:\s*7px\s*;", r"\1padding: var(--table-cell-padding);", html_text, flags=re.S)

    # 6) height classes + user text styles
    anchor = ".img-box img { width: 100%; height: 100%; object-fit: cover; }"
    if anchor in html_text and ".img-h-300" not in html_text:
        extra = """
    .img-h-300 { height: var(--img-h-300); }
    .img-h-250 { height: var(--img-h-250); }
    .img-h-180 { height: var(--img-h-180); }
    .img-h-150 { height: var(--img-h-150); }

    /* User text blocks (safe text-only editing area) */
    .user-text-block { margin-top: var(--user-block-gap); padding: 12px; border: 1px dashed #ccc; border-radius: 8px; background: #fff; }
    .user-text-title { font-weight: 700; margin-bottom: 6px; color: var(--primary-purple); }
    .user-text-block p { margin: 6px 0; }
    .user-text-block ul { margin: 6px 0 6px 18px; }
"""
        html_text = html_text.replace(anchor, anchor + extra)

    # 7) Convert inline height styles in img-box divs to classes
    # Examples:
    # <div class="img-box" style="height:300px;">
    # <div class="img-box" style="height: 150px; margin-top: 20px;">
    replacements = {
        300: "img-h-300",
        250: "img-h-250",
        180: "img-h-180",
        150: "img-h-150",
    }
    for px, cls in replacements.items():
        html_text = re.sub(
            rf'<div\s+class="img-box"\s+style="\s*height\s*:\s*{px}px\s*;\s*"\s*>',
            rf'<div class="img-box {cls}">',
            html_text,
        )
        # style that also contains margin-top etc -> strip style entirely (spacing will be controlled by variables)
        html_text = re.sub(
            rf'<div\s+class="img-box"\s+style="\s*height\s*:\s*{px}px\s*;\s*[^"]*"\s*>',
            rf'<div class="img-box {cls}">',
            html_text,
        )
    return html_text


def plain_text_to_safe_html(text: str) -> str:
    """
    Converts user plain-text to safe HTML fragments:
    - Blank lines separate paragraphs
    - Lines starting with "- " become bullet lists
    - Other lines within a paragraph are joined with <br>
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")

    blocks: List[str] = []
    buf: List[str] = []
    list_buf: List[str] = []

    def flush_paragraph() -> None:
        nonlocal buf
        if not buf:
            return
        escaped = [html.escape(x) for x in buf]
        blocks.append("<p>" + "<br>".join(escaped) + "</p>")
        buf = []

    def flush_list() -> None:
        nonlocal list_buf
        if not list_buf:
            return
        items = "".join([f"<li>{html.escape(x)}</li>" for x in list_buf])
        blocks.append("<ul>" + items + "</ul>")
        list_buf = []

    for raw in lines:
        line = raw.strip("\n")
        if line.strip() == "":
            flush_list()
            flush_paragraph()
            continue

        if line.lstrip().startswith("- "):
            flush_paragraph()
            item = line.lstrip()[2:].strip()
            list_buf.append(item)
        else:
            flush_list()
            buf.append(line)

    flush_list()
    flush_paragraph()

    if not blocks:
        return "<p></p>"
    return "\n".join(blocks)


def safe_html_to_plain_text(html_fragment: str) -> str:
    """
    Best-effort conversion from an HTML fragment back to plain text,
    so the user can re-edit text blocks.
    """
    s = html_fragment

    # list items
    s = re.sub(r"<\s*li[^>]*>", "- ", s, flags=re.I)
    s = re.sub(r"</\s*li\s*>", "\n", s, flags=re.I)
    s = re.sub(r"</\s*ul\s*>", "\n", s, flags=re.I)

    # paragraphs and line breaks
    s = re.sub(r"<\s*br\s*/?\s*>", "\n", s, flags=re.I)
    s = re.sub(r"</\s*p\s*>", "\n\n", s, flags=re.I)

    # strip remaining tags
    s = re.sub(r"<[^>]+>", "", s)
    s = html.unescape(s)

    # normalize whitespace
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


# -----------------------------
# Marker insertion helpers
# -----------------------------
def ensure_page_markers(html_text: str) -> str:
    if "<!--PAGE_START:" in html_text:
        return html_text

    starts = [m.start() for m in RAW_PAGE_START_RE.finditer(html_text)]
    if not starts:
        return html_text

    out = []
    last_pos = 0
    page_no = 1

    for start in starts:
        # copy content up to the page start
        out.append(html_text[last_pos:start])

        div_end = _find_matching_div_end(html_text, start)
        if div_end == -1:
            # can't safely wrap, bail out
            return html_text

        page_block = html_text[start:div_end]
        wrapped = f"<!--PAGE_START:{page_no}-->\n{page_block}\n<!--PAGE_END:{page_no}-->"
        out.append(wrapped)

        last_pos = div_end
        page_no += 1

    out.append(html_text[last_pos:])
    return "".join(out)


def ensure_table_markers(html_text: str) -> str:
    if "<!--TABLE_START:" in html_text:
        return html_text

    tables = list(RAW_TABLE_RE.finditer(html_text))
    if not tables:
        return html_text

    out = []
    last_pos = 0
    table_no = 1

    for m in tables:
        out.append(html_text[last_pos:m.start()])
        table_html = m.group(0)
        wrapped = f"<!--TABLE_START:{table_no}-->\n{table_html}\n<!--TABLE_END:{table_no}-->"
        out.append(wrapped)
        last_pos = m.end()
        table_no += 1

    out.append(html_text[last_pos:])
    return "".join(out)


def ensure_icon_markers(html_text: str) -> str:
    if "<!--ICON_GROUP_START:" in html_text:
        return html_text

    new_html = html_text

    # 1) process steps container (unique grid with repeat(4, 1fr))
    key = "grid-template-columns: repeat(4"
    idx = new_html.find(key)
    if idx != -1:
        div_start = new_html.rfind("<div", 0, idx)
        div_end = _find_matching_div_end(new_html, div_start) if div_start != -1 else -1
        if div_start != -1 and div_end != -1:
            block = new_html[div_start:div_end]
            wrapped = (
                "<!--ICON_GROUP_START:process_steps-->\n"
                + block
                + "\n<!--ICON_GROUP_END:process_steps-->"
            )
            new_html = new_html[:div_start] + wrapped + new_html[div_end:]

    # 2) centers list container (ul with fa-hospital-user icons)
    # heuristic: find first occurrence of "fa-hospital-user" then wrap the nearest <ul>...</ul>
    key2 = "fa-hospital-user"
    idx2 = new_html.find(key2)
    if idx2 != -1:
        ul_start = new_html.rfind("<ul", 0, idx2)
        ul_end = new_html.find("</ul>", idx2)
        if ul_start != -1 and ul_end != -1:
            ul_end += len("</ul>")
            block = new_html[ul_start:ul_end]
            wrapped = (
                "<!--ICON_GROUP_START:centers_list-->\n"
                + block
                + "\n<!--ICON_GROUP_END:centers_list-->"
            )
            new_html = new_html[:ul_start] + wrapped + new_html[ul_end:]

    return new_html


# -----------------------------
# Document model
# -----------------------------
@dataclass
class TemplateDocument:
    prefix: str
    pages: List[str]
    suffix: str

    @staticmethod
    def from_html(html_text: str) -> "TemplateDocument":
        matches = list(PAGE_BLOCK_RE.finditer(html_text))
        if not matches:
            html2 = ensure_page_markers(html_text)
            matches = list(PAGE_BLOCK_RE.finditer(html2))
            if not matches:
                return TemplateDocument(prefix=html_text, pages=[], suffix="")
            html_text = html2
            matches = list(PAGE_BLOCK_RE.finditer(html_text))

        first = matches[0]
        last = matches[-1]
        prefix = html_text[: first.start()]
        suffix = html_text[last.end() :]
        pages = [m.group(2).strip() for m in matches]
        return TemplateDocument(prefix=prefix, pages=pages, suffix=suffix)

    def _renumber_page_footers(self) -> None:
        new_pages: List[str] = []
        for i, page in enumerate(self.pages, start=1):
            page2 = re.sub(r"(>Page\s*)\d+(\s*<)", rf"\g<1>{i}\2", page)
            new_pages.append(page2)
        self.pages = new_pages

    def to_html(self) -> str:
        parts = [self.prefix]
        for i, page in enumerate(self.pages, start=1):
            parts.append(f"<!--PAGE_START:{i}-->\n{page}\n<!--PAGE_END:{i}-->")
        parts.append(self.suffix)
        return "".join(parts)


# -----------------------------
# Icon group helpers
# -----------------------------
def _extract_fa_items_from_div_grid(block_html: str) -> List[Tuple[str, str]]:
    """
    For process steps: returns list of (icon_class, label_text)
    """
    items: List[Tuple[str, str]] = []
    # each item: <div class="feature-item"> ... <i class="fas fa-..."></i> ... <h4>...</h4>
    for m in re.finditer(r'<div[^>]*class="feature-item"[^>]*>(.*?)</div>', block_html, re.S | re.I):
        chunk = m.group(1)
        icon_m = re.search(r'<i[^>]*class="([^"]*fa-[^"]*)"[^>]*>', chunk, re.I)
        label_m = re.search(r"<h4[^>]*>\s*([^<]+)\s*</h4>", chunk, re.I)
        icon_cls = icon_m.group(1).strip() if icon_m else "fas fa-circle"
        label = label_m.group(1).strip() if label_m else ""
        items.append((icon_cls, label))
    return items


def _build_div_grid_from_fa_items(items: List[Tuple[str, str]]) -> str:
    """
    Builds a feature-grid for process steps.
    """
    card_tpl = (
        '        <div class="feature-item">\n'
        '            <i class="{icon}"></i>\n'
        "            <h4>{label}</h4>\n"
        "        </div>\n"
    )
    inner = "".join([card_tpl.format(icon=html.escape(icon), label=html.escape(label)) for icon, label in items])
    return (
        '<div class="feature-grid" style="grid-template-columns: repeat(4, 1fr);">\n'
        + inner
        + "</div>"
    )


def _extract_li_items_from_ul(block_html: str) -> List[Tuple[str, str]]:
    """
    For centers list: returns list of (icon_class, label_text)
    """
    items: List[Tuple[str, str]] = []
    for m in re.finditer(r"<li[^>]*>(.*?)</li>", block_html, re.S | re.I):
        li = m.group(1)
        icon_m = re.search(r'<i[^>]*class="([^"]*fa-[^"]*)"[^>]*>', li, re.I)
        text = re.sub(r"<[^>]+>", "", li)
        text = html.unescape(text).strip()
        icon_cls = icon_m.group(1).strip() if icon_m else "fas fa-hospital-user"
        items.append((icon_cls, text))
    return items


def _build_ul_from_li_items(items: List[Tuple[str, str]]) -> str:
    li_html = []
    for icon, label in items:
        li_html.append(
            '            <li><i class="{icon}"></i> {label}</li>\n'.format(
                icon=html.escape(icon), label=html.escape(label)
            )
        )
    ul = (
        "        <ul style=\"list-style: none; padding-left: 0; font-size: 11pt; font-weight: bold;\">\n"
        + "".join(li_html)
        + "        </ul>"
    )
    return ul


# -----------------------------
# Main Engine
# -----------------------------
class ProposalEngine:
    """
    The "functional core" used by the GUI.
    """

    def __init__(self, base_dir: str):
        self.base_dir = base_dir
        self.assets_dir = os.path.join(self.base_dir, "proposal_assets")
        self.images_dir = os.path.join(self.assets_dir, "images")
        self.originals_dir = os.path.join(self.images_dir, "originals")
        os.makedirs(self.originals_dir, exist_ok=True)

        self.settings_path = os.path.join(self.assets_dir, "proposal_settings.json")

        # distributed template (next to .py files)
        self.distributed_template_path = os.path.join(self.base_dir, "proposal_template.html")
        # editable template (in assets)
        self.template_path = os.path.join(self.assets_dir, "proposal_template.html")

        # Image placeholders
        self.image_map: Dict[str, Dict[str, str]] = {
            "병원 전경": {"placeholder": "placeholder_hospital_view.jpg", "path": ""},
            "인증마크 모음": {"placeholder": "placeholder_cert_mark.jpg", "path": ""},
            "검진센터 내부": {"placeholder": "placeholder_center_interior.jpg", "path": ""},
            "MRI 장비": {"placeholder": "placeholder_mri.jpg", "path": ""},
            "CT 장비": {"placeholder": "placeholder_ct.jpg", "path": ""},
            "모바일 예약시스템": {"placeholder": "placeholder_mobile_system.jpg", "path": ""},
            "출장검진 버스": {"placeholder": "placeholder_bus.jpg", "path": ""},
            "검진 진행 모습": {"placeholder": "placeholder_exam_progress.jpg", "path": ""},
        }

        # store originals (by local file name) so we can re-resize if layout changes
        self.image_originals: Dict[str, str] = {}

        # Layout settings (numeric)
        self.layout_settings: Dict[str, int] = {
            "page_padding_mm": 20,
            "page_gap_px": 20,
            "img_default_height_px": 220,
            "img_margin_v_px": 10,
            "highlight_margin_v_px": 15,
            "table_margin_top_px": 10,
            "table_cell_padding_px": 7,
            "user_block_gap_px": 12,
            "img_h_300_px": 300,
            "img_h_250_px": 250,
            "img_h_180_px": 180,
            "img_h_150_px": 150,
        }

        self.image_target_sizes: Dict[str, Tuple[int, int]] = {}
        self._recompute_image_target_sizes()

        self.page_enabled: List[bool] = []

        self._ensure_template_file()
        self._load_settings()

    # -----------------------------
    # Layout
    # -----------------------------
    def _recompute_image_target_sizes(self) -> None:
        # Widths are chosen to roughly match 1-column and 2-column layouts in the template.
        # Heights are controlled by layout_settings (px).
        h300 = int(self.layout_settings.get("img_h_300_px", 300))
        h250 = int(self.layout_settings.get("img_h_250_px", 250))
        h180 = int(self.layout_settings.get("img_h_180_px", 180))
        h150 = int(self.layout_settings.get("img_h_150_px", 150))
        hdef = int(self.layout_settings.get("img_default_height_px", 220))

        self.image_target_sizes = {
            "병원 전경": (1000, h300),
            "인증마크 모음": (490, h150),
            "검진센터 내부": (1000, hdef),
            "MRI 장비": (490, h180),
            "CT 장비": (490, h180),
            "모바일 예약시스템": (1000, h250),
            "출장검진 버스": (490, h150),
            "검진 진행 모습": (1000, h150),
        }

    def get_layout_settings(self) -> Dict[str, int]:
        return dict(self.layout_settings)

    def set_layout_settings(self, new_settings: Dict[str, int]) -> None:
        # sanitize (keep int, minimums)
        def _clamp(name: str, v: int, lo: int, hi: int) -> int:
            try:
                vv = int(v)
            except Exception:
                vv = int(self.layout_settings.get(name, lo))
            return max(lo, min(hi, vv))

        self.layout_settings["page_padding_mm"] = _clamp("page_padding_mm", new_settings.get("page_padding_mm", 20), 5, 40)
        self.layout_settings["page_gap_px"] = _clamp("page_gap_px", new_settings.get("page_gap_px", 20), 0, 80)
        self.layout_settings["img_default_height_px"] = _clamp("img_default_height_px", new_settings.get("img_default_height_px", 220), 80, 600)
        self.layout_settings["img_margin_v_px"] = _clamp("img_margin_v_px", new_settings.get("img_margin_v_px", 10), 0, 80)
        self.layout_settings["highlight_margin_v_px"] = _clamp("highlight_margin_v_px", new_settings.get("highlight_margin_v_px", 15), 0, 80)
        self.layout_settings["table_margin_top_px"] = _clamp("table_margin_top_px", new_settings.get("table_margin_top_px", 10), 0, 80)
        self.layout_settings["table_cell_padding_px"] = _clamp("table_cell_padding_px", new_settings.get("table_cell_padding_px", 7), 2, 20)
        self.layout_settings["user_block_gap_px"] = _clamp("user_block_gap_px", new_settings.get("user_block_gap_px", 12), 0, 80)

        self.layout_settings["img_h_300_px"] = _clamp("img_h_300_px", new_settings.get("img_h_300_px", 300), 80, 800)
        self.layout_settings["img_h_250_px"] = _clamp("img_h_250_px", new_settings.get("img_h_250_px", 250), 80, 800)
        self.layout_settings["img_h_180_px"] = _clamp("img_h_180_px", new_settings.get("img_h_180_px", 180), 80, 800)
        self.layout_settings["img_h_150_px"] = _clamp("img_h_150_px", new_settings.get("img_h_150_px", 150), 80, 800)

        self._recompute_image_target_sizes()
        self._apply_layout_settings_to_template()
        self._rebuild_all_resized_images()
        self._save_settings()

    def _apply_layout_settings_to_template(self) -> None:
        html_text = self.load_template_html()
        html_text = _ensure_layout_support(html_text)

        html_text = _set_root_var(html_text, "page-padding", f"{self.layout_settings['page_padding_mm']}mm")
        html_text = _set_root_var(html_text, "page-gap", f"{self.layout_settings['page_gap_px']}px")
        html_text = _set_root_var(html_text, "img-box-height", f"{self.layout_settings['img_default_height_px']}px")
        html_text = _set_root_var(html_text, "img-box-margin-v", f"{self.layout_settings['img_margin_v_px']}px")
        html_text = _set_root_var(html_text, "highlight-margin-v", f"{self.layout_settings['highlight_margin_v_px']}px")
        html_text = _set_root_var(html_text, "table-margin-top", f"{self.layout_settings['table_margin_top_px']}px")
        html_text = _set_root_var(html_text, "table-cell-padding", f"{self.layout_settings['table_cell_padding_px']}px")
        html_text = _set_root_var(html_text, "user-block-gap", f"{self.layout_settings['user_block_gap_px']}px")

        html_text = _set_root_var(html_text, "img-h-300", f"{self.layout_settings['img_h_300_px']}px")
        html_text = _set_root_var(html_text, "img-h-250", f"{self.layout_settings['img_h_250_px']}px")
        html_text = _set_root_var(html_text, "img-h-180", f"{self.layout_settings['img_h_180_px']}px")
        html_text = _set_root_var(html_text, "img-h-150", f"{self.layout_settings['img_h_150_px']}px")

        self.save_template_html(html_text)

    # -----------------------------
    # Template file handling
    # -----------------------------
    def _ensure_template_file(self) -> None:
        os.makedirs(self.assets_dir, exist_ok=True)

        if os.path.exists(self.template_path):
            html_text = _safe_read_text(self.template_path)
            upgraded = ensure_page_markers(html_text)
            upgraded = ensure_table_markers(upgraded)
            upgraded = ensure_icon_markers(upgraded)
            upgraded = _ensure_layout_support(upgraded)
            if upgraded != html_text:
                _safe_write_text(self.template_path, upgraded)
            return

        # First run: copy distributed template into assets and upgrade
        if not os.path.exists(self.distributed_template_path):
            raise FileNotFoundError("proposal_template.html not found next to the program.")
        html_text = _safe_read_text(self.distributed_template_path)
        html_text = ensure_page_markers(html_text)
        html_text = ensure_table_markers(html_text)
        html_text = ensure_icon_markers(html_text)
        html_text = _ensure_layout_support(html_text)
        _safe_write_text(self.template_path, html_text)

    def load_template_html(self) -> str:
        return _safe_read_text(self.template_path)

    def save_template_html(self, html_text: str) -> None:
        _safe_write_text(self.template_path, html_text)

    def get_document(self) -> TemplateDocument:
        html_text = self.load_template_html()
        return TemplateDocument.from_html(html_text)

    def save_document(self, doc: TemplateDocument) -> None:
        doc._renumber_page_footers()
        self.save_template_html(doc.to_html())

    # -----------------------------
    # Settings
    # -----------------------------
    def _load_settings(self) -> None:
        os.makedirs(self.assets_dir, exist_ok=True)

        if not os.path.exists(self.settings_path):
            doc = self.get_document()
            self.page_enabled = [True] * len(doc.pages)
            self._apply_layout_settings_to_template()
            self._save_settings()
            return

        with open(self.settings_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # layout
        layout = data.get("layout", {})
        if isinstance(layout, dict):
            for k in list(self.layout_settings.keys()):
                if k in layout:
                    try:
                        self.layout_settings[k] = int(layout[k])
                    except Exception:
                        pass
        self._recompute_image_target_sizes()
        self._apply_layout_settings_to_template()

        # images: prefer originals
        self.image_originals = {}
        images_original = data.get("images_original", {})
        if isinstance(images_original, dict):
            for key, fname in images_original.items():
                if key in self.image_map and fname:
                    orig_path = os.path.join(self.originals_dir, str(fname))
                    if os.path.exists(orig_path):
                        self.image_originals[key] = str(fname)
                        resized_path = self._ensure_resized_from_original(key, orig_path)
                        if os.path.exists(resized_path):
                            self.image_map[key]["path"] = resized_path

        # backward compatibility: older settings that stored resized only
        if not self.image_originals:
            images = data.get("images", {})
            if isinstance(images, dict):
                for key, fname in images.items():
                    if key in self.image_map and fname:
                        resized_path = os.path.join(self.images_dir, str(fname))
                        if os.path.exists(resized_path):
                            # treat resized as original (copy to originals) so future re-resize is possible
                            orig_fname = os.path.basename(resized_path)
                            orig_path = os.path.join(self.originals_dir, orig_fname)
                            if not os.path.exists(orig_path):
                                shutil.copy2(resized_path, orig_path)
                            self.image_originals[key] = orig_fname
                            self.image_map[key]["path"] = resized_path

        # pages enabled
        enabled = data.get("page_enabled", None)
        doc = self.get_document()
        if isinstance(enabled, list):
            self.page_enabled = [bool(x) for x in enabled[: len(doc.pages)]]
            if len(self.page_enabled) < len(doc.pages):
                self.page_enabled += [True] * (len(doc.pages) - len(self.page_enabled))
        else:
            self.page_enabled = [True] * len(doc.pages)

        self._save_settings()

    def _save_settings(self) -> None:
        data_to_save: Dict[str, object] = {
            "page_enabled": self.page_enabled,
            "layout": self.layout_settings,
            "images_original": self.image_originals,
            "images": {},
        }

        # keep resized basenames too (for convenience)
        images_resized: Dict[str, str] = {}
        for key, meta in self.image_map.items():
            p = meta.get("path", "")
            if p and os.path.exists(p):
                images_resized[key] = os.path.basename(p)
        data_to_save["images"] = images_resized

        with open(self.settings_path, "w", encoding="utf-8") as f:
            json.dump(data_to_save, f, ensure_ascii=False, indent=2)

    def save_settings(self) -> None:
        self._save_settings()

    # -----------------------------
    # Images
    # -----------------------------
    @staticmethod
    def _resize_cover(img: Image.Image, target_w: int, target_h: int) -> Image.Image:
        """
        Center-crop and resize to exactly (target_w, target_h).
        """
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")

        w, h = img.size
        if w == 0 or h == 0:
            return img

        target_ratio = target_w / target_h
        src_ratio = w / h

        if src_ratio > target_ratio:
            # too wide -> crop width
            new_w = int(h * target_ratio)
            left = (w - new_w) // 2
            box = (left, 0, left + new_w, h)
        else:
            # too tall -> crop height
            new_h = int(w / target_ratio)
            top = (h - new_h) // 2
            box = (0, top, w, top + new_h)

        cropped = img.crop(box)
        resized = cropped.resize((target_w, target_h), Image.LANCZOS)
        return resized

    def _ensure_resized_from_original(self, image_key: str, original_path: str) -> str:
        target = self.image_target_sizes.get(image_key, (1000, 1000))
        h = _hash_file(original_path)
        safe_key = re.sub(r"[^0-9a-zA-Z가-힣_-]+", "_", image_key).strip("_")
        out_name = f"{safe_key}_{h}_{target[0]}x{target[1]}.jpg"
        out_path = os.path.join(self.images_dir, out_name)

        if os.path.exists(out_path):
            return out_path

        with Image.open(original_path) as img:
            img2 = self._resize_cover(img, target[0], target[1])
            img2.save(out_path, format="JPEG", quality=90, optimize=True)
        return out_path

    def _rebuild_all_resized_images(self) -> None:
        """
        Recreate resized images from originals (if available) after layout changes.
        """
        for key, orig_fname in list(self.image_originals.items()):
            if key not in self.image_map:
                continue
            orig_path = os.path.join(self.originals_dir, orig_fname)
            if not os.path.exists(orig_path):
                continue
            resized = self._ensure_resized_from_original(key, orig_path)
            if os.path.exists(resized):
                self.image_map[key]["path"] = resized

    def copy_resize_to_local(self, image_key: str, src_path: str) -> str:
        """
        Copy user-selected image into originals dir, then generate resized file into images dir.
        Returns the resized local path.
        """
        if not os.path.exists(src_path):
            raise FileNotFoundError(src_path)

        # 1) copy original into program folder
        h = _hash_file(src_path)
        safe_key = re.sub(r"[^0-9a-zA-Z가-힣_-]+", "_", image_key).strip("_")
        ext = os.path.splitext(src_path)[1].lower()
        if ext not in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff"]:
            ext = ".bin"
        orig_name = f"{safe_key}_{h}{ext}"
        orig_path = os.path.join(self.originals_dir, orig_name)
        if not os.path.exists(orig_path):
            shutil.copy2(src_path, orig_path)

        # 2) build resized from original
        resized_path = self._ensure_resized_from_original(image_key, orig_path)

        # remember original file name
        self.image_originals[image_key] = orig_name
        return resized_path

    @staticmethod
    def image_file_to_data_uri(path: str) -> str:
        with open(path, "rb") as f:
            raw = f.read()
        encoded = base64.b64encode(raw).decode("utf-8")
        return f"data:image/jpeg;base64,{encoded}"

    # -----------------------------
    # Pages
    # -----------------------------
    def get_pages(self) -> List[str]:
        doc = self.get_document()
        return doc.pages

    def set_page_enabled(self, idx: int, enabled: bool) -> None:
        doc = self.get_document()
        if not self.page_enabled or len(self.page_enabled) != len(doc.pages):
            self.page_enabled = [True] * len(doc.pages)
        if 0 <= idx < len(self.page_enabled):
            self.page_enabled[idx] = bool(enabled)
            self._save_settings()

    def move_page(self, idx: int, direction: int) -> None:
        doc = self.get_document()
        j = idx + direction
        if idx < 0 or j < 0 or idx >= len(doc.pages) or j >= len(doc.pages):
            return
        doc.pages[idx], doc.pages[j] = doc.pages[j], doc.pages[idx]
        self.save_document(doc)
        # move enabled flags too
        if self.page_enabled and len(self.page_enabled) == len(doc.pages):
            self.page_enabled[idx], self.page_enabled[j] = self.page_enabled[j], self.page_enabled[idx]
            self._save_settings()

    def duplicate_page(self, idx: int) -> None:
        doc = self.get_document()
        if idx < 0 or idx >= len(doc.pages):
            return
        doc.pages.insert(idx + 1, doc.pages[idx])
        self.save_document(doc)
        if not self.page_enabled or len(self.page_enabled) != len(doc.pages) - 1:
            self.page_enabled = [True] * (len(doc.pages) - 1)
        self.page_enabled.insert(idx + 1, True)
        self._save_settings()

    def add_new_page(self) -> None:
        doc = self.get_document()
        skeleton = (
            '<div class="page">\n'
            '  <div class="page-header"><h2>새 페이지</h2></div>\n'
            '  <div class="section-title">내용</div>\n'
            '  <div class="user-text-block" data-block-id="newpage_block1">\n'
            '    <div class="user-text-title">추가 텍스트</div>\n'
            '    <!--TEXT_BLOCK_START:newpage_block1-->\n'
            '    <p></p>\n'
            '    <!--TEXT_BLOCK_END:newpage_block1-->\n'
            '  </div>\n'
            '  <div class="page-footer"><span>NEW</span><span>Page 1</span></div>\n'
            '</div>'
        )
        doc.pages.append(skeleton)
        self.save_document(doc)
        if not self.page_enabled or len(self.page_enabled) != len(doc.pages) - 1:
            self.page_enabled = [True] * (len(doc.pages) - 1)
        self.page_enabled.append(True)
        self._save_settings()

    def delete_page(self, idx: int) -> None:
        doc = self.get_document()
        if idx < 0 or idx >= len(doc.pages):
            return
        doc.pages.pop(idx)
        self.save_document(doc)
        if self.page_enabled and idx < len(self.page_enabled):
            self.page_enabled.pop(idx)
        self._save_settings()

    def update_page_html(self, idx: int, new_html: str) -> None:
        # kept for backward compatibility (advanced usage)
        doc = self.get_document()
        if 0 <= idx < len(doc.pages):
            doc.pages[idx] = new_html
            self.save_document(doc)

    # -----------------------------
    # Text blocks (safe text-only editing)
    # -----------------------------
    def list_text_blocks(self, page_idx: int) -> List[Dict[str, str]]:
        doc = self.get_document()
        if page_idx < 0 or page_idx >= len(doc.pages):
            return []

        page_html = doc.pages[page_idx]
        blocks: List[Dict[str, str]] = []

        # find block wrappers
        for m in re.finditer(r'<div\s+class="user-text-block"[^>]*data-block-id="([^"]+)"[^>]*>', page_html, flags=re.I):
            block_id = m.group(1).strip()
            block_start = m.start()
            block_end = _find_matching_div_end(page_html, block_start)
            if block_end == -1:
                continue
            block_html = page_html[block_start:block_end]

            title_m = re.search(r'<div\s+class="user-text-title"[^>]*>\s*(.*?)\s*</div>', block_html, flags=re.I | re.S)
            title = html.unescape(re.sub(r"<[^>]+>", "", title_m.group(1)).strip()) if title_m else block_id

            body_m = re.search(rf'<!--TEXT_BLOCK_START:{re.escape(block_id)}-->\s*(.*?)\s*<!--TEXT_BLOCK_END:{re.escape(block_id)}-->', block_html, flags=re.S)
            body_html = body_m.group(1).strip() if body_m else ""
            plain = safe_html_to_plain_text(body_html)

            blocks.append({"id": block_id, "title": title, "text": plain})

        return blocks

    def add_text_block(self, page_idx: int, title: str = "추가 텍스트") -> str:
        doc = self.get_document()
        if page_idx < 0 or page_idx >= len(doc.pages):
            return ""

        page_no = page_idx + 1
        existing = self.list_text_blocks(page_idx)
        n = 1
        while True:
            block_id = f"page{page_no}_block{n}"
            if not any(b["id"] == block_id for b in existing):
                break
            n += 1

        block_html = (
            f'\n  <div class="user-text-block" data-block-id="{block_id}">\n'
            f'    <div class="user-text-title">{html.escape(title)}</div>\n'
            f'    <!--TEXT_BLOCK_START:{block_id}-->\n'
            f'    <p></p>\n'
            f'    <!--TEXT_BLOCK_END:{block_id}-->\n'
            f'  </div>\n'
        )

        page_html = doc.pages[page_idx]
        insert_at = page_html.lower().find('<div class="page-footer"')
        if insert_at != -1:
            page_html = page_html[:insert_at] + block_html + page_html[insert_at:]
        else:
            # best-effort: before last </div>
            last_close = page_html.lower().rfind("</div>")
            if last_close != -1:
                page_html = page_html[:last_close] + block_html + page_html[last_close:]
            else:
                page_html += block_html

        doc.pages[page_idx] = page_html
        self.save_document(doc)
        return block_id

    def delete_text_block(self, page_idx: int, block_id: str) -> None:
        doc = self.get_document()
        if page_idx < 0 or page_idx >= len(doc.pages):
            return

        page_html = doc.pages[page_idx]
        # locate the wrapper
        m = re.search(rf'<div\s+class="user-text-block"[^>]*data-block-id="{re.escape(block_id)}"[^>]*>', page_html, flags=re.I)
        if not m:
            return
        start = m.start()
        end = _find_matching_div_end(page_html, start)
        if end == -1:
            return

        page_html = page_html[:start] + page_html[end:]
        doc.pages[page_idx] = page_html
        self.save_document(doc)

    def save_text_block(self, page_idx: int, block_id: str, title: str, plain_text: str) -> None:
        doc = self.get_document()
        if page_idx < 0 or page_idx >= len(doc.pages):
            return

        page_html = doc.pages[page_idx]

        # 1) update title inside wrapper
        title_escaped = html.escape(title)

        # isolate the wrapper html first
        m = re.search(rf'<div\s+class="user-text-block"[^>]*data-block-id="{re.escape(block_id)}"[^>]*>', page_html, flags=re.I)
        if not m:
            return
        start = m.start()
        end = _find_matching_div_end(page_html, start)
        if end == -1:
            return

        wrapper = page_html[start:end]
        wrapper = re.sub(
            r'(<div\s+class="user-text-title"[^>]*>\s*)(.*?)(\s*</div>)',
            rf"\1{title_escaped}\3",
            wrapper,
            count=1,
            flags=re.S | re.I,
        )

        # 2) update body between TEXT markers
        new_body = plain_text_to_safe_html(plain_text)
        wrapper = re.sub(
            rf'(<!--TEXT_BLOCK_START:{re.escape(block_id)}-->\s*)(.*?)(\s*<!--TEXT_BLOCK_END:{re.escape(block_id)}-->)',
            rf"\1{new_body}\3",
            wrapper,
            flags=re.S,
        )

        page_html = page_html[:start] + wrapper + page_html[end:]
        doc.pages[page_idx] = page_html
        self.save_document(doc)

    # -----------------------------
    # Tables
    # -----------------------------
    def list_tables(self) -> List[int]:
        html_text = self.load_template_html()
        matches = list(TABLE_BLOCK_RE.finditer(html_text))
        return [int(m.group(1)) for m in matches]

    def get_table_html(self, table_no: int) -> str:
        html_text = self.load_template_html()
        m = TABLE_BLOCK_RE.search(html_text)
        for mm in TABLE_BLOCK_RE.finditer(html_text):
            if int(mm.group(1)) == table_no:
                return mm.group(2).strip()
        return ""

    def set_table_html(self, table_no: int, new_table_html: str) -> None:
        html_text = self.load_template_html()

        def repl(match: re.Match) -> str:
            if int(match.group(1)) != table_no:
                return match.group(0)
            return f"<!--TABLE_START:{table_no}-->\n{new_table_html}\n<!--TABLE_END:{table_no}-->"

        html_text2 = TABLE_BLOCK_RE.sub(repl, html_text)
        self.save_template_html(html_text2)

    def add_empty_row_to_table(self, table_no: int) -> None:
        table_html = self.get_table_html(table_no)
        if not table_html:
            return

        # naive: add before </table>
        # find column count from first row
        cols = 3
        m = re.search(r"<tr[^>]*>\s*(.*?)\s*</tr>", table_html, re.S | re.I)
        if m:
            cols = len(re.findall(r"<t[hd][^>]*>", m.group(1), re.I)) or cols

        row = "<tr>" + "".join(["<td></td>" for _ in range(cols)]) + "</tr>"
        table_html2 = re.sub(r"</table\s*>", row + "\n</table>", table_html, flags=re.I)
        self.set_table_html(table_no, table_html2)

    def clear_table(self, table_no: int) -> None:
        table_html = self.get_table_html(table_no)
        if not table_html:
            return
        table_html2 = re.sub(r"<tr[^>]*>.*?</tr>", "", table_html, flags=re.S | re.I)
        self.set_table_html(table_no, table_html2)

    def rescan_tables(self) -> None:
        html_text = self.load_template_html()
        # Remove existing table markers first
        html_text = re.sub(r"<!--TABLE_START:\d+-->\s*", "", html_text)
        html_text = re.sub(r"\s*<!--TABLE_END:\d+-->", "", html_text)
        html_text = ensure_table_markers(html_text)
        self.save_template_html(html_text)

    # -----------------------------
    # Icon groups
    # -----------------------------
    def get_icon_group_html(self, group_key: str) -> str:
        html_text = self.load_template_html()
        for m in ICON_GROUP_RE.finditer(html_text):
            if m.group(1) == group_key:
                return m.group(2).strip()
        return ""

    def set_icon_group_html(self, group_key: str, new_html: str) -> None:
        html_text = self.load_template_html()

        def repl(match: re.Match) -> str:
            if match.group(1) != group_key:
                return match.group(0)
            return f"<!--ICON_GROUP_START:{group_key}-->\n{new_html}\n<!--ICON_GROUP_END:{group_key}-->"

        html_text2 = ICON_GROUP_RE.sub(repl, html_text)
        self.save_template_html(html_text2)

    def get_process_steps(self) -> List[Dict[str, str]]:
        block = self.get_icon_group_html("process_steps")
        items = _extract_fa_items_from_div_grid(block)
        return [{"icon": icon, "label": label} for icon, label in items]

    def save_process_steps(self, items: List[Dict[str, str]]) -> None:
        tuples = [(it.get("icon", "fas fa-circle"), it.get("label", "")) for it in items]
        html_block = _build_div_grid_from_fa_items(tuples)
        self.set_icon_group_html("process_steps", html_block)

    def get_centers_items(self) -> List[Dict[str, str]]:
        block = self.get_icon_group_html("centers_list")
        items = _extract_li_items_from_ul(block)
        return [{"icon": icon, "label": label} for icon, label in items]

    def save_centers_items(self, items: List[Dict[str, str]]) -> None:
        tuples = [(it.get("icon", "fas fa-hospital-user"), it.get("label", "")) for it in items]
        ul = _build_ul_from_li_items(tuples)
        self.set_icon_group_html("centers_list", ul)

    def rescan_icon_groups(self) -> None:
        html_text = self.load_template_html()
        # Remove existing icon markers first
        html_text = re.sub(r"<!--ICON_GROUP_START:[^>]+-->\s*", "", html_text)
        html_text = re.sub(r"\s*<!--ICON_GROUP_END:[^>]+-->", "", html_text)
        html_text = ensure_icon_markers(html_text)
        self.save_template_html(html_text)

    # -----------------------------
    # Final HTML generation (apply replacements + embed images)
    # -----------------------------
    def build_output_html(
        self,
        recipient: str,
        proposer: str,
        tel: str,
        primary_color: str,
        accent_color: str,
    ) -> str:
        html_text = self.load_template_html()

        # Apply page selection
        doc = TemplateDocument.from_html(html_text)
        if len(self.page_enabled) != len(doc.pages):
            self.page_enabled = [True] * len(doc.pages)
            self._save_settings()

        selected_pages = [p for p, en in zip(doc.pages, self.page_enabled) if en]
        doc.pages = selected_pages
        html_text = doc.to_html()

        # Basic info replacements
        html_text = re.sub(r"(<strong>수신\s*:\s*</strong>\s*)([^<]+)", rf"\1{recipient}", html_text)
        html_text = re.sub(r"(<strong>제안\s*:\s*</strong>\s*)([^<]+)", rf"\1{proposer}", html_text)
        html_text = re.sub(r"(Tel\.\s*)([0-9\s\-]+)", rf"\1{tel}", html_text)

        # Colors
        html_text = re.sub(r"--primary-purple:\s*#[0-9A-Fa-f]{6}\s*;", f"--primary-purple: {primary_color};", html_text)
        html_text = re.sub(r"--accent-gold:\s*#[0-9A-Fa-f]{6}\s*;", f"--accent-gold: {accent_color};", html_text)

        # Images: replace placeholders with embedded Base64
        for key, meta in self.image_map.items():
            local_path = meta.get("path", "")
            placeholder = meta.get("placeholder", "")
            if local_path and os.path.exists(local_path) and placeholder:
                data_uri = self.image_file_to_data_uri(local_path)
                html_text = html_text.replace(f'src="{placeholder}"', f'src="{data_uri}"')

        # Remove editor markers from final output for cleanliness
        html_text = _remove_all_markers(html_text)
        return html_text
