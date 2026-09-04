# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import tempfile
import urllib.request
from pathlib import Path
from typing import Any, Dict, List

from PIL import Image, ImageDraw, ImageFont
import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from flyer_data import normalize_flyer_data

THEMES = {
    "기업형": {"bg": "#F3F6FA", "primary": "#183B66", "accent": "#2E74B5", "soft": "#EAF2FA", "text": "#1F2937"},
    "여름": {"bg": "#DDF5FF", "primary": "#0577B8", "accent": "#00A8E8", "soft": "#EDF9FF", "text": "#12324A"},
    "봄": {"bg": "#FFF0F5", "primary": "#B84371", "accent": "#EE7DA8", "soft": "#FFF8FB", "text": "#3D2430"},
    "업그레이드": {"bg": "#FFF0CF", "primary": "#B7442D", "accent": "#F06A3A", "soft": "#FFF9ED", "text": "#472B25"},
    "가정의달": {"bg": "#DDEEFF", "primary": "#225AA8", "accent": "#FF7F6C", "soft": "#F7FBFF", "text": "#20324B"},
    "미니멀": {"bg": "#F4F4F4", "primary": "#30343B", "accent": "#666D75", "soft": "#FFFFFF", "text": "#202428"},
}

FONT_URLS = {
    "regular": "https://raw.githubusercontent.com/orioncactus/pretendard/main/packages/pretendard/dist/public/static/Pretendard-Regular.otf",
    "medium": "https://raw.githubusercontent.com/orioncactus/pretendard/main/packages/pretendard/dist/public/static/Pretendard-Medium.otf",
    "semibold": "https://raw.githubusercontent.com/orioncactus/pretendard/main/packages/pretendard/dist/public/static/Pretendard-SemiBold.otf",
    "bold": "https://raw.githubusercontent.com/orioncactus/pretendard/main/packages/pretendard/dist/public/static/Pretendard-Bold.otf",
}


def _font_dir() -> Path:
    p = Path(tempfile.gettempdir()) / "nk_flyer_fonts"
    p.mkdir(parents=True, exist_ok=True)
    return p


def ensure_fonts() -> Dict[str, str]:
    out = {}
    for weight, url in FONT_URLS.items():
        path = _font_dir() / f"Pretendard-{weight}.otf"
        if not path.exists() or path.stat().st_size < 100000:
            urllib.request.urlretrieve(url, path)
        out[weight] = str(path)
    return out


def _register_fonts(paths: Dict[str, str]) -> Dict[str, str]:
    names = {k: f"NK-Pretendard-{k}" for k in paths}
    registered = set(pdfmetrics.getRegisteredFontNames())
    for k, name in names.items():
        if name not in registered:
            pdfmetrics.registerFont(TTFont(name, paths[k]))
    return names


def _hex(value: str):
    value = value.lstrip("#")
    return tuple(int(value[i:i+2], 16) for i in (0, 2, 4))


def _fit_pdf(c, text, x, y, width, font, size, min_size=5.5):
    text = str(text or "")
    s = size
    while s > min_size and pdfmetrics.stringWidth(text, font, s) > width:
        s -= 0.25
    c.setFont(font, s)
    c.drawString(x, y, text)


def _wrap_pdf(text: str, font: str, size: float, width: float) -> List[str]:
    text = str(text or "").strip()
    if not text:
        return []
    words = text.split()
    if len(words) <= 1:
        words, joiner = list(text), ""
    else:
        joiner = " "
    lines, current = [], ""
    for w in words:
        candidate = w if not current else current + joiner + w
        if pdfmetrics.stringWidth(candidate, font, size) <= width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = w
    if current:
        lines.append(current)
    return lines


def _qr(url: str):
    if not str(url or "").strip():
        return None
    q = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=7, border=2)
    q.add_data(url.strip())
    q.make(fit=True)
    return q.make_image(fill_color="black", back_color="white").convert("RGB")


class FlyerEngine:
    def __init__(self):
        self.font_paths = ensure_fonts()
        self.fonts = _register_fonts(self.font_paths)

    def render_pdf(self, data: Dict[str, Any], background_bytes: bytes | None = None) -> bytes:
        d = normalize_flyer_data(data)
        theme = THEMES.get(d.get("theme"), THEMES["기업형"])
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=A4)
        W, H = A4
        margin = 28

        c.setFillColor(theme["bg"])
        c.rect(0, 0, W, H, stroke=0, fill=1)
        if background_bytes:
            try:
                im = Image.open(io.BytesIO(background_bytes)).convert("RGB")
                tmp = io.BytesIO(); im.save(tmp, format="JPEG", quality=94); tmp.seek(0)
                c.drawImage(ImageReader(tmp), 0, 0, W, H, mask="auto")
            except Exception:
                pass

        primary, accent, text, soft = theme["primary"], theme["accent"], theme["text"], theme["soft"]
        c.setFillColor(primary)
        title = d.get("title") or "건강검진 안내문"
        size = 26
        while size > 17 and pdfmetrics.stringWidth(title, self.fonts["bold"], size) > W - 2 * margin:
            size -= 1
        c.setFont(self.fonts["bold"], size)
        c.drawString(margin, H - 54, title)

        meta = "  |  ".join(x for x in [
            d.get("target", ""),
            f"검진기간 {d.get('period','')}" if d.get("period") else "",
            f"접수기간 {d.get('application_period','')}" if d.get("application_period") else "",
        ] if x)
        c.setFillColor(text)
        _fit_pdf(c, meta, margin, H - 76, W - 2 * margin, self.fonts["medium"], 9.5, 7)

        y = H - 92
        c.setFillColor("#FFFFFF")
        c.roundRect(margin, y - 54, W - 2 * margin, 50, 8, stroke=0, fill=1)
        c.setStrokeColor(accent); c.setLineWidth(1.2)
        c.roundRect(margin, y - 54, W - 2 * margin, 50, 8, stroke=1, fill=0)
        c.setFillColor(primary); c.setFont(self.fonts["semibold"], 11)
        c.drawString(margin + 12, y - 20, d.get("event_title") or "EVENT")
        c.setFillColor(text)
        for i, line in enumerate(d.get("event_lines", [])[:2]):
            _fit_pdf(c, "• " + line, margin + 110, y - 20 - i * 15, W - 2 * margin - 125, self.fonts["medium"], 8.8, 6.8)

        table_top = y - 68
        packages = d.get("packages", [])
        row_h = 32
        header_h = 28
        max_rows = min(len(packages), 8)
        table_h = header_h + row_h * max_rows
        c.setFillColor("#FFFFFF")
        c.roundRect(margin, table_top - table_h, W - 2 * margin, table_h, 8, stroke=0, fill=1)
        c.setFillColor(primary)
        c.roundRect(margin, table_top - header_h, W - 2 * margin, header_h, 8, stroke=0, fill=1)
        c.setFillColor("#FFFFFF"); c.setFont(self.fonts["bold"], 10.5)
        c.drawString(margin + 12, table_top - 19, "플랜명")
        c.drawString(margin + 132, table_top - 19, "세부항목")
        c.drawRightString(W - margin - 12, table_top - 19, "금액")

        yy = table_top - header_h
        for idx, p in enumerate(packages[:max_rows]):
            yy -= row_h
            if idx % 2 == 0:
                c.setFillColor(soft); c.rect(margin, yy, W - 2 * margin, row_h, stroke=0, fill=1)
            c.setFillColor(primary); c.setFont(self.fonts["bold"], 9.2)
            _fit_pdf(c, p.get("name", ""), margin + 12, yy + 11, 110, self.fonts["bold"], 9.2, 6.8)
            c.setFillColor(text)
            _fit_pdf(c, p.get("detail", ""), margin + 132, yy + 11, W - 2 * margin - 235, self.fonts["medium"], 8.2, 6.2)
            c.setFont(self.fonts["bold"], 9.2)
            c.drawRightString(W - margin - 12, yy + 11, p.get("price") or p.get("male_price", ""))

        common_top = table_top - table_h - 10
        c.setFillColor("#FFFFFF"); c.roundRect(margin, common_top - 60, W - 2*margin, 60, 7, stroke=0, fill=1)
        c.setFillColor(primary); c.setFont(self.fonts["bold"], 10); c.drawString(margin + 10, common_top - 16, "공통항목")
        c.setFillColor(text); c.setFont(self.fonts["regular"], 6.6)
        for i, line in enumerate(_wrap_pdf(d.get("common_items", ""), self.fonts["regular"], 6.6, W - 2*margin - 20)[:5]):
            c.drawString(margin + 10, common_top - 30 - i*9, line)

        group_top = common_top - 72
        group_bottom = 110
        total_h = max(110, group_top - group_bottom)
        gap = 8
        col_w = (W - 2*margin - 2*gap) / 3
        for gi, g in enumerate(("A", "B", "C")):
            x = margin + gi*(col_w+gap)
            c.setFillColor("#FFFFFF"); c.roundRect(x, group_bottom, col_w, total_h, 7, stroke=0, fill=1)
            c.setFillColor(primary); c.rect(x, group_top - 24, col_w, 24, stroke=0, fill=1)
            c.setFillColor("#FFFFFF"); c.setFont(self.fonts["bold"], 10.5); c.drawString(x+8, group_top-17, f"{g}그룹")
            c.setFillColor(text)
            items = d.get("groups", {}).get(g, [])
            fs = 6.0 if g == "A" else 5.8
            step = 11
            for i, item in enumerate(items[:22]):
                yy2 = group_top - 37 - i*step
                if yy2 < group_bottom + 8:
                    break
                _fit_pdf(c, item, x+7, yy2, col_w-14, self.fonts["regular"], fs, 4.8)

        footer_y = 18
        q = _qr(d.get("qr_url", ""))
        qr_w = 68
        if q:
            qbuf = io.BytesIO(); q.save(qbuf, format="PNG"); qbuf.seek(0)
            c.drawImage(ImageReader(qbuf), margin, footer_y, qr_w, qr_w, mask="auto")
            c.linkURL(d["qr_url"], (margin, footer_y, margin+qr_w, footer_y+qr_w), relative=0)
        tx = margin + (qr_w + 10 if q else 0)
        c.setFillColor(primary); c.setFont(self.fonts["bold"], 19)
        c.drawString(tx, footer_y + 39, f"검진문의 {d.get('phone','1833-9988')}")
        c.setFillColor(text); c.setFont(self.fonts["regular"], 6.7)
        for i, line in enumerate(d.get("notes", [])[:3]):
            _fit_pdf(c, "• " + line, tx, footer_y + 22 - i*10, W-margin-tx, self.fonts["regular"], 6.7, 5.4)

        c.showPage(); c.save(); return buf.getvalue()

    def render_png(self, data: Dict[str, Any], background_bytes: bytes | None = None) -> bytes:
        d = normalize_flyer_data(data)
        theme = THEMES.get(d.get("theme"), THEMES["기업형"])
        W, H = 1240, 1754
        if background_bytes:
            try:
                img = Image.open(io.BytesIO(background_bytes)).convert("RGB").resize((W, H))
            except Exception:
                img = Image.new("RGB", (W, H), _hex(theme["bg"]))
        else:
            img = Image.new("RGB", (W, H), _hex(theme["bg"]))
        dr = ImageDraw.Draw(img)
        F = {k: lambda s, p=v: ImageFont.truetype(p, s) for k, v in self.font_paths.items()}
        primary, accent, text, soft = map(_hex, [theme["primary"], theme["accent"], theme["text"], theme["soft"]])
        x0, x1 = 60, W-60
        dr.text((x0, 65), d.get("title") or "건강검진 안내문", font=F["bold"](54), fill=primary)
        meta = "  |  ".join(x for x in [d.get("target", ""), d.get("period", ""), d.get("application_period", "")] if x)
        dr.text((x0, 132), meta, font=F["medium"](21), fill=text)
        dr.rounded_rectangle((x0, 175, x1, 285), radius=18, fill=(255,255,255), outline=accent, width=3)
        dr.text((x0+22, 194), d.get("event_title") or "EVENT", font=F["semibold"](28), fill=primary)
        for i, line in enumerate(d.get("event_lines", [])[:2]):
            dr.text((x0+235, 194+i*38), "• "+line, font=F["medium"](20), fill=text)

        y = 320
        packages = d.get("packages", [])[:8]
        dr.rounded_rectangle((x0, y, x1, y+65+76*len(packages)), radius=18, fill=(255,255,255))
        dr.rounded_rectangle((x0, y, x1, y+65), radius=18, fill=primary)
        dr.text((x0+22, y+17), "플랜명", font=F["bold"](25), fill=(255,255,255))
        dr.text((x0+300, y+17), "세부항목", font=F["bold"](25), fill=(255,255,255))
        dr.text((x1-135, y+17), "금액", font=F["bold"](25), fill=(255,255,255))
        for i, p in enumerate(packages):
            yy = y+65+i*76
            if i % 2 == 0: dr.rectangle((x0, yy, x1, yy+76), fill=soft)
            dr.text((x0+22, yy+22), p.get("name", ""), font=F["bold"](22), fill=primary)
            dr.text((x0+300, yy+22), p.get("detail", ""), font=F["medium"](19), fill=text)
            dr.text((x1-125, yy+22), p.get("price") or p.get("male_price", ""), font=F["bold"](22), fill=text)

        y2 = y+65+76*len(packages)+24
        dr.rounded_rectangle((x0, y2, x1, y2+135), radius=16, fill=(255,255,255))
        dr.text((x0+18, y2+15), "공통항목", font=F["bold"](24), fill=primary)
        common = d.get("common_items", "")
        # 고정 폭 기준 단순 줄바꿈
        lines, cur = [], ""
        for token in common.split(" | "):
            cand = token if not cur else cur+" | "+token
            if dr.textlength(cand, font=F["regular"](16)) < x1-x0-40: cur = cand
            else:
                if cur: lines.append(cur)
                cur = token
        if cur: lines.append(cur)
        for i, line in enumerate(lines[:4]): dr.text((x0+18, y2+52+i*21), line, font=F["regular"](16), fill=text)

        gy = y2+160; gap=14; cw=(x1-x0-2*gap)//3; gh=390
        for gi, g in enumerate(("A","B","C")):
            gx=x0+gi*(cw+gap)
            dr.rounded_rectangle((gx, gy, gx+cw, gy+gh), radius=14, fill=(255,255,255))
            dr.rectangle((gx, gy, gx+cw, gy+55), fill=primary)
            dr.text((gx+15, gy+13), f"{g}그룹", font=F["bold"](25), fill=(255,255,255))
            for i, item in enumerate(d.get("groups",{}).get(g,[])[:20]):
                yy=gy+68+i*15
                if yy>gy+gh-20: break
                dr.text((gx+12, yy), item, font=F["regular"](13), fill=text)

        fy=H-135
        q=_qr(d.get("qr_url", ""))
        if q:
            q=q.resize((100,100)); img.paste(q,(x0,fy-5))
            tx=x0+125
        else: tx=x0
        dr.text((tx, fy+8), f"검진문의 {d.get('phone','1833-9988')}", font=F["bold"](34), fill=primary)
        out=io.BytesIO(); img.save(out, format="PNG"); return out.getvalue()
