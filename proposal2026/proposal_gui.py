# proposal_gui.py
# -*- coding: utf-8 -*-
"""
Proposal Maker GUI (cleaner UI + safer text editing)
- Optional modern theme via ttkbootstrap (recommended)
- Shows color swatches for primary/accent colors
- "Page content" is now edited as plain text blocks (converted to safe HTML)
- Layout (spacing/sizes) can be tuned via CSS variable controls
- Page/table/icon edit features remain available
"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk, filedialog, colorchooser, messagebox
from tkinter.scrolledtext import ScrolledText

from proposal_engine import ProposalEngine


class ProposalEditorApp:
    def __init__(self, root: tk.Tk, modern_theme: bool = False):
        self.root = root
        self.root.title("제안서 생성기 (UI/UX 개선 + 텍스트 안전 편집)")

        # slightly cleaner default paddings
        self.modern_theme = modern_theme
        self._setup_style()

        base_dir = os.path.dirname(os.path.abspath(__file__))
        self.engine = ProposalEngine(base_dir)

        # Basic info
        self.recipient_var = tk.StringVar(value="김포시청")
        self.proposer_var = tk.StringVar(value="뉴고려병원")
        self.tel_var = tk.StringVar(value="031-980-9114")

        self.primary_color = "#4A148C"
        self.accent_color = "#D4AF37"

        # Layout vars (loaded from engine)
        self.layout_vars = {k: tk.IntVar(value=v) for k, v in self.engine.get_layout_settings().items()}

        # -------------------- UI --------------------
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_basic = ttk.Frame(self.nb)
        self.tab_layout = ttk.Frame(self.nb)
        self.tab_images = ttk.Frame(self.nb)
        self.tab_pages = ttk.Frame(self.nb)
        self.tab_page_text = ttk.Frame(self.nb)
        self.tab_tables = ttk.Frame(self.nb)
        self.tab_icons = ttk.Frame(self.nb)
        self.tab_export = ttk.Frame(self.nb)

        self.nb.add(self.tab_basic, text="기본 정보")
        self.nb.add(self.tab_layout, text="레이아웃/공백")
        self.nb.add(self.tab_images, text="이미지")
        self.nb.add(self.tab_pages, text="페이지 관리")
        self.nb.add(self.tab_page_text, text="페이지 내용(텍스트)")
        self.nb.add(self.tab_tables, text="표 편집")
        self.nb.add(self.tab_icons, text="아이콘 편집")
        self.nb.add(self.tab_export, text="제안서 생성")

        self._build_basic_tab()
        self._build_layout_tab()
        self._build_images_tab()
        self._build_pages_tab()
        self._build_page_text_tab()
        self._build_tables_tab()
        self._build_icons_tab()
        self._build_export_tab()

    # -----------------------------
    # Styling / theming
    # -----------------------------
    def _setup_style(self) -> None:
        try:
            style = ttk.Style()
            # If not using ttkbootstrap, try a nicer built-in theme
            if not self.modern_theme:
                for theme in ("clam", "vista", "xpnative", "alt", "default"):
                    if theme in style.theme_names():
                        style.theme_use(theme)
                        break

            style.configure("TFrame", padding=6)
            style.configure("TLabelframe", padding=10)
            style.configure("TLabelframe.Label", font=("Arial", 10, "bold"))
            style.configure("TButton", padding=6)
            style.configure("TNotebook.Tab", padding=(12, 6))
        except Exception:
            pass

    # -----------------------------
    # Basic tab
    # -----------------------------
    def _build_basic_tab(self) -> None:
        frm = ttk.Frame(self.tab_basic)
        frm.pack(fill="x", padx=10, pady=10)

        info = ttk.Labelframe(frm, text="기본 정보")
        info.pack(fill="x", pady=5)

        ttk.Label(info, text="수신").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(info, textvariable=self.recipient_var, width=40).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(info, text="제안").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(info, textvariable=self.proposer_var, width=40).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(info, text="전화번호").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(info, textvariable=self.tel_var, width=40).grid(row=2, column=1, sticky="w", pady=4)

        colors = ttk.Labelframe(frm, text="컬러")
        colors.pack(fill="x", pady=10)

        self.btn_primary = ttk.Button(colors, text="메인 컬러 선택", command=self._pick_primary_color)
        self.btn_primary.grid(row=0, column=0, sticky="w", pady=4)

        self.primary_swatch = tk.Canvas(colors, width=28, height=18, highlightthickness=1, highlightbackground="#999")
        self.primary_swatch.grid(row=0, column=1, sticky="w", padx=8)
        self.lbl_primary = ttk.Label(colors, text=self.primary_color)
        self.lbl_primary.grid(row=0, column=2, sticky="w")

        self.btn_accent = ttk.Button(colors, text="포인트 컬러 선택", command=self._pick_accent_color)
        self.btn_accent.grid(row=1, column=0, sticky="w", pady=4)

        self.accent_swatch = tk.Canvas(colors, width=28, height=18, highlightthickness=1, highlightbackground="#999")
        self.accent_swatch.grid(row=1, column=1, sticky="w", padx=8)
        self.lbl_accent = ttk.Label(colors, text=self.accent_color)
        self.lbl_accent.grid(row=1, column=2, sticky="w")

        ttk.Label(colors, text="(참고) 색상은 '제안서 생성' 시 최종 HTML에 반영됩니다.").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=6
        )

        self._refresh_color_swatches()

    def _refresh_color_swatches(self) -> None:
        # primary
        self.primary_swatch.delete("all")
        self.primary_swatch.create_rectangle(0, 0, 28, 18, fill=self.primary_color, outline=self.primary_color)
        self.lbl_primary.config(text=self.primary_color)
        # accent
        self.accent_swatch.delete("all")
        self.accent_swatch.create_rectangle(0, 0, 28, 18, fill=self.accent_color, outline=self.accent_color)
        self.lbl_accent.config(text=self.accent_color)

    def _pick_primary_color(self) -> None:
        c = colorchooser.askcolor(title="메인 컬러 선택", initialcolor=self.primary_color)
        if c and c[1]:
            self.primary_color = c[1]
            self._refresh_color_swatches()

    def _pick_accent_color(self) -> None:
        c = colorchooser.askcolor(title="포인트 컬러 선택", initialcolor=self.accent_color)
        if c and c[1]:
            self.accent_color = c[1]
            self._refresh_color_swatches()

    # -----------------------------
    # Layout tab (spacing/sizes)
    # -----------------------------
    def _build_layout_tab(self) -> None:
        outer = ttk.Frame(self.tab_layout)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        lf = ttk.Labelframe(outer, text="공백/크기 조정 (CSS 변수)")
        lf.pack(fill="x", pady=6)

        row = 0
        def add_spin(label: str, key: str, unit: str, lo: int, hi: int):
            nonlocal row
            ttk.Label(lf, text=label).grid(row=row, column=0, sticky="w", pady=4)
            sp = ttk.Spinbox(lf, from_=lo, to=hi, textvariable=self.layout_vars[key], width=8)
            sp.grid(row=row, column=1, sticky="w", pady=4)
            ttk.Label(lf, text=unit).grid(row=row, column=2, sticky="w", padx=6)
            row += 1

        add_spin("페이지 안쪽 여백", "page_padding_mm", "mm", 5, 40)
        add_spin("페이지 간격(화면 보기용)", "page_gap_px", "px", 0, 80)
        add_spin("기본 이미지 박스 높이", "img_default_height_px", "px", 80, 600)
        add_spin("이미지 위/아래 공백", "img_margin_v_px", "px", 0, 80)
        add_spin("강조 박스 위/아래 공백", "highlight_margin_v_px", "px", 0, 80)
        add_spin("표 위쪽 공백", "table_margin_top_px", "px", 0, 80)
        add_spin("표 셀 패딩(행 높이 느낌)", "table_cell_padding_px", "px", 2, 20)
        add_spin("텍스트 블록 위쪽 공백", "user_block_gap_px", "px", 0, 80)

        ttk.Separator(lf).grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1

        ttk.Label(lf, text="특수 이미지 높이(템플릿의 300/250/180/150 영역)").grid(row=row, column=0, columnspan=3, sticky="w", pady=4)
        row += 1
        add_spin("높이 300 영역", "img_h_300_px", "px", 80, 800)
        add_spin("높이 250 영역", "img_h_250_px", "px", 80, 800)
        add_spin("높이 180 영역", "img_h_180_px", "px", 80, 800)
        add_spin("높이 150 영역", "img_h_150_px", "px", 80, 800)

        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=10)

        ttk.Button(btns, text="레이아웃 적용(템플릿 저장 + 이미지 재리사이즈)", command=self._apply_layout).pack(side="left")
        ttk.Label(
            btns,
            text="※ 이미지 높이를 바꾸면, 저장된 이미지가 원본 복사본 기준으로 다시 리사이즈됩니다.",
        ).pack(side="left", padx=10)

    def _apply_layout(self) -> None:
        new_settings = {k: v.get() for k, v in self.layout_vars.items()}
        try:
            self.engine.set_layout_settings(new_settings)
            messagebox.showinfo("완료", "레이아웃 설정이 반영되었습니다.\n(템플릿 저장 + 이미지 재리사이즈 완료)")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # -----------------------------
    # Images tab
    # -----------------------------
    def _build_images_tab(self) -> None:
        frm = ttk.Frame(self.tab_images)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(frm, text="이미지를 선택하면 프로그램 폴더(proposal_assets/images)에 복사 + 리사이즈되어 저장됩니다.").pack(anchor="w", pady=4)

        self.img_rows = []
        for key in self.engine.image_map.keys():
            row = ttk.Frame(frm)
            row.pack(fill="x", pady=3)

            ttk.Label(row, text=key, width=18).pack(side="left")
            lbl = ttk.Label(row, text="(선택 안됨)", foreground="#666")
            lbl.pack(side="left", padx=6)

            btn = ttk.Button(row, text="이미지 선택", command=lambda k=key, l=lbl: self._browse_image(k, l))
            btn.pack(side="right")

            self.img_rows.append((key, lbl))

        self._refresh_image_labels()

    def _refresh_image_labels(self) -> None:
        for key, lbl in self.img_rows:
            p = self.engine.image_map[key].get("path", "")
            if p and os.path.exists(p):
                lbl.config(text=os.path.basename(p), foreground="#1a73e8")
            else:
                lbl.config(text="(선택 안됨)", foreground="#666")

    def _browse_image(self, key: str, label_widget: ttk.Label) -> None:
        filename = filedialog.askopenfilename(
            title=f"{key} 이미지 선택",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.gif;*.tif;*.tiff"), ("All Files", "*.*")]
        )
        if not filename:
            return

        try:
            local_path = self.engine.copy_resize_to_local(key, filename)
            self.engine.image_map[key]["path"] = local_path
            self.engine.save_settings()
            label_widget.config(text=os.path.basename(local_path), foreground="#1a73e8")
            messagebox.showinfo("완료", "이미지가 복사/리사이즈되어 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # -----------------------------
    # Pages tab
    # -----------------------------
    def _build_pages_tab(self) -> None:
        outer = ttk.Frame(self.tab_pages)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        left = ttk.Frame(outer)
        left.pack(side="left", fill="both", expand=True)

        ttk.Label(left, text="페이지 목록").pack(anchor="w")
        self.page_list = tk.Listbox(left, height=18)
        self.page_list.pack(fill="both", expand=True, pady=6)
        self.page_list.bind("<<ListboxSelect>>", lambda e: self._on_page_select())

        right = ttk.Frame(outer)
        right.pack(side="left", fill="y", padx=10)

        self.chk_var = tk.BooleanVar(value=True)
        self.chk_enabled = ttk.Checkbutton(right, text="페이지 포함", variable=self.chk_var, command=self._toggle_page_enabled)
        self.chk_enabled.pack(anchor="w", pady=3)

        ttk.Button(right, text="위로", command=lambda: self._move_page(-1)).pack(fill="x", pady=3)
        ttk.Button(right, text="아래로", command=lambda: self._move_page(1)).pack(fill="x", pady=3)
        ttk.Separator(right).pack(fill="x", pady=10)
        ttk.Button(right, text="복제", command=self._duplicate_page).pack(fill="x", pady=3)
        ttk.Button(right, text="새 페이지 추가", command=self._add_page).pack(fill="x", pady=3)
        ttk.Button(right, text="삭제", command=self._delete_page).pack(fill="x", pady=3)

        self._refresh_pages_list()

    def _refresh_pages_list(self) -> None:
        self.page_list.delete(0, "end")
        pages = self.engine.get_pages()
        for i in range(len(pages)):
            en = True
            if self.engine.page_enabled and i < len(self.engine.page_enabled):
                en = self.engine.page_enabled[i]
            prefix = "✓ " if en else "✗ "
            self.page_list.insert("end", f"{prefix}페이지 {i+1}")

        if pages:
            self.page_list.selection_set(0)
            self._on_page_select()

    def _selected_page_index(self) -> int:
        sel = self.page_list.curselection()
        return int(sel[0]) if sel else -1

    def _on_page_select(self) -> None:
        idx = self._selected_page_index()
        if idx < 0:
            return
        enabled = True
        if self.engine.page_enabled and idx < len(self.engine.page_enabled):
            enabled = self.engine.page_enabled[idx]
        self.chk_var.set(enabled)

        # Page text tab may not be built yet during startup
        if hasattr(self, "page_combo"):
            self._refresh_text_blocks_for_page(idx)

    def _toggle_page_enabled(self) -> None:
        idx = self._selected_page_index()
        if idx < 0:
            return
        self.engine.set_page_enabled(idx, self.chk_var.get())
        self._refresh_pages_list()

    def _move_page(self, direction: int) -> None:
        idx = self._selected_page_index()
        if idx < 0:
            return
        self.engine.move_page(idx, direction)
        self._refresh_pages_list()
        new_idx = max(0, min(idx + direction, self.page_list.size() - 1))
        self.page_list.selection_clear(0, "end")
        self.page_list.selection_set(new_idx)
        self._on_page_select()

    def _duplicate_page(self) -> None:
        idx = self._selected_page_index()
        if idx < 0:
            return
        self.engine.duplicate_page(idx)
        self._refresh_pages_list()

    def _add_page(self) -> None:
        self.engine.add_new_page()
        self._refresh_pages_list()

    def _delete_page(self) -> None:
        idx = self._selected_page_index()
        if idx < 0:
            return
        if messagebox.askyesno("확인", "선택한 페이지를 삭제하시겠습니까?"):
            self.engine.delete_page(idx)
            self._refresh_pages_list()

    # -----------------------------
    # Page text blocks tab (text-only)
    # -----------------------------
    def _build_page_text_tab(self) -> None:
        outer = ttk.Frame(self.tab_page_text)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        top = ttk.Frame(outer)
        top.pack(fill="x")

        ttk.Label(top, text="페이지 선택:").pack(side="left")
        self.page_combo = ttk.Combobox(top, state="readonly", width=18)
        self.page_combo.pack(side="left", padx=6)
        self.page_combo.bind("<<ComboboxSelected>>", lambda e: self._on_page_combo_changed())

        ttk.Label(top, text="텍스트 블록:").pack(side="left", padx=(20, 0))
        self.block_combo = ttk.Combobox(top, state="readonly", width=24)
        self.block_combo.pack(side="left", padx=6)
        self.block_combo.bind("<<ComboboxSelected>>", lambda e: self._on_block_combo_changed())

        ttk.Button(top, text="블록 추가", command=self._add_text_block).pack(side="left", padx=8)
        ttk.Button(top, text="블록 삭제", command=self._delete_text_block).pack(side="left", padx=4)

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True, pady=10)

        left = ttk.Frame(body)
        left.pack(side="left", fill="y")

        ttk.Label(left, text="블록 제목").pack(anchor="w")
        self.block_title_var = tk.StringVar(value="")
        ttk.Entry(left, textvariable=self.block_title_var, width=28).pack(anchor="w", pady=4)

        ttk.Button(left, text="저장", command=self._save_text_block).pack(anchor="w", pady=6)

        ttk.Label(
            left,
            text="입력 규칙:\n- 빈 줄: 문단 구분\n- '- '로 시작: 글머리표\n- 같은 문단의 줄바꿈: <br>로 처리",
            justify="left",
        ).pack(anchor="w", pady=8)

        right = ttk.Frame(body)
        right.pack(side="left", fill="both", expand=True, padx=12)

        ttk.Label(right, text="내용(텍스트만 입력)").pack(anchor="w")
        self.block_text = ScrolledText(right, height=18)
        self.block_text.pack(fill="both", expand=True, pady=4)

        self._refresh_page_combo()

    def _refresh_page_combo(self) -> None:
        pages = self.engine.get_pages()
        values = [f"페이지 {i+1}" for i in range(len(pages))]
        self.page_combo["values"] = values
        if values:
            self.page_combo.current(0)
            self._refresh_text_blocks_for_page(0)

    def _on_page_combo_changed(self) -> None:
        idx = self.page_combo.current()
        if idx < 0:
            return
        self._refresh_text_blocks_for_page(idx)

    def _refresh_text_blocks_for_page(self, page_idx: int) -> None:
        # keep combobox synced
        pages = self.engine.get_pages()
        values = [f"페이지 {i+1}" for i in range(len(pages))]
        self.page_combo["values"] = values
        if values and (self.page_combo.current() != page_idx):
            self.page_combo.current(page_idx)

        blocks = self.engine.list_text_blocks(page_idx)
        if not blocks:
            self.block_combo["values"] = ["(블록 없음)"]
            self.block_combo.current(0)
            self.block_title_var.set("")
            self.block_text.delete("1.0", "end")
            return

        self._current_blocks_cache = blocks  # store for lookup
        bvals = [f"{b['id']} | {b['title']}" for b in blocks]
        self.block_combo["values"] = bvals
        self.block_combo.current(0)
        self._load_block_to_editor(0)

    def _on_block_combo_changed(self) -> None:
        idx = self.block_combo.current()
        if idx < 0:
            return
        self._load_block_to_editor(idx)

    def _load_block_to_editor(self, idx: int) -> None:
        blocks = getattr(self, "_current_blocks_cache", [])
        if not blocks or idx < 0 or idx >= len(blocks):
            return
        b = blocks[idx]
        self.block_title_var.set(b.get("title", ""))
        self.block_text.delete("1.0", "end")
        self.block_text.insert("1.0", b.get("text", ""))

    def _add_text_block(self) -> None:
        page_idx = self.page_combo.current()
        if page_idx < 0:
            return
        block_id = self.engine.add_text_block(page_idx, title="추가 텍스트")
        if not block_id:
            messagebox.showerror("오류", "블록 추가에 실패했습니다.")
            return
        messagebox.showinfo("완료", "텍스트 블록이 추가되었습니다.")
        self._refresh_text_blocks_for_page(page_idx)

    def _delete_text_block(self) -> None:
        page_idx = self.page_combo.current()
        blocks = getattr(self, "_current_blocks_cache", [])
        idx = self.block_combo.current()
        if page_idx < 0 or not blocks or idx < 0 or idx >= len(blocks):
            return
        block_id = blocks[idx]["id"]
        if messagebox.askyesno("확인", "이 텍스트 블록을 삭제하시겠습니까?"):
            self.engine.delete_text_block(page_idx, block_id)
            self._refresh_text_blocks_for_page(page_idx)

    def _save_text_block(self) -> None:
        page_idx = self.page_combo.current()
        blocks = getattr(self, "_current_blocks_cache", [])
        idx = self.block_combo.current()
        if page_idx < 0 or not blocks or idx < 0 or idx >= len(blocks):
            messagebox.showerror("오류", "저장할 블록을 선택하세요.")
            return
        block_id = blocks[idx]["id"]
        title = self.block_title_var.get().strip() or "텍스트"
        text = self.block_text.get("1.0", "end").rstrip()
        try:
            self.engine.save_text_block(page_idx, block_id, title, text)
            messagebox.showinfo("완료", "텍스트 블록이 저장되었습니다.")
            self._refresh_text_blocks_for_page(page_idx)
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # -----------------------------
    # Tables tab
    # -----------------------------
    def _build_tables_tab(self) -> None:
        outer = ttk.Frame(self.tab_tables)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        top = ttk.Frame(outer)
        top.pack(fill="x")

        ttk.Label(top, text="표 선택:").pack(side="left")
        self.table_combo = ttk.Combobox(top, state="readonly", width=20)
        self.table_combo.pack(side="left", padx=6)
        self.table_combo.bind("<<ComboboxSelected>>", lambda e: self._load_table())

        ttk.Button(top, text="빈 행 추가", command=self._table_add_row).pack(side="left", padx=8)
        ttk.Button(top, text="표 내용 비우기", command=self._table_clear).pack(side="left", padx=4)
        ttk.Button(top, text="테이블 재스캔", command=self._table_rescan).pack(side="right")

        self.table_editor = ScrolledText(outer, height=18)
        self.table_editor.pack(fill="both", expand=True, pady=10)

        btns = ttk.Frame(outer)
        btns.pack(fill="x")

        ttk.Button(btns, text="저장", command=self._save_table).pack(side="left")

        self._refresh_tables()

    def _refresh_tables(self) -> None:
        tables = self.engine.list_tables()
        vals = [f"TABLE {n}" for n in tables]
        self.table_combo["values"] = vals
        if vals:
            self.table_combo.current(0)
            self._load_table()
        else:
            self.table_editor.delete("1.0", "end")

    def _selected_table_no(self) -> int:
        sel = self.table_combo.get().strip()
        m = sel.replace("TABLE", "").strip()
        try:
            return int(m)
        except Exception:
            return -1

    def _load_table(self) -> None:
        n = self._selected_table_no()
        if n <= 0:
            return
        html = self.engine.get_table_html(n)
        self.table_editor.delete("1.0", "end")
        self.table_editor.insert("1.0", html)

    def _save_table(self) -> None:
        n = self._selected_table_no()
        if n <= 0:
            return
        html = self.table_editor.get("1.0", "end").rstrip()
        self.engine.set_table_html(n, html)
        messagebox.showinfo("완료", "표가 저장되었습니다.")

    def _table_add_row(self) -> None:
        n = self._selected_table_no()
        if n <= 0:
            return
        self.engine.add_empty_row_to_table(n)
        self._load_table()

    def _table_clear(self) -> None:
        n = self._selected_table_no()
        if n <= 0:
            return
        if messagebox.askyesno("확인", "표 내용을 비우시겠습니까?"):
            self.engine.clear_table(n)
            self._load_table()

    def _table_rescan(self) -> None:
        self.engine.rescan_tables()
        self._refresh_tables()
        messagebox.showinfo("완료", "테이블을 다시 감지했습니다.")

    # -----------------------------
    # Icons tab
    # -----------------------------
    def _build_icons_tab(self) -> None:
        outer = ttk.Frame(self.tab_icons)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        top = ttk.Frame(outer)
        top.pack(fill="x")
        ttk.Button(top, text="아이콘 그룹 재스캔", command=self._icons_rescan).pack(side="right")

        # process steps
        lf1 = ttk.Labelframe(outer, text="검진 프로세스(아이콘+텍스트)")
        lf1.pack(fill="both", expand=True, pady=6)

        self.proc_list = tk.Listbox(lf1, height=6)
        self.proc_list.pack(side="left", fill="both", expand=True, padx=6, pady=6)

        ctrl = ttk.Frame(lf1)
        ctrl.pack(side="left", fill="y", padx=6, pady=6)

        ttk.Button(ctrl, text="추가", command=self._proc_add).pack(fill="x", pady=2)
        ttk.Button(ctrl, text="삭제", command=self._proc_del).pack(fill="x", pady=2)
        ttk.Button(ctrl, text="위", command=lambda: self._proc_move(-1)).pack(fill="x", pady=2)
        ttk.Button(ctrl, text="아래", command=lambda: self._proc_move(1)).pack(fill="x", pady=2)
        ttk.Button(ctrl, text="저장", command=self._proc_save).pack(fill="x", pady=8)

        self.proc_icon_var = tk.StringVar(value="fas fa-clipboard-list")
        self.proc_label_var = tk.StringVar(value="접수")
        ttk.Label(ctrl, text="아이콘 class").pack(anchor="w", pady=(10, 0))
        ttk.Entry(ctrl, textvariable=self.proc_icon_var, width=24).pack(anchor="w")
        ttk.Label(ctrl, text="라벨").pack(anchor="w", pady=(6, 0))
        ttk.Entry(ctrl, textvariable=self.proc_label_var, width=24).pack(anchor="w")

        # centers list
        lf2 = ttk.Labelframe(outer, text="진료과목/센터 목록(아이콘+텍스트)")
        lf2.pack(fill="both", expand=True, pady=6)

        self.cent_list = tk.Listbox(lf2, height=6)
        self.cent_list.pack(side="left", fill="both", expand=True, padx=6, pady=6)

        ctrl2 = ttk.Frame(lf2)
        ctrl2.pack(side="left", fill="y", padx=6, pady=6)

        ttk.Button(ctrl2, text="추가", command=self._cent_add).pack(fill="x", pady=2)
        ttk.Button(ctrl2, text="삭제", command=self._cent_del).pack(fill="x", pady=2)
        ttk.Button(ctrl2, text="위", command=lambda: self._cent_move(-1)).pack(fill="x", pady=2)
        ttk.Button(ctrl2, text="아래", command=lambda: self._cent_move(1)).pack(fill="x", pady=2)
        ttk.Button(ctrl2, text="저장", command=self._cent_save).pack(fill="x", pady=8)

        self.cent_icon_var = tk.StringVar(value="fas fa-hospital-user")
        self.cent_label_var = tk.StringVar(value="내과센터")
        ttk.Label(ctrl2, text="아이콘 class").pack(anchor="w", pady=(10, 0))
        ttk.Entry(ctrl2, textvariable=self.cent_icon_var, width=24).pack(anchor="w")
        ttk.Label(ctrl2, text="라벨").pack(anchor="w", pady=(6, 0))
        ttk.Entry(ctrl2, textvariable=self.cent_label_var, width=24).pack(anchor="w")

        self._refresh_icons_lists()

    def _icons_rescan(self) -> None:
        self.engine.rescan_icon_groups()
        self._refresh_icons_lists()
        messagebox.showinfo("완료", "아이콘 그룹을 다시 감지했습니다.")

    # --- process steps ---
    def _refresh_icons_lists(self) -> None:
        self.proc_items = self.engine.get_process_steps()
        self.proc_list.delete(0, "end")
        for it in self.proc_items:
            self.proc_list.insert("end", f"{it['icon']} | {it['label']}")

        self.cent_items = self.engine.get_centers_items()
        self.cent_list.delete(0, "end")
        for it in self.cent_items:
            self.cent_list.insert("end", f"{it['icon']} | {it['label']}")

    def _proc_selected(self) -> int:
        sel = self.proc_list.curselection()
        return int(sel[0]) if sel else -1

    def _proc_add(self) -> None:
        self.proc_items.append({"icon": self.proc_icon_var.get().strip(), "label": self.proc_label_var.get().strip()})
        self._refresh_icons_lists()

    def _proc_del(self) -> None:
        i = self._proc_selected()
        if i < 0:
            return
        self.proc_items.pop(i)
        self._refresh_icons_lists()

    def _proc_move(self, d: int) -> None:
        i = self._proc_selected()
        if i < 0:
            return
        j = i + d
        if j < 0 or j >= len(self.proc_items):
            return
        self.proc_items[i], self.proc_items[j] = self.proc_items[j], self.proc_items[i]
        self._refresh_icons_lists()
        self.proc_list.selection_set(j)

    def _proc_save(self) -> None:
        self.engine.save_process_steps(self.proc_items)
        messagebox.showinfo("완료", "검진 프로세스가 저장되었습니다.")

    # --- centers ---
    def _cent_selected(self) -> int:
        sel = self.cent_list.curselection()
        return int(sel[0]) if sel else -1

    def _cent_add(self) -> None:
        self.cent_items.append({"icon": self.cent_icon_var.get().strip(), "label": self.cent_label_var.get().strip()})
        self._refresh_icons_lists()

    def _cent_del(self) -> None:
        i = self._cent_selected()
        if i < 0:
            return
        self.cent_items.pop(i)
        self._refresh_icons_lists()

    def _cent_move(self, d: int) -> None:
        i = self._cent_selected()
        if i < 0:
            return
        j = i + d
        if j < 0 or j >= len(self.cent_items):
            return
        self.cent_items[i], self.cent_items[j] = self.cent_items[j], self.cent_items[i]
        self._refresh_icons_lists()
        self.cent_list.selection_set(j)

    def _cent_save(self) -> None:
        self.engine.save_centers_items(self.cent_items)
        messagebox.showinfo("완료", "센터 목록이 저장되었습니다.")

    # -----------------------------
    # Export tab
    # -----------------------------
    def _build_export_tab(self) -> None:
        frm = ttk.Frame(self.tab_export)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(frm, text="최종 HTML 생성 (이미지 Base64 포함)").pack(anchor="w", pady=4)

        btn = ttk.Button(frm, text="HTML 생성", command=self._export_html)
        btn.pack(anchor="w", pady=6)

        self.export_path_var = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.export_path_var, foreground="#1a73e8").pack(anchor="w", pady=4)

    def _export_html(self) -> None:
        out = filedialog.asksaveasfilename(
            title="저장할 HTML 파일 선택",
            defaultextension=".html",
            filetypes=[("HTML", "*.html")]
        )
        if not out:
            return

        try:
            html_text = self.engine.build_output_html(
                recipient=self.recipient_var.get().strip(),
                proposer=self.proposer_var.get().strip(),
                tel=self.tel_var.get().strip(),
                primary_color=self.primary_color,
                accent_color=self.accent_color,
            )
            with open(out, "w", encoding="utf-8") as f:
                f.write(html_text)
            self.export_path_var.set(f"저장 완료: {out}")
            messagebox.showinfo("완료", "HTML이 생성되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", str(e))


def launch():
    """
    Optional modern UI:
      pip install ttkbootstrap
    """
    try:
        import ttkbootstrap as tb  # type: ignore
        root = tb.Window(themename="flatly")
        app = ProposalEditorApp(root, modern_theme=True)
    except Exception:
        root = tk.Tk()
        app = ProposalEditorApp(root, modern_theme=False)
    root.mainloop()
