"""
G-PDF - 문서 변환 & PDF 압축 통합 프로그램
한글(hwp/hwpx), Word, Excel, PowerPoint → PDF 변환 + G-Fit 압축 + 원본 첨부
"""

import json
import os
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pdf_compressor import (
    compress_pdf, format_size, get_compression_method,
    is_pdf_valid,
)
from doc_converter import (
    is_convertible, convert_to_pdf, get_engine_status_text,
)

# 드래그&드롭 지원
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False


class GPDFApp:
    """G-PDF 메인 애플리케이션"""

    HEADER_BG = "#1a3a5c"
    BODY_BG = "#f0f2f5"
    PANEL_BG = "#e8eaed"
    CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".gpdf_config.json")

    # 지원 확장자
    _PDF_EXTS = {".pdf"}
    _DOC_EXTS = {".hwp", ".hwpx", ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"}
    _ALL_EXTS = _PDF_EXTS | _DOC_EXTS

    # 압축 프리셋
    COMPRESS_MODES = {
        "original":  {"label": "원본 출력",        "quality": 0,  "dpi": 0,   "lossless": False, "desc": "변환만, 압축 없음"},
        "lossless":  {"label": "무손실 최적화",    "quality": 95, "dpi": 300, "lossless": True,  "desc": "품질 유지, 구조만 최적화"},
        "standard":  {"label": "표준 압축 (권장)",  "quality": 65, "dpi": 150, "lossless": False, "desc": "품질/용량 균형"},
        "strong":    {"label": "강한 압축",        "quality": 40, "dpi": 120, "lossless": False, "desc": "용량 최소화 우선"},
    }

    def __init__(self):
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title("G-PDF")
        self.root.geometry("780x750")
        self.root.minsize(700, 650)
        self.root.configure(bg=self.BODY_BG)

        config = self._load_config()
        self.last_dir = config.get("last_dir", "")
        self.output_dir = config.get("output_dir", "")
        # 최근 저장 폴더 (최대 4개)
        self.recent_output_dirs = [
            d for d in config.get("recent_output_dirs", [])
            if os.path.isdir(d)
        ][:4]

        self.file_list = []       # 변환/압축할 메인 파일
        self.attach_list = []     # PDF에 첨부할 원본 파일
        self.is_processing = False
        self.cancel_event = threading.Event()

        self._create_header()
        self._create_file_section()
        self._create_attach_section()
        self._create_output_section()
        self._create_options_section()
        self._create_progress_section()
        self._create_result_section()
        self._create_statusbar()
        self._bind_events()

    # ================================================================
    #  UI 구성
    # ================================================================

    def _create_header(self):
        header = tk.Frame(self.root, bg=self.HEADER_BG, height=50)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="G-PDF", font=("Georgia", 20, "bold"),
                 bg=self.HEADER_BG, fg="white").pack(side="left", padx=15)
        tk.Label(header, text="문서 변환 & PDF 압축", font=("맑은 고딕", 11),
                 bg=self.HEADER_BG, fg="#aabbcc").pack(side="left", padx=5)

    def _create_file_section(self):
        frame = tk.LabelFrame(
            self.root, text=" PDF로 변환/압축할 문서 ",
            font=("맑은 고딕", 10, "bold"), bg=self.BODY_BG, padx=10, pady=6
        )
        frame.pack(fill="x", padx=12, pady=(8, 3))

        btn_frame = tk.Frame(frame, bg=self.BODY_BG)
        btn_frame.pack(fill="x")

        btn_s = {"font": ("맑은 고딕", 9), "relief": "flat", "cursor": "hand2",
                 "padx": 8, "pady": 3}

        tk.Button(btn_frame, text="파일 추가", bg="#3498db", fg="white",
                  command=self._add_files, **btn_s).pack(side="left", padx=(0, 4))
        tk.Button(btn_frame, text="폴더 추가", bg="#2980b9", fg="white",
                  command=self._add_folder, **btn_s).pack(side="left", padx=(0, 4))
        tk.Button(btn_frame, text="목록 비우기", bg="#e74c3c", fg="white",
                  command=self._clear_files, **btn_s).pack(side="left")

        self.file_count_label = tk.Label(
            btn_frame, text="0개", font=("맑은 고딕", 9), bg=self.BODY_BG, fg="#666"
        )
        self.file_count_label.pack(side="right")

        list_frame = tk.Frame(frame, bg=self.BODY_BG)
        list_frame.pack(fill="x", pady=(4, 0))

        sb = tk.Scrollbar(list_frame)
        sb.pack(side="right", fill="y")

        self.file_listbox = tk.Listbox(
            list_frame, height=5, font=("맑은 고딕", 9),
            yscrollcommand=sb.set, selectmode="extended"
        )
        self.file_listbox.pack(fill="x", expand=True)
        sb.config(command=self.file_listbox.yview)

        hint = "드래그로 추가 가능" if HAS_DND else "버튼으로 추가"
        tk.Label(frame, text=f"PDF / 한글 / Word / Excel / PPT 지원  ({hint}  |  Delete키로 제거)",
                 font=("맑은 고딕", 8), bg=self.BODY_BG, fg="#999").pack(anchor="w")

    def _create_attach_section(self):
        frame = tk.LabelFrame(
            self.root, text=" 첨부 파일 (PDF 안에 원본 보관) ",
            font=("맑은 고딕", 10, "bold"), bg=self.BODY_BG, padx=10, pady=6
        )
        frame.pack(fill="x", padx=12, pady=(0, 3))

        btn_frame = tk.Frame(frame, bg=self.BODY_BG)
        btn_frame.pack(fill="x")

        btn_s = {"font": ("맑은 고딕", 9), "relief": "flat", "cursor": "hand2",
                 "padx": 8, "pady": 3}

        tk.Button(btn_frame, text="첨부 추가", bg="#8e44ad", fg="white",
                  command=self._add_attach, **btn_s).pack(side="left", padx=(0, 4))
        tk.Button(btn_frame, text="선택 삭제", bg="#95a5a6", fg="white",
                  command=self._remove_attach, **btn_s).pack(side="left", padx=(0, 4))
        tk.Button(btn_frame, text="전체 삭제", bg="#e74c3c", fg="white",
                  command=self._clear_attach, **btn_s).pack(side="left")

        self.attach_count_label = tk.Label(
            btn_frame, text="0개", font=("맑은 고딕", 9), bg=self.BODY_BG, fg="#666"
        )
        self.attach_count_label.pack(side="right")

        list_frame = tk.Frame(frame, bg=self.BODY_BG)
        list_frame.pack(fill="x", pady=(3, 0))

        sb = tk.Scrollbar(list_frame)
        sb.pack(side="right", fill="y")

        self.attach_listbox = tk.Listbox(
            list_frame, height=5, font=("맑은 고딕", 9),
            yscrollcommand=sb.set, selectmode="extended"
        )
        self.attach_listbox.pack(fill="x", expand=True)
        sb.config(command=self.attach_listbox.yview)
        self.attach_listbox.bind("<Delete>", lambda e: self._remove_attach())

        attach_hint = "드래그로 추가 가능  |  " if HAS_DND else ""
        tk.Label(frame,
                 text=f"PDF 출력 시 첨부된 원본이 함께 보관됩니다 ({attach_hint}Adobe Reader에서 추출 가능, 용량 증가)",
                 font=("맑은 고딕", 8), bg=self.BODY_BG, fg="#999").pack(anchor="w")

    def _create_output_section(self):
        frame = tk.LabelFrame(
            self.root, text=" 출력 폴더 ",
            font=("맑은 고딕", 10, "bold"), bg=self.BODY_BG, padx=10, pady=6
        )
        frame.pack(fill="x", padx=12, pady=(0, 3))

        row = tk.Frame(frame, bg=self.BODY_BG)
        row.pack(fill="x")

        tk.Label(row, text="저장 위치:", font=("맑은 고딕", 9),
                 bg=self.BODY_BG).pack(side="left", padx=(0, 6))

        self._PRESET_SAME = "원본과 같은 폴더"
        self._PRESET_BROWSE = "저장 폴더 지정..."

        self._output_presets = [self._PRESET_SAME] + self.recent_output_dirs + [self._PRESET_BROWSE]
        self._output_var = tk.StringVar(value=self._PRESET_SAME)

        self._output_combo = ttk.Combobox(
            row, textvariable=self._output_var,
            values=self._output_presets,
            state="readonly", font=("맑은 고딕", 9), width=36,
        )
        self._output_combo.pack(side="left", fill="x", expand=True)
        self._output_combo.bind("<<ComboboxSelected>>", self._on_output_select)

    def _create_options_section(self):
        frame = tk.LabelFrame(
            self.root, text=" 압축 설정 ",
            font=("맑은 고딕", 10, "bold"), bg=self.BODY_BG, padx=10, pady=6
        )
        frame.pack(fill="x", padx=12, pady=(0, 3))

        # 1행: 압축 모드 라디오
        radio_row = tk.Frame(frame, bg=self.BODY_BG)
        radio_row.pack(fill="x")

        self._mode_var = tk.StringVar(value="standard")

        for key, info in self.COMPRESS_MODES.items():
            tk.Radiobutton(
                radio_row, text=info["label"], variable=self._mode_var, value=key,
                font=("맑은 고딕", 9), bg=self.BODY_BG, activebackground=self.BODY_BG,
                command=self._on_mode_change,
            ).pack(side="left", padx=(0, 6))

        # 구분선
        tk.Frame(radio_row, width=1, bg="#ccc").pack(side="left", fill="y", padx=(8, 8), pady=2)

        # PDF/A 체크
        self._pdfa_var = tk.BooleanVar(value=True)
        self._pdfa_cb = tk.Checkbutton(
            radio_row, text="PDF/A", variable=self._pdfa_var,
            font=("맑은 고딕", 9), bg=self.BODY_BG, activebackground=self.BODY_BG,
        )
        self._pdfa_cb.pack(side="left")
        tk.Label(radio_row, text="(장기보존 국제표준)",
                 font=("맑은 고딕", 8), bg=self.BODY_BG, fg="#999").pack(side="left", padx=(2, 0))

        # 2행: 설명
        self._desc_label = tk.Label(
            frame, text="", font=("맑은 고딕", 8), bg=self.BODY_BG, fg="#888", anchor="w"
        )
        self._desc_label.pack(fill="x", pady=(2, 0))

        self._on_mode_change()

    def _create_progress_section(self):
        frame = tk.Frame(self.root, bg=self.BODY_BG)
        frame.pack(fill="x", padx=12, pady=(0, 3))

        self.progress_label = tk.Label(
            frame, text="대기 중", font=("맑은 고딕", 9),
            bg=self.BODY_BG, fg="#666"
        )
        self.progress_label.pack(anchor="w")

        self.progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(frame, variable=self.progress_var, maximum=100).pack(fill="x", pady=2)

        btn_row = tk.Frame(frame, bg=self.BODY_BG)
        btn_row.pack(fill="x", pady=(4, 0))

        btn_s = {"font": ("맑은 고딕", 10), "relief": "flat", "cursor": "hand2",
                 "padx": 15, "pady": 5}

        self.btn_start = tk.Button(
            btn_row, text="변환/압축 시작", bg="#27ae60", fg="white",
            command=self._start_process, **btn_s
        )
        self.btn_start.pack(side="left", padx=(0, 5))

        self.btn_cancel = tk.Button(
            btn_row, text="취소", bg="#e74c3c", fg="white",
            command=self._cancel_process, state="disabled", **btn_s
        )
        self.btn_cancel.pack(side="left")

    def _create_result_section(self):
        frame = tk.LabelFrame(
            self.root, text=" 결과 ",
            font=("맑은 고딕", 10, "bold"), bg=self.BODY_BG, padx=10, pady=6
        )
        frame.pack(fill="both", expand=True, padx=12, pady=(0, 3))

        self.result_text = tk.Text(
            frame, height=6, font=("맑은 고딕", 9), wrap="word",
            state="disabled", bg="white"
        )
        self.result_text.pack(fill="both", expand=True)

    def _create_statusbar(self):
        bar = tk.Frame(self.root, bg=self.HEADER_BG, height=26)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        self.status_text = tk.Label(
            bar, text="준비", font=("맑은 고딕", 9),
            bg=self.HEADER_BG, fg="#aabbcc", padx=10
        )
        self.status_text.pack(side="left")

        compress_method = get_compression_method()
        tk.Label(bar, text=f"압축: {compress_method}", font=("맑은 고딕", 8),
                 bg=self.HEADER_BG, fg="#778899", padx=8).pack(side="right")

        engine_text = get_engine_status_text()
        tk.Label(bar, text=engine_text, font=("맑은 고딕", 8),
                 bg=self.HEADER_BG, fg="#778899", padx=8).pack(side="right")

    def _bind_events(self):
        if HAS_DND:
            self.file_listbox.drop_target_register(DND_FILES)
            self.file_listbox.dnd_bind("<<Drop>>", self._on_drop)
            self.attach_listbox.drop_target_register(DND_FILES)
            self.attach_listbox.dnd_bind("<<Drop>>", self._on_drop_attach)
        self.file_listbox.bind("<Delete>", self._remove_file_selected)

    # ================================================================
    #  이벤트
    # ================================================================

    def _on_mode_change(self):
        mode = self._mode_var.get()
        desc = self.COMPRESS_MODES.get(mode, {}).get("desc", "")
        self._desc_label.config(text=desc)
        if mode == "original":
            self._pdfa_cb.config(state="disabled")
            self._pdfa_var.set(False)
        else:
            self._pdfa_cb.config(state="normal")

    # ================================================================
    #  메인 파일 관리
    # ================================================================

    def _is_supported(self, path):
        return os.path.splitext(path)[1].lower() in self._ALL_EXTS

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="파일 선택", initialdir=self.last_dir or "",
            filetypes=[
                ("지원 파일", "*.pdf *.hwp *.hwpx *.docx *.doc *.xlsx *.xls *.pptx *.ppt"),
                ("PDF", "*.pdf"), ("한글", "*.hwp *.hwpx"),
                ("Word", "*.docx *.doc"), ("Excel", "*.xlsx *.xls"),
                ("PowerPoint", "*.pptx *.ppt"), ("모든 파일", "*.*"),
            ]
        )
        if not paths:
            return
        self.last_dir = os.path.dirname(paths[0])
        self._save_config()
        added = 0
        for p in paths:
            if self._is_supported(p) and p not in self.file_list:
                self.file_list.append(p)
                added += 1
        self._refresh_file_list()
        if added:
            self.status_text.config(text=f"{added}개 파일 추가")

    def _add_folder(self):
        folder = filedialog.askdirectory(title="폴더 선택", initialdir=self.last_dir or "")
        if not folder:
            return
        self.last_dir = folder
        self._save_config()
        added = 0
        for f in sorted(os.listdir(folder)):
            ext = os.path.splitext(f)[1].lower()
            if ext in self._ALL_EXTS:
                full = os.path.join(folder, f)
                if full not in self.file_list:
                    self.file_list.append(full)
                    added += 1
        self._refresh_file_list()
        if added:
            self.status_text.config(text=f"폴더에서 {added}개 파일 추가")

    def _on_drop(self, event):
        raw = event.data.strip()
        if "{" in raw:
            import re
            paths = re.findall(r'\{([^}]+)\}', raw)
        else:
            paths = raw.split()
        added = 0
        for p in paths:
            p = p.strip('"')
            if self._is_supported(p) and p not in self.file_list:
                self.file_list.append(p)
                added += 1
        if added:
            self._refresh_file_list()
            self.status_text.config(text=f"{added}개 파일 추가")

    def _on_drop_attach(self, event):
        raw = event.data.strip()
        if "{" in raw:
            import re
            paths = re.findall(r'\{([^}]+)\}', raw)
        else:
            paths = raw.split()
        added = 0
        for p in paths:
            p = p.strip('"')
            if os.path.isfile(p) and p not in self.attach_list:
                self.attach_list.append(p)
                added += 1
        if added:
            self._refresh_attach_list()
            self.status_text.config(text=f"첨부 {added}개 추가")
            if self._mode_var.get() != "original":
                self._pdfa_var.set(True)

    def _clear_files(self):
        self.file_list.clear()
        self._refresh_file_list()

    def _remove_file_selected(self, event=None):
        for idx in sorted(self.file_listbox.curselection(), reverse=True):
            if idx < len(self.file_list):
                self.file_list.pop(idx)
        self._refresh_file_list()

    def _refresh_file_list(self):
        self.file_listbox.delete(0, "end")
        for p in self.file_list:
            name = os.path.basename(p)
            ext = os.path.splitext(p)[1].lower()
            try:
                size = format_size(os.path.getsize(p))
            except OSError:
                size = "?"
            tag = "  [PDF변환]" if ext in self._DOC_EXTS else ""
            self.file_listbox.insert("end", f"{name}  ({size}){tag}")
        self.file_count_label.config(text=f"{len(self.file_list)}개")

    # ================================================================
    #  첨부 파일 관리
    # ================================================================

    def _add_attach(self):
        paths = filedialog.askopenfilenames(
            title="첨부할 원본 파일 선택", initialdir=self.last_dir or "",
            filetypes=[
                ("오피스 문서", "*.xlsx *.xls *.hwp *.hwpx *.docx *.doc *.pptx *.ppt"),
                ("이미지", "*.jpg *.jpeg *.png *.tif *.tiff *.bmp"),
                ("데이터", "*.csv *.xml *.json *.txt"),
                ("모든 파일", "*.*"),
            ]
        )
        if not paths:
            return
        added = 0
        for p in paths:
            if p not in self.attach_list:
                self.attach_list.append(p)
                added += 1
        self._refresh_attach_list()
        if added:
            self.status_text.config(text=f"첨부 {added}개 추가")
            # 첨부 파일 추가 시 PDF/A 자동 체크
            if self._mode_var.get() != "original":
                self._pdfa_var.set(True)

    def _remove_attach(self):
        for idx in sorted(self.attach_listbox.curselection(), reverse=True):
            if idx < len(self.attach_list):
                self.attach_list.pop(idx)
        self._refresh_attach_list()

    def _clear_attach(self):
        self.attach_list.clear()
        self._refresh_attach_list()

    def _refresh_attach_list(self):
        self.attach_listbox.delete(0, "end")
        for p in self.attach_list:
            name = os.path.basename(p)
            try:
                size = format_size(os.path.getsize(p))
            except OSError:
                size = "?"
            self.attach_listbox.insert("end", f"{name}  ({size})")
        self.attach_count_label.config(text=f"{len(self.attach_list)}개")

    # ================================================================
    #  출력 폴더
    # ================================================================

    def _on_output_select(self, _event=None):
        val = self._output_var.get()
        if val != self._PRESET_BROWSE:
            if val != self._PRESET_SAME:
                self.output_dir = val
                self._save_config()
            return

        # 저장 폴더 지정... 클릭 시 다이얼로그
        folder = filedialog.askdirectory(title="저장 폴더 선택", initialdir=self.output_dir or "")
        if folder:
            self.output_dir = folder
            # 최근 목록 갱신: 맨 앞에 추가, 중복 제거, 최대 4개
            dirs = [folder] + [d for d in self.recent_output_dirs if d != folder]
            self.recent_output_dirs = dirs[:4]
            self._save_config()
            self._output_combo["values"] = [self._PRESET_SAME] + self.recent_output_dirs + [self._PRESET_BROWSE]
            self._output_var.set(folder)
        else:
            # 취소 시 이전 값 복원
            prev = self.output_dir if (self.output_dir and os.path.isdir(self.output_dir)) else self._PRESET_SAME
            self._output_var.set(prev)

    # ================================================================
    #  실행
    # ================================================================

    def _resolve_output_dir(self, input_path):
        val = self._output_var.get()
        if val == self._PRESET_SAME:
            return os.path.dirname(input_path)
        return val  # 직접 선택한 경로

    def _get_output_path(self, input_path):
        name, ext = os.path.splitext(os.path.basename(input_path))
        out_ext = ".pdf" if ext.lower() in self._DOC_EXTS else ext
        suffix = "_G-Fit"

        out_dir = self._resolve_output_dir(input_path)
        output_path = os.path.join(out_dir, f"{name}{suffix}{out_ext}")

        counter = 1
        while os.path.exists(output_path):
            output_path = os.path.join(out_dir, f"{name}{suffix}({counter}){out_ext}")
            counter += 1
        return output_path

    def _start_process(self):
        if self.is_processing:
            return
        if not self.file_list:
            messagebox.showwarning("알림", "파일을 추가해주세요.", parent=self.root)
            return
        val = self._output_var.get()
        if val != self._PRESET_SAME and not os.path.isdir(val):
            messagebox.showwarning("알림", "출력 폴더가 유효하지 않습니다.", parent=self.root)
            return

        # 파일 존재 확인
        missing = [os.path.basename(p) for p in self.file_list if not os.path.isfile(p)]
        if missing:
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{chr(10).join(missing[:5])}", parent=self.root)
            return

        self.is_processing = True
        self.cancel_event.clear()
        self.btn_start.config(state="disabled", text="처리 중...")
        self.btn_cancel.config(state="normal")
        self.progress_var.set(0)
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.config(state="disabled")

        mode = self._mode_var.get()
        pdfa = self._pdfa_var.get()
        files = list(self.file_list)
        attachments = list(self.attach_list)

        thread = threading.Thread(
            target=self._process_worker,
            args=(files, mode, pdfa, attachments),
            daemon=True
        )
        thread.start()

    def _cancel_process(self):
        if self.is_processing:
            self.cancel_event.set()
            self.btn_cancel.config(state="disabled", text="취소 중...")

    def _process_worker(self, files, mode, pdfa, attachments):
        total = len(files)
        success_count = 0
        fail_count = 0
        total_out_size = 0

        mode_info = self.COMPRESS_MODES.get(mode, self.COMPRESS_MODES["standard"])
        do_compress = mode != "original"

        for i, input_path in enumerate(files):
            if self.cancel_event.is_set():
                self.root.after(0, self._append_result, "\n[중단됨] 사용자가 취소했습니다.\n")
                break

            filename = os.path.basename(input_path)
            ext = os.path.splitext(input_path)[1].lower()
            is_doc = ext in self._DOC_EXTS

            try:
                # ── 1단계: 문서 변환 (필요 시) ──
                pdf_input = input_path
                temp_pdf = None

                if is_doc:
                    self.root.after(0, self._update_progress,
                                    (i / total) * 100, f"[{i+1}/{total}] {filename} PDF 변환 중...")

                    temp_pdf = os.path.join(tempfile.gettempdir(),
                                            os.path.splitext(filename)[0] + "_gpdf_conv.pdf")
                    conv = convert_to_pdf(input_path, temp_pdf)
                    if not conv["success"]:
                        fail_count += 1
                        self.root.after(0, self._append_result,
                                        f"[변환실패] {filename}: {conv['error']}\n")
                        continue

                    pdf_input = temp_pdf
                    self.root.after(0, self._append_result,
                                    f"[변환] {filename} -> PDF ({conv['engine']})\n")

                output_path = self._get_output_path(input_path)

                # ── 2단계: 압축 (또는 단순 복사) ──
                # 첨부 파일 구성: 공통 첨부 + 문서 변환 시 원본 자동 첨부
                file_attachments = list(attachments)
                if is_doc and input_path not in file_attachments:
                    file_attachments.append(input_path)

                if do_compress:
                    self.root.after(0, self._update_progress,
                                    (i / total) * 100, f"[{i+1}/{total}] {filename} 압축 중...")

                    result = compress_pdf(
                        pdf_input, output_path,
                        image_quality=mode_info["quality"],
                        image_dpi=mode_info["dpi"],
                        lossless=mode_info["lossless"],
                        pdfa=pdfa,
                        attachments=file_attachments if file_attachments else None,
                    )

                    if result["success"]:
                        success_count += 1
                        out_size = os.path.getsize(output_path)
                        total_out_size += out_size
                        attach_msg = f" +첨부{len(file_attachments)}개" if file_attachments else ""
                        self.root.after(0, self._append_result,
                            f"[완료] {filename} -> {os.path.basename(output_path)} "
                            f"({format_size(out_size)}, {result['ratio']:.1f}% 감소){attach_msg}\n")
                    else:
                        fail_count += 1
                        self.root.after(0, self._append_result,
                            f"[실패] {filename}: {result['error']}\n")
                else:
                    # 원본 출력 모드: 변환만 하고 복사
                    import shutil
                    shutil.copy2(pdf_input, output_path)

                    # 첨부 파일이 있으면 pikepdf로 추가
                    if file_attachments:
                        try:
                            from pdf_compressor import _attach_files, HAS_PIKEPDF
                            if HAS_PIKEPDF:
                                _attach_files(output_path, file_attachments)
                        except Exception:
                            pass

                    success_count += 1
                    out_size = os.path.getsize(output_path)
                    total_out_size += out_size
                    attach_msg = f" +첨부{len(file_attachments)}개" if file_attachments else ""
                    self.root.after(0, self._append_result,
                        f"[완료] {filename} -> {os.path.basename(output_path)} "
                        f"({format_size(out_size)}){attach_msg}\n")

                # 임시 파일 정리
                if temp_pdf and os.path.exists(temp_pdf):
                    try:
                        os.remove(temp_pdf)
                    except OSError:
                        pass

            except Exception as e:
                fail_count += 1
                self.root.after(0, self._append_result, f"[오류] {filename}: {e}\n")

        summary = (
            f"\n{'─'*50}\n"
            f"완료: 성공 {success_count}개 / 실패 {fail_count}개\n"
            f"총 출력: {format_size(total_out_size)}\n"
        )
        self.root.after(0, self._on_done, summary)

    # ================================================================
    #  UI 업데이트
    # ================================================================

    def _update_progress(self, pct, text):
        self.progress_var.set(pct)
        self.progress_label.config(text=text)

    def _append_result(self, text):
        self.result_text.config(state="normal")
        self.result_text.insert("end", text)
        self.result_text.see("end")
        self.result_text.config(state="disabled")

    def _on_done(self, summary):
        self.is_processing = False
        self.btn_start.config(state="normal", text="변환/압축 시작")
        self.btn_cancel.config(state="disabled", text="취소")
        self.progress_var.set(100)
        self.progress_label.config(text="완료")
        self._append_result(summary)
        self.root.title("G-PDF")

        # 첫 번째 입력 파일 기준으로 출력 폴더 결정
        open_dir = None
        if self.file_list:
            candidate = self._resolve_output_dir(self.file_list[0])
            if os.path.isdir(candidate):
                open_dir = candidate
        if open_dir:
            if messagebox.askyesno("완료", "처리가 완료되었습니다.\n출력 폴더를 열겠습니까?",
                                   parent=self.root):
                os.startfile(open_dir)

    # ================================================================
    #  설정
    # ================================================================

    def _load_config(self):
        try:
            with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_config(self):
        config = self._load_config()
        config["last_dir"] = self.last_dir
        config["output_dir"] = self.output_dir
        config["recent_output_dirs"] = self.recent_output_dirs
        try:
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = GPDFApp()
    app.run()
