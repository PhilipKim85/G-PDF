"""
G-Fit 압축 옵션 위젯 (Tkinter 공용)

다른 프로젝트(G-서명2, G-photoreport 등)에서 PDF 출력 시
압축 모드를 선택할 수 있는 컴팩트한 UI 위젯.

사용법:
    from compress_widget import CompressOptionWidget, apply_compression

    # UI에 위젯 추가
    widget = CompressOptionWidget(parent_frame)
    widget.pack(...)

    # PDF 저장 후 압축 적용
    mode, pdfa, attachments = widget.get()
    if mode != "original":
        apply_compression(pdf_path, mode, pdfa, attachments=attachments)
"""

import os
import tkinter as tk
from tkinter import filedialog

# 압축 모드 정의
COMPRESS_MODES = {
    "original":  {"label": "원본 출력",       "desc": "압축 없이 그대로 저장"},
    "lossless":  {"label": "무손실 최적화",   "desc": "품질 유지, 구조만 최적화"},
    "standard":  {"label": "표준 압축 (권장)", "desc": "품질/용량 균형"},
    "strong":    {"label": "강한 압축",       "desc": "용량 최소화 우선"},
}

# 각 모드별 compress_pdf 파라미터
COMPRESS_PARAMS = {
    "lossless":  {"image_quality": 95, "image_dpi": 300, "lossless": True,  "pdfa": False},
    "standard":  {"image_quality": 65, "image_dpi": 150, "lossless": False, "pdfa": False},
    "strong":    {"image_quality": 40, "image_dpi": 120, "lossless": False, "pdfa": False},
}


class CompressOptionWidget(tk.Frame):
    """PDF 압축 모드 선택 위젯 — 라디오 버튼 + PDF/A + 원본 첨부"""

    def __init__(self, parent, default_mode="standard", bg=None, **kwargs):
        if bg is None:
            bg = parent.cget("bg") if hasattr(parent, "cget") else "#f0f0f0"
        super().__init__(parent, bg=bg, **kwargs)

        self._mode_var = tk.StringVar(value=default_mode)
        self._pdfa_var = tk.BooleanVar(value=True)
        self._attach_var = tk.BooleanVar(value=False)
        self._attach_files = []  # 첨부 파일 경로 리스트

        # ── 라벨프레임 ──
        frame = tk.LabelFrame(self, text="  PDF 출력 품질  ",
                              font=("맑은 고딕", 9), bg=bg, padx=8, pady=4)
        frame.pack(fill="x")

        # ── 1행: 압축 모드 라디오 버튼 ──
        radio_row = tk.Frame(frame, bg=bg)
        radio_row.pack(fill="x")

        for mode_key, mode_info in COMPRESS_MODES.items():
            tk.Radiobutton(
                radio_row,
                text=mode_info["label"],
                variable=self._mode_var,
                value=mode_key,
                font=("맑은 고딕", 9),
                bg=bg,
                activebackground=bg,
                command=self._on_mode_change,
            ).pack(side="left", padx=(0, 6))

        # ── 구분선 ──
        sep = tk.Frame(radio_row, width=1, bg="#ccc")
        sep.pack(side="left", fill="y", padx=(8, 8), pady=2)

        # ── PDF/A 체크박스 ──
        self._pdfa_cb = tk.Checkbutton(
            radio_row,
            text="PDF/A",
            variable=self._pdfa_var,
            font=("맑은 고딕", 9),
            bg=bg,
            activebackground=bg,
            command=self._on_pdfa_change,
        )
        self._pdfa_cb.pack(side="left")

        tk.Label(
            radio_row, text="(장기보존 국제표준)",
            font=("맑은 고딕", 8), bg=bg, fg="#999",
        ).pack(side="left", padx=(2, 0))

        # ── 2행: 원본 첨부 옵션 ──
        attach_row = tk.Frame(frame, bg=bg)
        attach_row.pack(fill="x", pady=(3, 0))

        self._attach_cb = tk.Checkbutton(
            attach_row,
            text="원본 파일 첨부",
            variable=self._attach_var,
            font=("맑은 고딕", 9),
            bg=bg,
            activebackground=bg,
            command=self._on_attach_change,
        )
        self._attach_cb.pack(side="left")

        tk.Label(
            attach_row, text="Excel·한글 등 원본을 PDF 안에 보관",
            font=("맑은 고딕", 8), bg=bg, fg="#999",
        ).pack(side="left", padx=(2, 0))

        # 첨부 파일 추가 버튼
        self._attach_btn = tk.Button(
            attach_row, text="파일 선택...",
            font=("맑은 고딕", 8), relief="flat", cursor="hand2",
            command=self._select_attach_files, state="disabled",
        )
        self._attach_btn.pack(side="left", padx=(8, 0))

        # 첨부 파일 개수 라벨
        self._attach_count = tk.Label(
            attach_row, text="",
            font=("맑은 고딕", 8), bg=bg, fg="#2980b9",
        )
        self._attach_count.pack(side="left", padx=(4, 0))

        # 첨부 파일 초기화 버튼
        self._attach_clear_btn = tk.Button(
            attach_row, text="초기화",
            font=("맑은 고딕", 8), relief="flat", cursor="hand2",
            fg="#e74c3c", command=self._clear_attach_files,
        )

        # ── 3행: 설명 라벨 ──
        self._desc_label = tk.Label(
            frame, text="", font=("맑은 고딕", 8),
            bg=bg, fg="#888", anchor="w"
        )
        self._desc_label.pack(fill="x", pady=(2, 0))

        self._on_mode_change()

    def _on_mode_change(self):
        mode = self._mode_var.get()
        desc = COMPRESS_MODES.get(mode, {}).get("desc", "")
        self._desc_label.config(text=desc)

        if mode == "original":
            self._pdfa_cb.config(state="disabled")
            self._pdfa_var.set(False)
            self._attach_cb.config(state="disabled")
            self._attach_var.set(False)
            self._on_attach_change()
        else:
            self._pdfa_cb.config(state="normal")
            self._pdfa_var.set(True)
            self._on_pdfa_change()

    def _on_pdfa_change(self):
        if self._pdfa_var.get():
            self._attach_cb.config(state="normal")
        else:
            self._attach_cb.config(state="disabled")
            self._attach_var.set(False)
            self._on_attach_change()

    def _on_attach_change(self):
        if self._attach_var.get():
            self._attach_btn.config(state="normal")
        else:
            self._attach_btn.config(state="disabled")
            self._attach_files.clear()
            self._update_attach_label()

    def _select_attach_files(self):
        files = filedialog.askopenfilenames(
            title="첨부할 원본 파일 선택",
            filetypes=[
                ("오피스 문서", "*.xlsx *.xls *.hwp *.hwpx *.docx *.doc *.pptx"),
                ("이미지", "*.jpg *.jpeg *.png *.tif *.tiff *.bmp"),
                ("데이터", "*.csv *.xml *.json *.txt"),
                ("모든 파일", "*.*"),
            ]
        )
        if files:
            for f in files:
                if f not in self._attach_files:
                    self._attach_files.append(f)
            self._update_attach_label()

    def _clear_attach_files(self):
        self._attach_files.clear()
        self._update_attach_label()

    def _update_attach_label(self):
        count = len(self._attach_files)
        if count > 0:
            names = [os.path.basename(f) for f in self._attach_files[:3]]
            text = ", ".join(names)
            if count > 3:
                text += f" 외 {count - 3}개"
            self._attach_count.config(text=f"📎 {count}개: {text}")
            self._attach_clear_btn.pack(side="left", padx=(4, 0))
        else:
            self._attach_count.config(text="")
            self._attach_clear_btn.pack_forget()

    def get(self):
        """현재 선택된 (mode, pdfa, attachments) 튜플 반환"""
        attachments = list(self._attach_files) if self._attach_var.get() else []
        return self._mode_var.get(), self._pdfa_var.get(), attachments

    def set_mode(self, mode):
        """모드를 프로그래밍 방식으로 설정"""
        if mode in COMPRESS_MODES:
            self._mode_var.set(mode)
            self._on_mode_change()

    def set_attach_files(self, file_paths):
        """첨부 파일을 프로그래밍 방식으로 설정 (외부에서 자동 추가용)"""
        self._attach_files = [f for f in file_paths if os.path.isfile(f)]
        if self._attach_files:
            self._attach_var.set(True)
            self._on_attach_change()
        self._update_attach_label()

    def get_params(self):
        """compress_pdf()에 전달할 파라미터 dict 반환 (original이면 None)"""
        mode, pdfa, attachments = self.get()
        if mode == "original":
            return None
        params = dict(COMPRESS_PARAMS[mode])
        params["pdfa"] = pdfa
        if attachments:
            params["attachments"] = attachments
        return params


def apply_compression(pdf_path, mode="standard", pdfa=False, output_path=None,
                      attachments=None):
    """
    PDF 파일에 G-Fit 압축을 적용하는 헬퍼 함수.

    Args:
        pdf_path: 원본 PDF 경로
        mode: "lossless" | "standard" | "strong"
        pdfa: PDF/A 출력 여부
        output_path: 출력 경로 (None이면 원본 덮어쓰기)
        attachments: 첨부할 파일 경로 리스트

    Returns:
        dict: compress_pdf 결과 (success, original_size, compressed_size, ratio, ...)
              또는 mode가 "original"이면 None
    """
    if mode == "original":
        return None

    from pdf_compressor import compress_pdf

    params = dict(COMPRESS_PARAMS.get(mode, COMPRESS_PARAMS["standard"]))
    params["pdfa"] = pdfa

    if output_path is None:
        # 임시 파일로 압축 후 원본 교체
        tmp_path = pdf_path + ".gfit_tmp"
        try:
            result = compress_pdf(
                pdf_path, tmp_path,
                image_quality=params["image_quality"],
                image_dpi=params["image_dpi"],
                lossless=params["lossless"],
                pdfa=params["pdfa"],
                attachments=attachments,
            )
            if result["success"]:
                os.replace(tmp_path, pdf_path)
            return result
        except Exception as e:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
            return {"success": False, "error": str(e)}
    else:
        return compress_pdf(
            pdf_path, output_path,
            image_quality=params["image_quality"],
            image_dpi=params["image_dpi"],
            lossless=params["lossless"],
            pdfa=params["pdfa"],
            attachments=attachments,
        )
