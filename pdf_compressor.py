"""
G-Fit PDF 압축 엔진

압축 파이프라인:
1. Ghostscript - 이미지 다운샘플링 + JPEG 재압축 + 폰트 서브셋 (손실 압축)
2. pikepdf - 스트림 최적화 + 객체 중복 제거 (무손실 최적화)
3. 워터마크 삽입 - 텍스트 워터마크 삽입 (왼쪽 하단, 중복 스킵)

무손실 모드: Ghostscript를 건너뛰고 pikepdf만 사용 (이미지 품질 보존)
PDF/A 모드: Ghostscript -dPDFA=1 으로 장기보존 표준 출력
"""

import os
import subprocess
import shutil
import threading
import fitz  # PyMuPDF
from PIL import Image
import io

# pikepdf (무손실 최적화)
try:
    import pikepdf
    HAS_PIKEPDF = True
except ImportError:
    HAS_PIKEPDF = False

# 워터마크 설정
_WM_TEXT = "사용기관: 강남구"
_WM_FONT_SIZE = 9
_WM_IMG_CACHE = None  # 워터마크 PNG 캐시


def _get_watermark_image() -> bytes:
    """워터마크 텍스트를 작은 PNG 이미지로 렌더링하여 캐시.
    폰트를 PDF에 임베딩하지 않으므로 파일 크기 증가 없음."""
    global _WM_IMG_CACHE
    if _WM_IMG_CACHE is not None:
        return _WM_IMG_CACHE

    from PIL import ImageDraw, ImageFont
    font_path = r"C:\Windows\Fonts\malgun.ttf"
    font_size = 24  # 렌더링용 (고해상도로 그린 뒤 PDF에 작게 배치)
    try:
        font = ImageFont.truetype(font_path, font_size)
    except Exception:
        font = ImageFont.load_default()

    # 텍스트 크기 측정
    dummy = Image.new("RGBA", (1, 1))
    draw = ImageDraw.Draw(dummy)
    bbox = draw.textbbox((0, 0), _WM_TEXT, font=font)
    tw = bbox[2] - bbox[0] + 4
    th = bbox[3] - bbox[1] + 4

    # 투명 배경에 회색 텍스트
    img = Image.new("RGBA", (tw, th), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    draw.text((2, 2 - bbox[1]), _WM_TEXT, font=font, fill=(128, 128, 128, 255))

    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    _WM_IMG_CACHE = buf.getvalue()
    return _WM_IMG_CACHE


# 기본 경로 (소스 실행 / PyInstaller 빌드 모두 지원)
def _get_base_dir():
    """PyInstaller --onefile일 때는 sys._MEIPASS, 아니면 스크립트 위치"""
    import sys
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

_BASE_DIR = _get_base_dir()

# Ghostscript 경로 후보
_GS_PATHS = [
    os.path.join(_BASE_DIR, "data", "gswin64c.exe"),
    os.path.join(_BASE_DIR, "data", "gswin32c.exe"),
    r"C:\Program Files\gs\gs10.07.0\bin\gswin64c.exe",
    r"C:\Program Files\gs\gs10.04.0\bin\gswin64c.exe",
    r"C:\Program Files\gs\gs10.03.1\bin\gswin64c.exe",
    r"C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe",
    r"C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe",
    r"C:\Program Files (x86)\PDF Compressor\data\gswin32c.exe",
]

_gs_cache = None


def _find_ghostscript() -> str | None:
    """사용 가능한 Ghostscript를 찾아 캐시한다."""
    global _gs_cache
    if _gs_cache is not None:
        return _gs_cache if _gs_cache else None

    for path in _GS_PATHS:
        if os.path.isfile(path):
            _gs_cache = path
            return path
    for name in ("gswin64c", "gswin32c", "gs"):
        found = shutil.which(name)
        if found:
            _gs_cache = found
            return found
    _gs_cache = ""
    return None


def is_pdf_valid(path: str) -> tuple[bool, str]:
    """PDF 파일이 유효한지 검사한다. (암호화, 손상 등)"""
    try:
        doc = fitz.open(path)
        if doc.is_encrypted:
            doc.close()
            return False, "암호화된 PDF입니다. 암호를 해제한 후 다시 시도하세요."
        page_count = len(doc)
        doc.close()
        if page_count == 0:
            return False, "페이지가 없는 PDF입니다."
        return True, ""
    except Exception as e:
        return False, f"PDF를 열 수 없습니다: {e}"


def compress_pdf(input_path: str, output_path: str, image_quality: int = 50,
                 image_dpi: int = 150, lossless: bool = False, pdfa: bool = False,
                 pdf_type: str = "auto",
                 callback=None, cancel_event: threading.Event | None = None,
                 attachments: list | None = None) -> dict:
    """
    PDF 압축

    Args:
        input_path: 원본 PDF 경로
        output_path: 저장 PDF 경로
        image_quality: JPEG 품질 (1~100)
        image_dpi: 이미지 해상도 DPI
        lossless: True면 이미지 품질 손실 없이 구조만 최적화
        pdfa: True면 PDF/A-1b 규격으로 출력
        pdf_type: "auto" | "text" | "scanned"
        callback: 진행 콜백 (page_num, total_pages)
        cancel_event: 취소 이벤트 (set되면 중단)
        attachments: 첨부할 파일 경로 리스트 (PDF/A-3 원본 첨부)

    Returns:
        dict: {success, original_size, compressed_size, ratio, error, method}
    """
    result = {
        "success": False,
        "original_size": 0,
        "compressed_size": 0,
        "ratio": 0.0,
        "error": None,
        "method": "",
    }

    try:
        # 입력 파일 검증
        if not os.path.isfile(input_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {input_path}")

        valid, msg = is_pdf_valid(input_path)
        if not valid:
            raise ValueError(msg)

        result["original_size"] = os.path.getsize(input_path)
        methods = []

        # 취소 확인 헬퍼
        def _check_cancel():
            if cancel_event and cancel_event.is_set():
                raise InterruptedError("사용자가 압축을 취소했습니다.")

        _check_cancel()

        # 1단계: 압축
        if lossless:
            # 무손실: PyMuPDF 구조 최적화만
            _optimize_structure(input_path, output_path, callback)
            methods.append("무손실 최적화")
        else:
            gs_path = _find_ghostscript()
            if gs_path:
                detected = _compress_ghostscript(input_path, output_path, image_quality,
                                      image_dpi, gs_path, pdfa, pdf_type, callback)
                methods.append(f"Ghostscript({detected})")
            else:
                _compress_pymupdf(input_path, output_path, image_quality,
                                   image_dpi, callback)
                methods.append("PyMuPDF")

        _check_cancel()

        # 2단계: pikepdf 무손실 최적화
        if HAS_PIKEPDF:
            _optimize_pikepdf(output_path)
            methods.append("pikepdf")

        _check_cancel()

        # 3단계: 워터마크 삽입
        if _WM_TEXT:
            _add_watermark(output_path)

        # 4단계: 원본 파일 첨부 (PDF/A-3)
        if attachments and HAS_PIKEPDF:
            _attach_files(output_path, attachments)
            methods.append("원본첨부")

        result["compressed_size"] = os.path.getsize(output_path)
        if result["original_size"] > 0:
            result["ratio"] = (1 - result["compressed_size"] / result["original_size"]) * 100
        result["method"] = " + ".join(methods)
        result["success"] = True

    except InterruptedError as e:
        result["error"] = str(e)
        _safe_remove(output_path)
    except Exception as e:
        result["error"] = str(e)
        _safe_remove(output_path)

    return result


def _safe_remove(path: str):
    """파일 안전 삭제"""
    try:
        if os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


# ============ Ghostscript 압축 ============

def _is_scanned_pdf(pdf_path: str) -> bool:
    """PDF가 스캔 문서(이미지 기반)인지 판별.
    페이지 면적의 70% 이상을 이미지가 차지하면 스캔 PDF로 판단."""
    try:
        doc = fitz.open(pdf_path)
        try:
            check_pages = min(3, len(doc))
            scanned_count = 0
            for i in range(check_pages):
                page = doc[i]
                page_area = page.rect.width * page.rect.height
                if page_area == 0:
                    continue
                img_area = 0
                for img in page.get_images(full=True):
                    xref = img[0]
                    rects = page.get_image_rects(xref)
                    for r in rects:
                        img_area += r.width * r.height
                if img_area / page_area > 0.5:
                    scanned_count += 1
            return scanned_count > check_pages / 2
        finally:
            doc.close()
    except Exception:
        return False


def _compress_ghostscript(input_path: str, output_path: str,
                           quality: int, dpi: int, gs_path: str,
                           pdfa: bool = False, pdf_type: str = "auto",
                           callback=None):
    """Ghostscript로 PDF 재생성 (이미지 재압축 + 폰트 서브셋).
    Returns: 감지된 PDF 유형 문자열 ("텍스트" / "스캔")"""
    if callback:
        try:
            doc = fitz.open(input_path)
            total = len(doc)
            doc.close()
            callback(0, total)
        except Exception:
            pass

    # PDF 유형 판별
    if pdf_type == "auto":
        scanned = _is_scanned_pdf(input_path)
    else:
        scanned = (pdf_type == "scanned")

    detected_label = "스캔" if scanned else "텍스트"

    # quality → PDFSETTINGS 매핑
    if quality >= 70:
        pdf_settings = "/prepress"
    elif quality >= 40:
        pdf_settings = "/ebook"
    else:
        pdf_settings = "/screen"

    cmd = [
        gs_path,
        "-sDEVICE=pdfwrite",
        "-dBATCH", "-dNOPAUSE", "-dSAFER",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={pdf_settings}",
        # 이미지 다운샘플링
        f"-dColorImageResolution={dpi}",
        f"-dGrayImageResolution={dpi}",
        f"-dMonoImageResolution={max(dpi * 2, 200)}",
        "-dDownsampleColorImages=true",
        "-dDownsampleGrayImages=true",
        "-dDownsampleMonoImages=true",
        "-dColorImageDownsampleType=/Bicubic",
        "-dGrayImageDownsampleType=/Bicubic",
        # 폰트 + 구조 최적화
        "-dSubsetFonts=true",
        "-dCompressFonts=true",
        "-dEmbedAllFonts=true",
        "-dDetectDuplicateImages=true",
        "-dCompressPages=true",
        f"-sOutputFile={output_path}",
        input_path,
    ]

    # 스캔 PDF일 때만 JPEG 재압축 옵션 추가
    if scanned:
        cmd.insert(-2, "-dPassThroughJPEGImages=false")
        cmd.insert(-2, "-dColorImageDownsampleThreshold=1.0")
        cmd.insert(-2, "-dGrayImageDownsampleThreshold=1.0")
        cmd.insert(-2, "-dMonoImageDownsampleThreshold=1.0")

    # PDF/A 모드
    if pdfa:
        cmd.insert(-2, "-dPDFA=1")
        cmd.insert(-2, "-dPDFACompatibilityPolicy=1")
        cmd.insert(-2, "-sColorConversionStrategy=RGB")

    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True, timeout=600,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("Ghostscript 처리 시간이 초과되었습니다 (10분). 파일이 너무 크거나 복잡합니다.")

    if proc.returncode != 0:
        stderr = proc.stderr.strip() if proc.stderr else ""
        raise RuntimeError(f"Ghostscript 오류 (코드 {proc.returncode}): {stderr[:200]}")
    if not os.path.exists(output_path):
        raise RuntimeError("Ghostscript가 출력 파일을 생성하지 못했습니다.")

    return detected_label


# ============ 원본 파일 첨부 (PDF/A-3) ============

# 확장자 → MIME 타입 매핑
_MIME_TYPES = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls":  "application/vnd.ms-excel",
    ".hwp":  "application/x-hwp",
    ".hwpx": "application/haansofthwpx",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".doc":  "application/msword",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png":  "image/png",
    ".tif":  "image/tiff",
    ".tiff": "image/tiff",
    ".bmp":  "image/bmp",
    ".csv":  "text/csv",
    ".txt":  "text/plain",
    ".xml":  "application/xml",
    ".json": "application/json",
}


def _attach_files(pdf_path: str, file_paths: list):
    """pikepdf로 원본 파일을 PDF에 첨부한다."""
    if not HAS_PIKEPDF:
        return

    tmp_path = pdf_path + ".attach_tmp"
    try:
        pdf = pikepdf.open(pdf_path)
        try:
            for fpath in file_paths:
                if not os.path.isfile(fpath):
                    continue
                fname = os.path.basename(fpath)
                ext = os.path.splitext(fname)[1].lower()
                mime = _MIME_TYPES.get(ext, "application/octet-stream")

                with open(fpath, "rb") as f:
                    data = f.read()

                pdf.attachments[fname] = pikepdf.AttachedFileSpec(
                    pdf, data=data,
                    description=f"원본 파일: {fname}",
                    mime_type=mime,
                    relationship=pikepdf.Name("/Source"),
                )

            pdf.save(tmp_path)
        finally:
            pdf.close()

        os.replace(tmp_path, pdf_path)
    except Exception:
        _safe_remove(tmp_path)


# ============ pikepdf 무손실 최적화 ============

def _optimize_pikepdf(pdf_path: str):
    """pikepdf로 스트림 최적화 + 객체 중복 제거"""
    tmp_path = pdf_path + ".tmp"
    try:
        pdf = pikepdf.open(pdf_path)
        try:
            pdf.save(tmp_path,
                     object_stream_mode=pikepdf.ObjectStreamMode.generate,
                     compress_streams=True,
                     recompress_flate=True,
                     linearize=True)
        finally:
            pdf.close()

        # 작아졌을 때만 교체
        if os.path.getsize(tmp_path) < os.path.getsize(pdf_path):
            os.replace(tmp_path, pdf_path)
        else:
            _safe_remove(tmp_path)
    except Exception:
        _safe_remove(tmp_path)


# ============ PyMuPDF 폴백 압축 ============

def _compress_pymupdf(input_path: str, output_path: str,
                       quality: int, dpi: int, callback=None):
    """PyMuPDF로 이미지 재압축 (Ghostscript 없을 때 폴백)"""
    doc = fitz.open(input_path)
    try:
        total_pages = len(doc)
        processed_xrefs = set()

        for page_num in range(total_pages):
            if callback:
                callback(page_num, total_pages)

            page = doc[page_num]
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                if xref in processed_xrefs:
                    continue
                processed_xrefs.add(xref)

                try:
                    base_image = doc.extract_image(xref)
                    if not base_image:
                        continue

                    image_bytes = base_image["image"]
                    pil_img = Image.open(io.BytesIO(image_bytes))

                    # 투명 이미지는 건너뜀
                    if pil_img.mode in ("RGBA", "PA", "LA"):
                        continue

                    # DPI 기반 리샘플링
                    orig_w, orig_h = pil_img.size
                    scale = dpi / 300.0
                    if scale < 1.0:
                        new_w = max(1, int(orig_w * scale))
                        new_h = max(1, int(orig_h * scale))
                        if new_w < orig_w or new_h < orig_h:
                            pil_img = pil_img.resize((new_w, new_h), Image.Resampling.LANCZOS)

                    # 그레이스케일 유지
                    if pil_img.mode == "L":
                        out_mode = "L"
                    else:
                        pil_img = pil_img.convert("RGB")
                        out_mode = "RGB"

                    buf = io.BytesIO()
                    pil_img.save(buf, format="JPEG", quality=quality, optimize=True)
                    new_bytes = buf.getvalue()

                    if len(new_bytes) < len(image_bytes):
                        new_w, new_h = pil_img.size
                        doc.update_stream(xref, new_bytes)
                        doc.xref_set_key(xref, "Filter", "/DCTDecode")
                        doc.xref_set_key(xref, "Width", str(new_w))
                        doc.xref_set_key(xref, "Height", str(new_h))
                        cs = "/DeviceGray" if out_mode == "L" else "/DeviceRGB"
                        doc.xref_set_key(xref, "ColorSpace", cs)
                        doc.xref_set_key(xref, "BitsPerComponent", "8")
                        doc.xref_set_key(xref, "DecodeParms", "null")
                        doc.xref_set_key(xref, "Length", str(len(new_bytes)))
                except Exception:
                    continue

        doc.save(output_path, garbage=4, deflate=True, clean=True, use_objstms=True)
    finally:
        doc.close()


# ============ 무손실 구조 최적화 ============

def _optimize_structure(input_path: str, output_path: str, callback=None):
    """이미지 품질 손실 없이 PDF 구조만 최적화 (폰트 서브셋 불가, 스트림 재압축)"""
    if callback:
        try:
            doc = fitz.open(input_path)
            callback(0, len(doc))
            doc.close()
        except Exception:
            pass

    # Ghostscript로 구조 최적화 (이미지 다운샘플링 없음)
    gs_path = _find_ghostscript()
    if gs_path:
        cmd = [
            gs_path,
            "-sDEVICE=pdfwrite",
            "-dBATCH", "-dNOPAUSE", "-dSAFER",
            "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS=/prepress",
            # 이미지 다운샘플링 비활성화 (무손실)
            "-dDownsampleColorImages=false",
            "-dDownsampleGrayImages=false",
            "-dDownsampleMonoImages=false",
            # 폰트+구조 최적화
            "-dSubsetFonts=true",
            "-dCompressFonts=true",
            "-dEmbedAllFonts=true",
            "-dDetectDuplicateImages=true",
            "-dCompressPages=true",
            f"-sOutputFile={output_path}",
            input_path,
        ]
        try:
            proc = subprocess.run(
                cmd, capture_output=True, text=True, timeout=600,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
            )
        except subprocess.TimeoutExpired:
            raise RuntimeError("처리 시간 초과 (10분)")
        if proc.returncode != 0:
            raise RuntimeError(f"Ghostscript 오류: {proc.stderr.strip()[:200]}")
    else:
        # GS 없으면 PyMuPDF로 구조만 최적화
        doc = fitz.open(input_path)
        try:
            doc.save(output_path, garbage=4, deflate=True, clean=True, use_objstms=True)
        finally:
            doc.close()


# ============ 워터마크 ============

def _has_same_watermark(page, wm_text: str, margin: float = 10) -> tuple:
    """페이지 왼쪽 하단에 동일한 워터마크가 있는지 확인.
    Returns (has_same, bottom_y): has_same=True이면 동일 워터마크 존재, bottom_y는 기존 텍스트의 최하단 y좌표"""
    ph = page.rect.height
    # 왼쪽 하단 영역 검색 (하단 60pt, 왼쪽 250pt)
    search_rect = fitz.Rect(0, ph - 60, 250, ph)

    # 해당 영역의 전체 텍스트를 가져옴
    area_text = page.get_text("text", clip=search_rect)
    # 공백 모두 제거하고 비교 (띄어쓰기, 콜론 위치 등 무시)
    normalized_area = "".join(area_text.split())
    normalized_wm = "".join(wm_text.split())
    found_same = normalized_wm in normalized_area

    # 기존 텍스트의 최하단 y좌표 추적
    bottom_y = ph - margin
    blocks = page.get_text("blocks", clip=search_rect)
    for block in blocks:
        if block[6] != 0:  # 텍스트 블록만
            continue
        if block[3] > ph - 60:  # 하단 영역 텍스트
            bottom_y = min(bottom_y, block[1] - (_WM_FONT_SIZE + 8))

    return found_same, bottom_y


def _add_watermark(pdf_path: str):
    """모든 페이지에 워터마크 이미지 삽입 (왼쪽 하단, 중복 시 스킵).
    텍스트를 미리 PNG로 렌더링하여 삽입 → 폰트 임베딩 없음."""
    try:
        wm_bytes = _get_watermark_image()
        doc = fitz.open(pdf_path)
        try:
            modified = False
            for pg in doc:
                ph = pg.rect.height
                has_same, alt_y = _has_same_watermark(pg, _WM_TEXT)

                if has_same:
                    continue  # 동일 워터마크 이미 존재 → 스킵

                wm_margin = 10
                wm_h = _WM_FONT_SIZE + 4
                # 기본 위치: 페이지 최하단
                y_bottom = ph - wm_margin
                y_top = y_bottom - wm_h

                # 기존 텍스트가 같은 위치에 있으면 그 위에 배치
                if alt_y < y_top:
                    y_top = alt_y
                    y_bottom = y_top + wm_h

                # 이미지 비율 유지하며 높이에 맞춰 배치
                wm_rect = fitz.Rect(wm_margin, y_top, wm_margin + 120, y_bottom)
                pg.insert_image(wm_rect, stream=wm_bytes, keep_proportion=True)
                modified = True
            if modified:
                doc.saveIncr()
        finally:
            doc.close()
    except Exception:
        pass  # 워터마크 실패해도 압축 결과는 유지


# ============ 미리보기 ============

def generate_preview(pdf_path: str, page_num: int = 0, zoom: float = 1.0) -> Image.Image:
    """PDF 페이지를 PIL Image로 렌더링"""
    doc = fitz.open(pdf_path)
    try:
        if page_num >= len(doc):
            page_num = 0
        page = doc[page_num]
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        return img.copy()
    finally:
        doc.close()


def get_pdf_info(pdf_path: str) -> dict:
    """PDF 기본 정보 반환"""
    doc = fitz.open(pdf_path)
    try:
        info = {
            "pages": len(doc),
            "size": os.path.getsize(pdf_path),
            "encrypted": doc.is_encrypted,
        }
        if len(doc) > 0:
            page = doc[0]
            info["width"] = page.rect.width
            info["height"] = page.rect.height
        return info
    finally:
        doc.close()


# ============ 유틸리티 ============

def get_compression_method() -> str:
    """현재 사용 가능한 압축 엔진 조합"""
    parts = []
    gs = _find_ghostscript()
    if gs:
        parts.append("Ghostscript")
    else:
        parts.append("PyMuPDF")
    if HAS_PIKEPDF:
        parts.append("pikepdf")
    return " + ".join(parts)


def format_size(size_bytes: int) -> str:
    """바이트를 읽기 쉬운 크기로 변환"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
