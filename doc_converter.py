"""
G-Fit 문서 → PDF 변환 엔진

지원 형식: .hwp, .hwpx, .docx, .doc, .xlsx, .xls, .pptx, .ppt
변환 엔진 우선순위:
  1. MS Office COM (Word, Excel, PowerPoint) — 설치되어 있으면 사용
  2. 한컴오피스 COM (한글) — hwp/hwpx 전용
  3. LibreOffice headless — 위 두 가지 불가 시 대체

뷰어만 설치된 경우:
  - 한글뷰어: hwp → pdf 변환 불가 (COM 미지원)
  - MS Office: COM 지원되므로 변환 가능
  - LibreOffice: headless 모드 변환 가능
"""

import os
import sys
import shutil
import subprocess
import tempfile
from pathlib import Path


def _get_base_dir():
    """실행 파일 또는 스크립트 기준 디렉토리"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


_BASE_DIR = _get_base_dir()

# 확장자별 변환 카테고리
_WORD_EXTS = {".docx", ".doc"}
_EXCEL_EXTS = {".xlsx", ".xls"}
_PPT_EXTS = {".pptx", ".ppt"}
_HWP_EXTS = {".hwp", ".hwpx"}
_ALL_CONVERTIBLE = _WORD_EXTS | _EXCEL_EXTS | _PPT_EXTS | _HWP_EXTS


def is_convertible(file_path: str) -> bool:
    """변환 가능한 파일인지 확인"""
    ext = os.path.splitext(file_path)[1].lower()
    return ext in _ALL_CONVERTIBLE


def get_available_engines() -> dict:
    """사용 가능한 변환 엔진 목록 반환"""
    engines = {}

    # MS Office COM
    try:
        import win32com.client
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Quit()
            engines["word"] = True
        except Exception:
            engines["word"] = False
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Quit()
            engines["excel"] = True
        except Exception:
            engines["excel"] = False
        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Quit()
            engines["powerpoint"] = True
        except Exception:
            engines["powerpoint"] = False
    except ImportError:
        engines["word"] = False
        engines["excel"] = False
        engines["powerpoint"] = False

    # 한컴오피스 COM
    try:
        import win32com.client
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.Quit()
        engines["hancom"] = True
    except Exception:
        engines["hancom"] = False

    return engines



def convert_to_pdf(input_path: str, output_path: str = None) -> dict:
    """
    문서 파일을 PDF로 변환한다.

    Args:
        input_path: 원본 문서 경로 (.hwp, .docx, .xlsx 등)
        output_path: 출력 PDF 경로 (None이면 같은 폴더에 .pdf)

    Returns:
        dict: {success, output_path, engine, error}
    """
    input_path = os.path.abspath(input_path)
    ext = os.path.splitext(input_path)[1].lower()

    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
    output_path = os.path.abspath(output_path)

    result = {"success": False, "output_path": output_path, "engine": "", "error": None}

    if not os.path.isfile(input_path):
        result["error"] = f"파일을 찾을 수 없습니다: {input_path}"
        return result

    if ext not in _ALL_CONVERTIBLE:
        result["error"] = f"지원하지 않는 형식입니다: {ext}"
        return result

    # hwp/hwpx → rhwp(SVG→PDF)
    if ext in _HWP_EXTS:
        if _convert_rhwp(input_path, output_path, result):
            return result

    # MS Office 문서 → COM
    if ext in _WORD_EXTS:
        if _convert_word_com(input_path, output_path, result):
            return result
    elif ext in _EXCEL_EXTS:
        if _convert_excel_com(input_path, output_path, result):
            return result
    elif ext in _PPT_EXTS:
        if _convert_ppt_com(input_path, output_path, result):
            return result

    # 모든 엔진 실패
    if not result["error"]:
        result["error"] = (
            "변환 가능한 프로그램을 찾을 수 없습니다.\n"
            "다음 중 하나가 필요합니다:\n"
            "- Microsoft Office (Word/Excel/PowerPoint)\n"
            "- rhwp (한글/HWP)"
        )
    return result


def _convert_word_com(input_path, output_path, result) -> bool:
    """Word COM으로 변환"""
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            doc = word.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=17)  # 17 = wdFormatPDF
            doc.Close(SaveChanges=False)
            result["success"] = True
            result["engine"] = "Microsoft Word"
            return True
        except Exception as e:
            result["error"] = f"Word 변환 실패: {e}"
            return False
        finally:
            word.Quit()
    except ImportError:
        return False
    except Exception:
        return False


def _convert_excel_com(input_path, output_path, result) -> bool:
    """Excel COM으로 변환"""
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(input_path)
            wb.ExportAsFixedFormat(0, output_path)  # 0 = xlTypePDF
            wb.Close(SaveChanges=False)
            result["success"] = True
            result["engine"] = "Microsoft Excel"
            return True
        except Exception as e:
            result["error"] = f"Excel 변환 실패: {e}"
            return False
        finally:
            excel.Quit()
    except ImportError:
        return False
    except Exception:
        return False


def _convert_ppt_com(input_path, output_path, result) -> bool:
    """PowerPoint COM으로 변환"""
    try:
        import win32com.client
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        try:
            presentation = ppt.Presentations.Open(input_path, WithWindow=False)
            presentation.SaveAs(output_path, FileFormat=32)  # 32 = ppSaveAsPDF
            presentation.Close()
            result["success"] = True
            result["engine"] = "Microsoft PowerPoint"
            return True
        except Exception as e:
            result["error"] = f"PowerPoint 변환 실패: {e}"
            return False
        finally:
            ppt.Quit()
    except ImportError:
        return False
    except Exception:
        return False


def _convert_hancom_com(input_path, output_path, result) -> bool:
    """한컴오피스 COM으로 hwp/hwpx 변환 — 먼저 닫고 검증"""
    try:
        import win32com.client
        import time
    except ImportError:
        return False

    abs_output = os.path.abspath(output_path)
    methods = [
        ("FileSaveAsPdf", None),
        ("FileSaveAs_S", "PDF"),
    ]

    for action_name, fmt in methods:
        # 매 시도마다 새로 열고 닫기 (파일 잠금 방지)
        try:
            if os.path.exists(output_path):
                os.remove(output_path)

            hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            hwp.XHwpWindows.Item(0).Visible = True
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            hwp.Open(input_path, "HWP", "forceopen:true")
            time.sleep(0.5)

            hwp.HAction.GetDefault(action_name,
                                   hwp.HParameterSet.HFileOpenSave.HSet)
            hwp.HParameterSet.HFileOpenSave.filename = abs_output
            if fmt:
                hwp.HParameterSet.HFileOpenSave.Format = fmt
            hwp.HAction.Execute(action_name,
                                hwp.HParameterSet.HFileOpenSave.HSet)
            time.sleep(2)

            # 반드시 먼저 닫기
            hwp.XHwpWindows.Item(0).Visible = False
            hwp.Clear(1)
            hwp.Quit()
            time.sleep(0.5)

            # 닫은 후 검증
            if _is_valid_pdf(output_path):
                result["success"] = True
                result["engine"] = "한컴오피스"
                return True

        except Exception:
            try:
                hwp.XHwpWindows.Item(0).Visible = False
                hwp.Clear(1)
                hwp.Quit()
            except Exception:
                pass

    result["error"] = "한컴오피스에서 유효한 PDF를 생성하지 못했습니다"
    return False


def _is_valid_pdf(path):
    """PDF 파일이 존재하고 유효한 내용이 있는지 확인"""
    if not os.path.isfile(path):
        return False
    size = os.path.getsize(path)
    if size < 1000:
        return False
    try:
        with open(path, "rb") as f:
            header = f.read(5)
        if header != b"%PDF-":
            return False
        # PyMuPDF로 페이지 수 확인
        import fitz
        doc = fitz.open(path)
        pages = len(doc)
        has_content = False
        if pages > 0:
            text = doc[0].get_text()
            has_content = len(text.strip()) > 0
        doc.close()
        return pages > 0 and has_content
    except Exception:
        return False


def _find_rhwp() -> str | None:
    """rhwp CLI 실행 파일 경로를 찾는다."""
    # 번들 경로
    bundled = os.path.join(_BASE_DIR, "data", "rhwp.exe")
    if os.path.isfile(bundled):
        return bundled
    # PATH
    found = shutil.which("rhwp")
    if found:
        return found
    return None


def _convert_rhwp(input_path, output_path, result) -> bool:
    """rhwp CLI로 HWP/HWPX → SVG → PDF 변환"""
    rhwp = _find_rhwp()
    if not rhwp:
        return False

    try:
        import tempfile as _tmpmod

        # 1단계: HWP → SVG (폰트 서브셋 임베딩으로 한글 깨짐 방지)
        svg_dir = _tmpmod.mkdtemp(prefix="gpdf_rhwp_")
        proc = subprocess.run(
            [rhwp, "export-svg", input_path, "--embed-fonts", "-o", svg_dir],
            capture_output=True, text=True, timeout=120, encoding="utf-8", errors="replace",
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        )

        if proc.returncode != 0:
            result["error"] = f"rhwp 변환 실패: {proc.stderr[:200]}"
            return False

        # SVG 파일 찾기
        svg_files = sorted([
            os.path.join(svg_dir, f) for f in os.listdir(svg_dir)
            if f.lower().endswith(".svg")
        ])

        if not svg_files:
            result["error"] = "rhwp가 SVG를 생성하지 못했습니다"
            return False

        # 2단계: SVG → PDF (PyMuPDF 사용)
        import fitz

        merged = fitz.open()
        for svg_file in svg_files:
            svg_doc = fitz.open(svg_file)
            pdf_bytes = svg_doc.convert_to_pdf()
            svg_doc.close()
            pdf_page = fitz.open("pdf", pdf_bytes)
            merged.insert_pdf(pdf_page)
            pdf_page.close()

        if len(merged) == 0:
            result["error"] = "SVG 파싱 실패"
            return False

        merged.save(output_path)
        merged.close()

        # SVG 임시 폴더 정리
        import shutil as _shutil
        _shutil.rmtree(svg_dir, ignore_errors=True)

        if os.path.isfile(output_path) and os.path.getsize(output_path) > 0:
            result["success"] = True
            result["engine"] = "rhwp"
            return True

        result["error"] = "rhwp PDF 생성 실패"
        return False

    except subprocess.TimeoutExpired:
        result["error"] = "rhwp 변환 시간 초과 (2분)"
        return False
    except ImportError as e:
        result["error"] = f"SVG→PDF 변환 라이브러리 없음: {e}"
        return False
    except Exception as e:
        result["error"] = f"rhwp 오류: {e}"
        return False



def get_engine_status_text() -> str:
    """현재 사용 가능한 엔진 상태를 한 줄 텍스트로 반환"""
    engines = get_available_engines()
    available = []
    if _find_rhwp():
        available.append("rhwp(HWP)")
    if engines.get("word") or engines.get("excel") or engines.get("powerpoint"):
        available.append("MS Office")
    if engines.get("hancom"):
        available.append("한컴오피스")

    if available:
        return "변환 엔진: " + ", ".join(available)
    else:
        return "변환 엔진 없음 (MS Office/rhwp 필요)"
