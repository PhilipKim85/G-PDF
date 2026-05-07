# G-PDF

강남구청 세무관리과에서 자체개발한 PDF 변환·압축 도구

## 기능

- **문서 → PDF 변환**: HWP/HWPX, Word, Excel, PowerPoint → PDF
  - MS Office COM, 한컴오피스 COM, LibreOffice headless 엔진 자동 선택
- **PDF 압축 (G-Fit)**: 최대 98% 압축률
  - Ghostscript 이미지 다운샘플링 + JPEG 재압축 + 폰트 서브셋 (손실 압축)
  - pikepdf 스트림 최적화 + 객체 중복 제거 (무손실 최적화)
- **PDF/A 장기보존 표준 출력**: PDF/A-1 규격 변환
- **PDF/A-3 원본 파일 첨부**: 원본 문서를 PDF에 임베딩
- **워터마크 삽입**: 텍스트 워터마크 자동 삽입
- **드래그 & 드롭 지원**: 파일을 끌어다 놓기만 하면 변환/압축

## 요구사항

### 실행 환경 (빌드된 exe)
- Windows 10/11
- 별도 Python 설치 불필요 (PyInstaller로 단일 exe 빌드)

### 개발 환경
- Python 3.9+
- Windows 10/11

### 의존 라이브러리

```
PyMuPDF (fitz)    - PDF 처리
pikepdf           - PDF 무손실 최적화
Pillow            - 이미지 처리
tkinterdnd2       - 드래그&드롭 지원 (선택)
```

### 외부 도구 (번들 포함)
- Ghostscript (`gswin64c.exe`) - PDF 압축 엔진
- rhwp (`rhwp.exe`) - HWP 변환

## 사용 방법

### 빌드된 exe 실행
```
G-PDF.exe
```

### 소스에서 실행
```bash
pip install pymupdf pikepdf Pillow tkinterdnd2
python main.py
```

### PyInstaller 빌드
```bash
pip install pyinstaller
pyinstaller G-PDF.spec
```

## 외부 의존 프로그램 설치

이 프로그램은 다음 외부 프로그램을 필요로 합니다.
사용자가 직접 다운로드하여 `data/` 폴더에 넣어주세요.

1. **Ghostscript**
   - 다운로드: https://www.ghostscript.com/releases/gsdnld.html
   - 필요 파일: `gswin64c.exe`, `gsdll64.dll`
   - 라이선스: AGPL-3.0

2. **rhwp**
   - 다운로드: https://github.com/nicedoc/rhwp
   - 필요 파일: `rhwp.exe`
   - 라이선스: MIT

## 프로젝트 구조

```
G-PDF/
├── main.py              # 메인 애플리케이션 (GUI)
├── pdf_compressor.py    # G-Fit PDF 압축 엔진
├── doc_converter.py     # 문서 → PDF 변환 엔진
├── compress_widget.py   # 압축 옵션 UI 위젯
├── data/                # 번들 외부 도구
│   ├── gswin64c.exe
│   ├── gsdll64.dll
│   └── rhwp.exe
├── LICENSE              # AGPL-3.0 라이선스
├── NOTICE.txt           # 저작권 및 서드파티 고지
└── LICENSES/            # 의존 라이브러리 라이선스 전문
```

## 라이선스

본 소프트웨어는 **GNU Affero General Public License v3.0 (AGPL-3.0)**으로 배포됩니다.

- 사용·수정·배포 자유
- 수정·배포 시 같은 AGPL-3.0으로 소스코드 공개 필요
- 자세한 내용은 [LICENSE](LICENSE) 파일 참조

## 저작권

```
Copyright (C) 2025-2026 서울특별시 강남구 (Gangnam-gu, Seoul)
```

- **저작권자**: 서울특별시 강남구
- **개발자**: 김택중 (강남구청 세무관리과)

## 의존성 라이선스

| 라이브러리 | 라이선스 | 용도 |
|-----------|---------|------|
| Ghostscript | AGPL-3.0 | PDF 압축 엔진 |
| PyMuPDF | AGPL-3.0 | PDF 처리 |
| pikepdf | MPL-2.0 | PDF 무손실 최적화 |
| pypdf | BSD-3-Clause | PDF 처리 |
| Pillow | MIT-CMU | 이미지 처리 |
| rhwp | MIT | HWP 변환 |
| LibreOffice | MPL-2.0/LGPL-3.0 | 문서 변환 |
| tkinterdnd2 | MIT | 드래그&드롭 |

각 라이선스 전문은 [LICENSES/](LICENSES/) 폴더를 참조하세요.
