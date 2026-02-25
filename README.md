# PDF 통합 처리기

PDF 파일을 분할, 병합, 이미지로 변환할 수 있는 통합 프로그램입니다.

## 주요 기능

### 1. PDF 분할
- **페이지 범위로 분할**: 특정 페이지 범위를 별도 파일로 추출
- **각 페이지를 개별 파일로 분할**: 모든 페이지를 개별 PDF 파일로 저장
- **페이지 단위로 분할**: N페이지씩 묶어서 여러 파일로 분할

### 2. PDF 병합
- 여러 PDF 파일을 하나로 병합
- 이미지 파일(PNG, JPG, BMP, TIFF)도 PDF로 변환하여 병합
- 파일 순서 조정 가능
- 페이지 크기 및 방향 설정

### 3. PDF → 이미지 변환
- PDF의 각 페이지를 이미지 파일로 변환
- 지원 형식: PNG, JPG, BMP, TIFF
- 해상도(DPI) 설정: 72, 96, 150, 200, 300, 600
- 전체 페이지 또는 특정 범위만 변환
- 파일명 패턴 설정 가능

### 4. 미리보기 기능
- PDF 페이지 실시간 미리보기
- 이미지 파일 미리보기
- 페이지 네비게이션

## 버전

### Python 버전
- Python 3.x 필요
- 필요한 라이브러리: PyPDF2, PyMuPDF (fitz), Pillow, tkinter

### C# 버전
- .NET 6.0 이상 필요
- WPF 기반 Windows 애플리케이션
- 필요한 라이브러리: PDFsharp

## 설치 및 실행

### Python 버전

#### 필수 패키지 설치
```bash
pip install -r requirements.txt
```

#### 프로그램 실행
```bash
python pdf_processor.py
```

#### EXE 파일 생성
```bash
python build_exe.py
```
또는
```bash
빌드_최종.bat
```

생성된 파일: `dist/PDF통합처리기.exe`

### C# 버전

#### 필수 요구사항
- .NET 6.0 SDK 이상
- Windows 10/11

#### 프로그램 실행 (개발 모드)
```bash
dotnet run
```

#### EXE 파일 생성
```bash
C샵_EXE_생성.bat
```
또는
```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o dist_csharp_single
```

생성된 파일: `dist_csharp_single/PDF통합처리기.exe`

#### 특징
- **완전한 단일 EXE 파일**: 추가 DLL 없이 단독 실행 가능
- **파일 크기**: 약 180MB (모든 라이브러리 및 .NET 런타임 포함)
- **자동 DLL 추출**: 첫 실행 시 필요한 네이티브 라이브러리를 자동으로 추출
- **이식성**: 파일 하나만 복사하면 어디서든 실행 가능

#### 사용된 라이브러리
- **PDFsharp**: PDF 분할 및 병합
- **PdfiumViewer**: PDF 렌더링 및 이미지 변환
- **PdfiumViewer.Native.x86_64.v8-xfa**: 네이티브 PDF 렌더링 엔진 (내장)
- **System.Drawing.Common**: 이미지 처리

## 사용 방법

### PDF 분할
1. "PDF 분할" 탭 선택
2. "PDF 파일 선택" 버튼으로 분할할 PDF 선택
3. 분할 방식 선택:
   - 페이지 범위: 시작/끝 페이지 입력
   - 개별 파일: 각 페이지를 별도 파일로
   - 페이지 단위: N페이지씩 묶어서 분할
4. 출력 폴더 선택 (선택사항)
5. "PDF 분할 실행" 버튼 클릭

### PDF 병합
1. "PDF 병합" 탭 선택
2. "추가" 버튼으로 병합할 파일 추가
3. ↑↓ 버튼으로 파일 순서 조정
4. 출력 파일명과 저장 위치 설정
5. 페이지 크기와 방향 선택
6. "PDF 병합 실행" 버튼 클릭

### PDF → 이미지 변환
1. "PDF → 이미지" 탭 선택
2. "PDF 파일 선택" 버튼으로 변환할 PDF 선택
3. 이미지 형식 선택 (PNG, JPG, BMP, TIFF)
4. 해상도(DPI) 설정
5. 변환할 페이지 범위 선택
6. 출력 폴더 및 파일명 패턴 설정
7. "이미지로 변환" 버튼 클릭

## 파일명 패턴

변환 시 사용할 수 있는 패턴:
- `{filename}`: 원본 PDF 파일명
- `{page}`: 페이지 번호

예시:
- `{filename}_page_{page}` → `문서_page_1.png`, `문서_page_2.png`
- `output_{page}` → `output_1.png`, `output_2.png`

## 지원 파일 형식

### 입력
- PDF: `.pdf`
- 이미지 (병합용): `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tiff`, `.gif`

### 출력
- PDF: `.pdf`
- 이미지: `.png`, `.jpg`, `.bmp`, `.tiff`

## 라이선스

이 프로그램은 개인 및 상업적 용도로 자유롭게 사용할 수 있습니다.

## 문의

문제가 발생하거나 기능 제안이 있으시면 이슈를 등록해주세요.
