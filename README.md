# PPEditer

**Planning Proposal Editor** — PPT/PPTX 기획서 편집기

---

## 소개

PPEditer는 기획서(Planning Proposal) 작성 및 편집을 위한 오픈소스 프레젠테이션 편집기입니다.
C# + WPF로 구현되어 Windows 표준 환경(메뉴, 단축키, 테마)을 완벽히 준수합니다.

---

## 기술 스택 및 라이선스

| 구성 요소 | 라이선스 | 역할 |
|----------|---------|------|
| **C# / WPF** (.NET 8) | MIT (Microsoft) | UI 프레임워크 |
| **DocumentFormat.OpenXml SDK** | MIT (Microsoft) | PPTX 읽기·쓰기 |
| **PDFsharp** | MIT (empira) | PDF 내보내기 |

> 모든 의존 라이브러리가 **MIT 라이선스**이므로, 본 프로젝트의 MIT 라이선스와 완전히 호환됩니다.

---

## 기능

| 기능 | 설명 |
|------|------|
| **파일 열기** | `.pptx`, `.ppt` 파일 열기 |
| **슬라이드 편집** | 텍스트 인라인 편집 (더블클릭) |
| **슬라이드 관리** | 추가 / 삭제 / 복제 / 순서 변경 |
| **저장** | `.pptx` 형식으로 저장 |
| **PDF 내보내기** | 전체 슬라이드를 PDF로 내보내기 (150 DPI) |
| **실행 취소/다시 실행** | Ctrl+Z / Ctrl+Y (최대 50단계) |
| **확대/축소** | 창 맞춤, 25%~300% 줌 |
| **최근 파일** | 레지스트리에 저장 (최대 10개) |

---

## 단축키

| 단축키 | 동작 |
|--------|------|
| `Ctrl+N` | 새 프레젠테이션 |
| `Ctrl+O` | 파일 열기 |
| `Ctrl+S` | 저장 |
| `Ctrl+Shift+S` | 다른 이름으로 저장 |
| `Ctrl+Shift+E` | PDF 내보내기 |
| `Ctrl+Z` | 실행 취소 |
| `Ctrl+Y` | 다시 실행 |
| `Ctrl+M` | 새 슬라이드 추가 |
| `Ctrl+D` | 슬라이드 복제 |
| `Ctrl+Shift+Delete` | 슬라이드 삭제 |
| `Ctrl+Shift+↑/↓` | 슬라이드 순서 이동 |
| `Ctrl+=` / `Ctrl+-` | 확대 / 축소 |
| `Ctrl+0` | 창 맞춤 |

---

## 빌드 및 실행

### 요구사항

- .NET 8 SDK
- Windows (WPF는 Windows 전용)

### 빌드

```bash
cd PPEditer
dotnet build
```

### 실행

```bash
dotnet run
# 또는 파일을 직접 열기
dotnet run -- 파일.pptx
```

### 배포 (단일 실행 파일)

```bash
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

---

## 프로젝트 구조

```
PPEditer/
├── PPEditer.sln
└── PPEditer/
    ├── PPEditer.csproj
    ├── App.xaml / App.xaml.cs
    ├── MainWindow.xaml / MainWindow.xaml.cs
    ├── Models/
    │   └── PresentationModel.cs      # PPTX 문서 모델 + 실행취소
    ├── Rendering/
    │   └── SlideRenderer.cs          # OpenXml → WPF Canvas 변환
    ├── Export/
    │   └── PdfExporter.cs            # PDFsharp 기반 PDF 내보내기
    ├── Controls/
    │   ├── SlideThumbnailPanel       # 왼쪽 썸네일 패널
    │   └── SlideEditorCanvas         # 메인 편집 영역
    └── Dialogs/
        └── AboutDialog               # 정보 대화상자
```

---

## 라이선스

MIT License — Copyright © 2026 Noh JinMoon

---

## 회사 정보

**HANDTECH (핸텍) — 상상공작소**
저작권자: 노진문 (Noh JinMoon)
