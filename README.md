# PPEditer

**Planning Proposal Editor** — PPTX 기획서 편집기

---

## 소개

PPEditer는 기획서(Planning Proposal) 작성 및 편집을 위한 오픈소스 프레젠테이션 편집기입니다.
C# + WPF로 구현되어 Windows 표준 환경(메뉴, 단축키, 테마)을 완벽히 준수하며,
Microsoft PowerPoint(.pptx) 파일과 완전히 호환됩니다.

---

## 기술 스택 및 라이선스

| 구성 요소 | 라이선스 | 역할 |
|---|---|---|
| **C# / WPF** (.NET 8) | MIT (Microsoft) | UI 프레임워크 |
| **DocumentFormat.OpenXml SDK v3** | MIT (Microsoft) | PPTX 읽기·쓰기 |
| **PDFsharp** | MIT (empira) | PDF 내보내기 |

> 모든 의존 라이브러리가 **MIT 라이선스**이므로, 본 프로젝트의 MIT 라이선스와 완전히 호환됩니다.

---

## 주요 기능

### 파일 관리
| 기능 | 설명 |
|---|---|
| 새로 만들기 / 열기 | `.pptx` 파일 생성 및 열기 |
| 저장 / 다른 이름으로 저장 | `.pptx` 형식으로 저장 |
| 최근 파일 | 레지스트리에 저장 (최대 10개) |
| PDF 내보내기 | 전체 슬라이드를 PDF로 내보내기 (150 DPI) |
| 암호 설정 | PPTX 열기 암호 설정 및 제거 |

### 슬라이드 편집
| 기능 | 설명 |
|---|---|
| 슬라이드 추가 / 삭제 / 복제 | 현재 위치 기준으로 슬라이드 관리 |
| 순서 변경 | 드래그 또는 단축키로 슬라이드 순서 이동 |
| 실행 취소 / 다시 실행 | Ctrl+Z / Ctrl+Y (최대 50단계) |

### 도형 및 텍스트
| 기능 | 설명 |
|---|---|
| 18종 도형 그리기 | 직선, 사각형, 삼각형, 원, 폴리곤, 화살표 등 |
| 텍스트 편집 | 더블클릭으로 인라인 텍스트 편집 |
| 도형 이동 / 크기 조절 | 드래그 및 방향키 미세 이동 |
| 그룹 / 언그룹 | Ctrl+G / Ctrl+Shift+G |
| 겹침 순서 | 맨 앞 / 앞으로 / 뒤로 / 맨 뒤 4단계 |
| 특수 삽입 | 수식·과학 기호, 특수 문자, 이모지 |

### 서식 지정
| 기능 | 설명 |
|---|---|
| 글자 속성 | 글꼴, 크기, 굵기, 기울임, 밑줄, 색상, 첨자 |
| 문단 속성 | 정렬, 줄 간격, 들여쓰기 |
| 도형 속성 | 채우기, 테두리, 위치/크기, 세로 정렬 |
| 문서 정보 | 슬라이드 크기, 방향, 암호 |

### 효과 넣기
| 기능 | 설명 |
|---|---|
| 슬라이드 전환 효과 | 없음 / 밝기변화 / 밀어내기 / 나타나기 / 집어내기 / 구겨던지기 / 모핑 |
| 개체 애니메이션 | 없음 / 밝기변화 / 날아오기 / 펄스 / 튕김 / 닦아내기 |
| 시간 설정 | 0.1초 단위 |
| 반복 / 자동 실행 | 반복 횟수 (0=무한), 슬라이드 표시 시 자동 재생 옵션 |

### 슬라이드 쇼
| 기능 | 설명 |
|---|---|
| 전체 화면 재생 | F5로 시작 |
| 슬라이드 탐색 | 클릭, PageDown/Up, 방향키 |
| 애니메이션 순서 재생 | 자동 실행 또는 클릭 시 재생 |

### 보기 및 환경
| 기능 | 설명 |
|---|---|
| 확대/축소 | 25% ~ 300%, 창 맞춤 |
| 테마 | 밝은 테마 / 어두운 테마 |
| 다국어 | 한국어 / English |

---

## 단축키

| 단축키 | 동작 |
|---|---|
| `Ctrl+N` | 새로 만들기 |
| `Ctrl+O` | 열기 |
| `Ctrl+S` | 저장 |
| `Ctrl+Shift+S` | 다른 이름으로 저장 |
| `Ctrl+Shift+E` | PDF 내보내기 |
| `Ctrl+Z` | 실행 취소 |
| `Ctrl+Y` | 다시 실행 |
| `Ctrl+M` | 슬라이드 추가 |
| `Ctrl+D` | 슬라이드 복제 |
| `Ctrl+Shift+Delete` | 슬라이드 삭제 |
| `Ctrl+Shift+↑/↓` | 슬라이드 순서 이동 |
| `Ctrl+G` | 그룹으로 묶기 |
| `Ctrl+Shift+G` | 그룹 해제 |
| `Ctrl+Shift+T` | 텍스트 상자 삽입 |
| `Ctrl+Shift+F` | 글자 속성 |
| `Ctrl+Shift+P` | 문단 속성 |
| `Ctrl+Shift+I` | 문서 정보 |
| `Ctrl+Shift+M` | 수식/과학 기호 |
| `Ctrl+Shift+K` | 특수 문자 |
| `Ctrl+Shift+J` | 이모지 |
| `Ctrl+B / I / U` | 굵게 / 기울임 / 밑줄 (텍스트 편집 중) |
| `Ctrl++ / Ctrl+-` | 확대 / 축소 |
| `Ctrl+0` | 창 맞춤 |
| `F5` | 슬라이드 쇼 시작 |
| `Space / PageDown / →` | 다음 슬라이드/애니메이션 (쇼 중) |
| `Backspace / PageUp / ←` | 이전 슬라이드 (쇼 중) |
| `Esc` | 슬라이드 쇼 종료 / 텍스트 편집 취소 |

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
    │   ├── PresentationModel.cs      # PPTX 문서 모델 + 실행취소
    │   ├── TransitionKind.cs         # 슬라이드 전환 효과 모델
    │   └── AnimationKind.cs          # 개체 애니메이션 모델
    ├── Rendering/
    │   └── SlideRenderer.cs          # OpenXml → WPF Canvas 변환
    ├── Conversion/
    │   └── PptxConverter.cs          # FlowDocument ↔ OpenXml 변환
    ├── Export/
    │   └── PdfExporter.cs            # PDFsharp 기반 PDF 내보내기
    ├── Controls/
    │   ├── SlideThumbnailPanel       # 왼쪽 썸네일 패널
    │   └── SlideEditorCanvas         # 메인 편집 영역
    ├── Resources/
    │   ├── Strings.ko.xaml           # 한국어 문자열
    │   ├── Strings.en.xaml           # 영어 문자열
    │   ├── Theme.Light.xaml          # 밝은 테마
    │   └── Theme.Dark.xaml           # 어두운 테마
    └── Dialogs/
        ├── AboutDialog               # 정보 대화상자
        ├── UserManualDialog          # 사용자 메뉴얼
        ├── SlideShowWindow           # 슬라이드 쇼
        ├── TransitionDialog          # 슬라이드 전환 효과 설정
        ├── AnimationDialog           # 개체 애니메이션 설정
        └── ...                       # 글자/문단/도형/문서 속성 등
```

---

## 히스토리

변경 이력은 [HISTORY.md](./HISTORY.md)를 참조하세요.

---

## 라이선스

MIT License — Copyright © 2026 Noh JinMoon

---

## 회사 정보

**HANDTECH (핸텍) — 상상공작소**
저작권자: 노진문 (Noh JinMoon)
