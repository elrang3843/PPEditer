# PPEditer

**Planning Proposal Editor** — PPTX 기획서 편집기

[한국어](#한국어) | [English](#english)

---

<a name="한국어"></a>

## 한국어

### 소개

PPEditer는 기획서(Planning Proposal) 작성에 특화된 오픈소스 프레젠테이션 편집기입니다.  
Microsoft PowerPoint(`.pptx`) 파일과 완전히 호환되며, 설치 없이 단일 실행 파일 하나로 동작합니다.

---

### 시스템 요구사항

| 항목 | 최소 사양 |
|------|----------|
| 운영체제 | Windows 10 / 11 (64비트) |
| 런타임 | .NET 8 런타임 포함 (Self-contained 빌드) |
| 디스크 | 약 100 MB |

---

### 설치 및 실행

#### 방법 1 — 릴리즈 다운로드 (권장)

1. [Releases](../../releases) 페이지에서 최신 버전의 `PPEditer.exe`를 다운로드합니다.
2. 원하는 폴더에 저장합니다. (설치 프로그램 없음)
3. `PPEditer.exe`를 더블클릭하여 실행합니다.

> Self-contained 빌드이므로 .NET을 별도로 설치할 필요 없습니다.

#### 방법 2 — PPTX 파일 연결

1. `.pptx` 파일을 우클릭 → **연결 프로그램** → `PPEditer.exe` 선택
2. 이후 `.pptx` 파일을 더블클릭하면 PPEditer로 바로 열립니다.

---

### 빠른 시작

| 작업 | 방법 |
|------|------|
| 새 프레젠테이션 만들기 | `파일 > 새로 만들기` 또는 `Ctrl+N` |
| 기존 파일 열기 | `파일 > 열기` 또는 `Ctrl+O`, 또는 파일을 창으로 드래그 |
| 저장 | `Ctrl+S` (처음 저장 시 경로 지정) |
| 텍스트 편집 | 텍스트 상자를 **더블클릭** |
| 슬라이드 쇼 | `F5` |

---

### 주요 기능

#### 파일 관리

| 기능 | 설명 |
|------|------|
| 새로 만들기 / 열기 / 저장 | `.pptx` 형식 완전 호환 |
| 다른 이름으로 저장 | `Ctrl+Shift+S` |
| PDF 내보내기 | 전체 슬라이드 → PDF (150 DPI), `Ctrl+Shift+E` |
| 최근 파일 | 최대 10개 기록 |
| 자동 저장 | 1 / 2 / 5 / 10 / 30분 간격 설정 |
| 열기 암호 | PPTX 암호 설정 및 제거 |

#### 탬플릿

| 기능 | 설명 |
|------|------|
| 탬플릿 불러오기 | `.pptx` 파일을 탬플릿 편집 모드로 열기 |
| 탬플릿으로 새로 만들기 | 저장된 탬플릿 목록에서 선택하여 새 파일 생성 |
| 탬플릿 저장 | `%AppData%\PPEditer\Templates\` 폴더에 저장 |
| 탬플릿 편집 모드 | 제목 표시줄에 `[탬플릿 편집 중]` 표시 |

#### 슬라이드 관리

| 기능 | 설명 |
|------|------|
| 슬라이드 추가 / 삭제 / 복제 | 현재 위치 기준 (`Ctrl+M` / `Ctrl+Shift+Del` / `Ctrl+D`) |
| 순서 변경 | `Ctrl+Shift+↑/↓` 또는 패널에서 드래그 |
| 복사 / 잘라내기 / 붙여넣기 | 슬라이드 패널에서 `Ctrl+C / X / V`, 우클릭 메뉴 |
| 실행 취소 / 다시 실행 | `Ctrl+Z / Y` (최대 50단계) |
| 슬라이드 노트 | `보기 > 노트 표시/숨김` |

#### 도형 및 텍스트

| 기능 | 설명 |
|------|------|
| 18종 도형 그리기 | 직선, 사각형, 삼각형, 원/타원, 화살표, 폴리곤 등 |
| 텍스트 편집 | 더블클릭 인라인 편집 |
| 이동 / 크기 조절 | 드래그, 핸들, 방향키 미세 이동 |
| 회전 | 우클릭 → 회전 |
| 복사 / 잘라내기 / 붙여넣기 | `Ctrl+C / X / V`, 다중 선택 일괄 처리 |
| 그룹 / 언그룹 | `Ctrl+G / Ctrl+Shift+G` |
| 겹침 순서 | 맨 앞 / 앞으로 / 뒤로 / 맨 뒤 |
| 특수 삽입 | 수식·과학 기호, 특수 문자, 이모지 |

#### 서식

| 기능 | 설명 |
|------|------|
| 글자 속성 | 글꼴, 크기, 굵기, 기울임, 밑줄, 색상, 첨자 (`Ctrl+Shift+F`) |
| 문단 속성 | 정렬, 줄 간격, 들여쓰기 (`Ctrl+Shift+P`) |
| 도형 속성 | 채우기, 테두리, 위치/크기, 세로 정렬 |
| HSV 색상 선택기 | 채도·명도 평면 + HEX 입력 |

#### 효과 및 슬라이드 쇼

| 기능 | 설명 |
|------|------|
| 슬라이드 전환 | 없음 / 밝기변화 / 밀어내기 / 나타나기 / 집어내기 / 구겨던지기 / 모핑 |
| 개체 애니메이션 | 없음 / 밝기변화 / 날아오기 / 펄스 / 튕김 / 닦아내기 |
| 전체 화면 쇼 | `F5` |
| 발표자 도구 | 현재/다음 슬라이드·노트·경과 시간 동시 확인 |
| 모니터 설정 | 발표 화면과 발표자 화면 모니터 개별 선택 |

#### 보기 및 환경

| 기능 | 설명 |
|------|------|
| 확대/축소 | 25% ~ 300%, 창 맞춤 (`Ctrl+0`) |
| 테마 | 밝은 테마 / 어두운 테마 |
| 다국어 | 한국어 / English 런타임 전환 |
| 워터마크 | 편집·슬라이드 쇼·인쇄에 텍스트 워터마크 |
| 인쇄 | `Ctrl+P` |

---

### 단축키 전체 목록

| 단축키 | 동작 |
|--------|------|
| `Ctrl+N` | 새로 만들기 |
| `Ctrl+O` | 열기 |
| `Ctrl+S` | 저장 |
| `Ctrl+Shift+S` | 다른 이름으로 저장 |
| `Ctrl+P` | 인쇄 |
| `Ctrl+Shift+E` | PDF 내보내기 |
| `Ctrl+Z` | 실행 취소 |
| `Ctrl+Y` | 다시 실행 |
| `Ctrl+C / X / V` | 복사 / 잘라내기 / 붙여넣기 |
| `Ctrl+M` | 슬라이드 추가 |
| `Ctrl+D` | 슬라이드 복제 |
| `Ctrl+Shift+Delete` | 슬라이드 삭제 |
| `Ctrl+Shift+↑ / ↓` | 슬라이드 순서 이동 |
| `Ctrl+G` | 그룹으로 묶기 |
| `Ctrl+Shift+G` | 그룹 해제 |
| `Ctrl+Shift+T` | 텍스트 상자 삽입 |
| `Ctrl+Shift+F` | 글자 속성 |
| `Ctrl+Shift+P` | 문단 속성 |
| `Ctrl+Shift+I` | 문서 정보 |
| `Ctrl+Shift+M` | 수식 / 과학 기호 |
| `Ctrl+Shift+K` | 특수 문자 |
| `Ctrl+Shift+J` | 이모지 |
| `Ctrl+B / I / U` | 굵게 / 기울임 / 밑줄 (텍스트 편집 중) |
| `Ctrl++ / Ctrl+-` | 확대 / 축소 |
| `Ctrl+0` | 창 맞춤 |
| `F5` | 슬라이드 쇼 시작 |
| `Space / PageDown / →` | 다음 (쇼 중) |
| `Backspace / PageUp / ←` | 이전 (쇼 중) |
| `Esc` | 슬라이드 쇼 종료 / 편집 취소 |

---

<a name="english"></a>

## English

### Overview

PPEditer is an open-source presentation editor specialized for planning proposals.  
It is fully compatible with Microsoft PowerPoint (`.pptx`) files and runs as a single self-contained executable — no installation required.

---

### System Requirements

| Item | Requirement |
|------|-------------|
| OS | Windows 10 / 11 (64-bit) |
| Runtime | Bundled (.NET 8 self-contained) |
| Disk | ~100 MB |

---

### Installation & Launch

#### Option 1 — Download Release (Recommended)

1. Go to the [Releases](../../releases) page and download the latest `PPEditer.exe`.
2. Place it in any folder. (No installer needed)
3. Double-click `PPEditer.exe` to launch.

> The build is self-contained — no separate .NET installation is required.

#### Option 2 — Associate with .pptx Files

1. Right-click any `.pptx` file → **Open with** → select `PPEditer.exe`.
2. After that, double-clicking `.pptx` files will open them directly in PPEditer.

---

### Quick Start

| Task | How |
|------|-----|
| New presentation | `File > New` or `Ctrl+N` |
| Open a file | `File > Open` or `Ctrl+O`, or drag a file into the window |
| Save | `Ctrl+S` (prompts for path on first save) |
| Edit text | **Double-click** a text box |
| Slide show | `F5` |

---

### Key Features

#### File Management

| Feature | Description |
|---------|-------------|
| New / Open / Save | Full `.pptx` compatibility |
| Save As | `Ctrl+Shift+S` |
| PDF Export | All slides → PDF (150 DPI), `Ctrl+Shift+E` |
| Recent Files | Up to 10 entries |
| Auto Save | Configurable interval: 1 / 2 / 5 / 10 / 30 min |
| Password | Set or remove PPTX open password |

#### Templates

| Feature | Description |
|---------|-------------|
| Open as Template | Open any `.pptx` in template editing mode |
| New from Template | Pick a saved template to start a new file |
| Save Template | Stored in `%AppData%\PPEditer\Templates\` |
| Template Mode | Title bar shows `[Template Editing]` |

#### Slide Management

| Feature | Description |
|---------|-------------|
| Add / Delete / Duplicate | Relative to current slide (`Ctrl+M` / `Ctrl+Shift+Del` / `Ctrl+D`) |
| Reorder | `Ctrl+Shift+↑/↓` or drag in panel |
| Copy / Cut / Paste slides | `Ctrl+C / X / V` in slide panel, or right-click menu |
| Undo / Redo | `Ctrl+Z / Y` (up to 50 steps) |
| Slide Notes | `View > Show/Hide Notes` |

#### Shapes & Text

| Feature | Description |
|---------|-------------|
| 18 shape types | Line, rectangle, triangle, circle, arrow, polygon, and more |
| Text editing | Inline editing via double-click |
| Move / Resize | Drag, handles, or arrow keys for fine adjustment |
| Rotate | Right-click → Rotate |
| Copy / Cut / Paste | `Ctrl+C / X / V`, multi-select batch |
| Group / Ungroup | `Ctrl+G / Ctrl+Shift+G` |
| Z-order | Bring to front / forward / backward / send to back |
| Special insert | Math/science symbols, special characters, emoji |

#### Formatting

| Feature | Description |
|---------|-------------|
| Character | Font, size, bold, italic, underline, color, superscript (`Ctrl+Shift+F`) |
| Paragraph | Alignment, line spacing, indent (`Ctrl+Shift+P`) |
| Shape | Fill, border, position/size, vertical alignment |
| Color picker | HSV plane + HEX input |

#### Effects & Slide Show

| Feature | Description |
|---------|-------------|
| Slide transition | None / Fade / Push / Wipe / Flip / Crumple / Morph |
| Object animation | None / FadeIn / FlyIn / Pulse / Bounce / WipeIn |
| Full-screen show | `F5` |
| Presenter view | Current/next slide, notes, elapsed time |
| Monitor settings | Assign presentation and presenter screens independently |

#### View & Environment

| Feature | Description |
|---------|-------------|
| Zoom | 25% – 300%, fit to window (`Ctrl+0`) |
| Theme | Light / Dark |
| Language | Korean / English (runtime switch) |
| Watermark | Text watermark in editor, slide show, and print |
| Print | `Ctrl+P` |

---

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+N` | New |
| `Ctrl+O` | Open |
| `Ctrl+S` | Save |
| `Ctrl+Shift+S` | Save As |
| `Ctrl+P` | Print |
| `Ctrl+Shift+E` | Export PDF |
| `Ctrl+Z` | Undo |
| `Ctrl+Y` | Redo |
| `Ctrl+C / X / V` | Copy / Cut / Paste |
| `Ctrl+M` | Add slide |
| `Ctrl+D` | Duplicate slide |
| `Ctrl+Shift+Delete` | Delete slide |
| `Ctrl+Shift+↑ / ↓` | Move slide up / down |
| `Ctrl+G` | Group |
| `Ctrl+Shift+G` | Ungroup |
| `Ctrl+Shift+T` | Insert text box |
| `Ctrl+Shift+F` | Character properties |
| `Ctrl+Shift+P` | Paragraph properties |
| `Ctrl+Shift+I` | Document info |
| `Ctrl+Shift+M` | Math / science symbols |
| `Ctrl+Shift+K` | Special characters |
| `Ctrl+Shift+J` | Emoji |
| `Ctrl+B / I / U` | Bold / Italic / Underline (while editing text) |
| `Ctrl++ / Ctrl+-` | Zoom in / out |
| `Ctrl+0` | Fit to window |
| `F5` | Start slide show |
| `Space / PageDown / →` | Next (during show) |
| `Backspace / PageUp / ←` | Previous (during show) |
| `Esc` | End show / cancel edit |

---

## 개발자용 빌드 / Build from Source

```bash
# Requirements: .NET 8 SDK, Windows

cd PPEditer
dotnet build
dotnet run                   # blank window
dotnet run -- file.pptx      # open file on startup

# Single-file release build
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

---

## 변경 이력 / History

전체 변경 이력은 [HISTORY.md](./HISTORY.md)를 참조하세요.  
Full changelog: [HISTORY.md](./HISTORY.md)

---

## 기술 스택 / Tech Stack

| Component | License | Role |
|-----------|---------|------|
| C# / WPF (.NET 8) | MIT (Microsoft) | UI framework |
| DocumentFormat.OpenXml SDK v3 | MIT (Microsoft) | PPTX read/write |
| PDFsharp | MIT (empira) | PDF export |

All dependencies are MIT-licensed and fully compatible with this project's MIT license.

---

## 라이선스 / License

MIT License — Copyright © 2026 Noh JinMoon (노진문)

**HANDTECH (핸텍) — 상상공작소**
