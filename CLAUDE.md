# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## 빌드 및 실행

```bash
dotnet build PPEditer.sln                        # 디버그 빌드
dotnet build -c Release PPEditer.sln             # 릴리스 빌드
dotnet run --project PPEditer/PPEditer.csproj    # 실행
```

테스트 및 린트 설정 없음.

---

## 기술 스택

- **.NET 8.0 / WPF** (`net8.0-windows`, `UseWPF=true`)
- **DocumentFormat.OpenXml v3.2.0** — PPTX 읽기/쓰기
- **PDFsharp v6.1.1** — PDF 내보내기
- Nullable enabled, ImplicitUsings enabled

---

## 아키텍처 개요

### 폴더 역할

| 폴더 | 역할 |
|------|------|
| `Models/` | 문서 상태 (`PresentationModel`), 데이터 타입 (`CharStyle`, `ShapeStyle`, `DrawTool`, `RgbColor`) |
| `Services/` | PPTX↔WPF 변환 (`PptxConverter`), 폰트 목록 (`FontService`), 설정 (`AppSettings`) |
| `Rendering/` | PPTX SlidePart → WPF Canvas 렌더링 (`SlideRenderer`) |
| `Controls/` | 슬라이드 편집 캔버스 (`SlideEditorCanvas`), 썸네일 패널 (`SlideThumbnailPanel`) |
| `Dialogs/` | 모달 대화상자 — 개체 속성, 문자 속성, 색상 선택, 문서 정보 등 |
| `Export/` | PDF 내보내기 (`PdfExporter`) |
| `Resources/` | 아이콘, 다국어 문자열 (`Strings.ko.xaml` / `Strings.en.xaml`), 테마 |

---

## 핵심 데이터 흐름

### 파일 열기 → 렌더링
```
PresentationModel.Open(filePath)
  → 파일을 MemoryStream으로 로드 → PresentationDocument.Open(stream, editable)

MainWindow → EditorCanvas.ShowSlide(model, slideIndex)
  → SlideRenderer.BuildCanvas(slidePart, slideW, slideH)
  → ShapeTree 순회 → BuildShape / BuildPicture / BuildGroupShape
  → WPF Canvas (EMU 좌표 기준) → Viewbox로 감싸 UI 크기에 맞춤
```

### 텍스트 편집 → 커밋
```
사용자 텍스트 도형 클릭
  → SlideEditorCanvas가 RichTextBox 편집기 오버레이
  → PptxConverter.ToFlowDocument(TextBody) → WPF FlowDocument 변환
  → (편집)
  → LostFocus 시 SlideEditorCanvas.CommitEdit()
  → PptxConverter.FromFlowDocument(doc) → PptxParagraph 목록
  → PresentationModel.UpdateShapeContent(slideIdx, treeIdx, paragraphs)
```

### Undo/Redo
- 매 변경 전 `PushUndo()` → 전체 MemoryStream을 `byte[]` 스냅샷 (최대 50단계)
- `Undo()` / `Redo()` → 스냅샷 복원 후 UI 재구성

---

## 주요 클래스 역할

### `PresentationModel`
문서 상태 전담. 슬라이드 추가/삭제/이동, 도형 이동/크기조정/삭제/그룹화, Z-순서, GetSlidePart 등 모든 PPTX 변경 연산을 제공. 직접 PPTX XML을 조작.

### `PptxConverter`
PPTX TextBody ↔ WPF FlowDocument 양방향 변환. 변환 시 WPF `Run`이 직접 담지 못하는 속성(문자 간격·줄 색·외곽선·윗줄)은 `Run.Tag`에 `CharExtraProps` 인스턴스로 저장.

```csharp
public sealed class CharExtraProps
{
    public int       SpacingPt100   { get; set; }  // PPTX spc (1/100pt)
    public RgbColor? UnderlineColor { get; set; }
    public bool      HasOutline     { get; set; }
    public bool      HasOverline    { get; set; }  // WPF 전용, PPTX 저장 안 됨
}
```

### `SlideRenderer`
읽기 전용 렌더링. 좌표계는 EMU 기준 WPF 픽셀 (EmuToPx = value / 914400 * 96). GroupShape 자식은 슬라이드 절대좌표 → 컨테이너 Canvas에서 그룹 오프셋 차감.

### `SlideEditorCanvas`
편집 상호작용 전담. 마우스·키보드 이벤트를 처리하고 아래 이벤트로 MainWindow에 위임:

| 이벤트 | 의미 |
|--------|------|
| `TextCommitted` | 텍스트 편집 완료 |
| `ShapeMoved / ShapeResized / ShapeDeleted` | 도형 변형 |
| `ShapePropertiesRequested` | 개체 속성 대화상자 요청 |
| `ShapeDrawn` | 도형 그리기 완료 |
| `ShapesGroupRequested / ShapeUngroupRequested` | 그룹화 요청 |
| `CharPropertiesRequested` | 문자 속성 대화상자 요청 |

---

## 중요 제약사항

### OpenXml SDK v3 구조체 열거형
SDK v3에서 `ShapeTypeValues`, `PresetLineDashValues`, `TextAlignmentTypeValues` 등은 **struct 타입**. `switch` 패턴 매칭 불가, `==` 비교 또는 `.InnerText` 문자열 비교 사용.

```csharp
// 오류: switch case 불가
// switch (val) { case A.PresetLineDashValues.Dash: ... }

// 정상:
if (val == A.PresetLineDashValues.Dash) { ... }
// 또는 string 비교:
if (val?.InnerText == "dash") { ... }
```

존재하지 않는 정적 속성은 직접 생성:
```csharp
new A.ShapeTypeValues("isoTri")   // IsoscelesTriangle
```

### 네임스페이스 충돌 주의
`using DocumentFormat.OpenXml.Presentation;`을 추가할 경우 `TextElement`, `RgbColor` 등이 WPF/Models 타입과 충돌함. 별칭 사용:
```csharp
using WpfTE    = System.Windows.Documents.TextElement;
using RgbColor = PPEditer.Models.RgbColor;
```

### 좌표 단위
- **EMU** (English Metric Units): PPTX 내부 단위. 1인치 = 914400 EMU
- **WPF 픽셀**: `EmuToPx(emu) = emu / 914400.0 * 96.0`
- 슬라이드 기본 크기: 9144000 × 5143500 EMU (가로 10인치 × 세로 7.5인치 미만)

### XAML DynamicResource
툴바·버튼 텍스트는 반드시 `Content="{DynamicResource Tb_*}"` 속성 형식 사용.  
태그 사이 `{DynamicResource ...}` 텍스트는 리터럴로 출력됨.

---

## 다국어 문자열

`Resources/Strings.ko.xaml` / `Strings.en.xaml` 두 파일에 동일 Key로 등록.  
코드에서 접근: `Application.Current.TryFindResource(key) is string s ? s : fallback`  
MainWindow에서는 `private string S(string key, string fallback = "")` 헬퍼 사용.
