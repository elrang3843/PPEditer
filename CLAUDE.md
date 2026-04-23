# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run

Requires .NET 8 SDK and Windows (WPF is Windows-only).

```bash
cd PPEditer
dotnet build
dotnet run                   # blank window
dotnet run -- file.pptx      # open a file on startup
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

There are no automated tests.

## Architecture

The app is a WPF PPTX editor. Key namespaces and their roles:

```
PPEditer.MainWindow               ← top-level shell, owns all event wiring
PPEditer.Controls
  SlideEditorCanvas               ← interactive editor (~1550 lines), fires events upward
  SlideThumbnailPanel             ← slide list sidebar
PPEditer.Models
  PresentationModel               ← in-memory PPTX document, all mutations go here (~2100 lines)
  CharStyle / ParagraphStyle / ShapeStyle / DrawTool        ← data/enum types
  TransitionKind / AnimationKind / WatermarkKind            ← enum types
PPEditer.Services
  PptxConverter                   ← FlowDocument ↔ OpenXml TextBody (bidirectional)
  AppSettings                     ← singleton; persists config to settings.json
  FontService                     ← enumerates installed fonts, fsType licensing check
  ScreenHelper / MonitorInfo      ← multi-monitor enumeration for Presenter View
PPEditer.Rendering
  SlideRenderer                   ← SlidePart → WPF Canvas (read-only)
  WatermarkRenderer               ← draws watermark overlay
PPEditer.Export
  PdfExporter                     ← PDFsharp-based PDF export
  PrintExporter                   ← WPF PrintDialog-based print
PPEditer.Dialogs                  ← all modal windows (CharProperties, ParaProperties,
                                     ShapeProperties, ColorPickerPopup, AnimationDialog,
                                     TransitionDialog, SlideShowWindow, PresenterViewWindow,
                                     DocInfoDialog, DisplaySettingsDialog, UserManualDialog,
                                     LicenseDialog, AboutDialog, EmojiPickerDialog,
                                     MathSymbolDialog, CharMapDialog, PrintLayoutDialog)
```

### Data flow for a text edit

1. Double-click on slide → `SlideEditorCanvas.StartEdit()` calls `PptxConverter.ToFlowDocument()`.
2. `ScaleFlowDoc(doc, _editorScale)` scales all font sizes up so text appears correctly sized outside the Viewbox.
3. A `RichTextBox` is placed on `EditorOverlay` — a `Canvas` that is a **sibling of the Viewbox** (outside it), so the 1 px caret renders at integer screen pixels. Position computed via `TranslatePoint`.
4. Escape/blur fires `CommitEdit(save:true)`.
5. `ScaleFlowDoc(doc, 1.0/_editorScale)` reverses the scale, then `PptxConverter.FromFlowDocument()` converts back.
6. `TextCommitted` event → `MainWindow` calls `PresentationModel.UpdateShapeContent()` → mutates OpenXml → `slidePart.Slide.Save()`.
7. `EditorCanvas.Invalidate(treeIdx)` rebuilds the canvas child for that shape.

**`_editorScale`** = `(screen pixel width of shape) / (native canvas width of shape)`, i.e. the Viewbox scale factor. `ScaleFlowDoc` walks every `Paragraph` and `Run` multiplying `FontSize`, `LineHeight`, `Margin`, and `TextIndent`.

**`SuppressLostFocusCommit`** flag: set `true` before opening any dialog from within edit mode to prevent the `LostFocus` handler from committing early. The handler also skips commit when focus moves to a `MenuItem` or `MenuBase` (insert menu scenario).

**Stale canvas after CommitEdit**: `CommitEdit` may call `Rebuild()` which replaces `_nativeCanvas`. In `Canvas_MouseDown`, check whether `_nativeCanvas` changed after `CommitEdit` and switch to the new canvas before hit-testing.

### Undo/redo

`PresentationModel` keeps two `Stack<byte[]>` (max 50 entries). Every mutating call starts with `PushUndo()`, which serializes the entire `MemoryStream` to a `byte[]`. `Undo()` / `Redo()` swap stacks and call `PresentationDocument.Open(memoryStream)` to restore state.

### Shape selection and clipboard

`SlideEditorCanvas` tracks a primary `_selectedIdx` (tree index) and `_multiTreeIdxs` for multi-select (rubber-band drag). The static `_shapeClipboard` (`List<OpenXmlCompositeElement>`) is shared across instances.

Paste: `CloneNode(true)` each element, assign a new unique NV ID via `GetMaxNvId`, shift `A.Offset` by +360 000 EMU (~1 cm), `AppendChild` to the shape tree. `ShapesPasteRequested` carries the cloned elements to `MainWindow` → `PresentationModel.PasteShapes`.

### SlideEditorCanvas events (fired upward to MainWindow)

| Event | Payload |
|-------|---------|
| `TextCommitted` | slideIdx, treeIdx, `PptxParagraph[]` |
| `ShapeMoved / ShapeResized` | slideIdx, treeIdx, new EMU rect |
| `ShapeDeleted` | slideIdx, treeIdx |
| `ShapesDeleted` | slideIdx, `int[]` treeIdxs |
| `ShapeDrawn` | slideIdx, `OpenXmlCompositeElement` |
| `ShapesPasteRequested` | slideIdx, `OpenXmlCompositeElement[]` |
| `ShapeRotated` | slideIdx, treeIdx, angle |
| `ShapeOrderChanged` | slideIdx, treeIdx, direction |
| `ShapesGroupRequested / ShapeUngroupRequested` | slideIdx, treeIdx |
| `CharPropertiesRequested / ParaPropertiesRequested` | (none) |
| `ShapePropertiesRequested` | slideIdx, treeIdx |

### Unit system

All PPTX coordinates and sizes are in EMU (English Metric Units).

```
1 inch    = 914 400 EMU
1 cm      = 360 000 EMU
1 pt      = 12 700 EMU
Font size = hundredths of a point  (e.g. 1800 = 18 pt)
Line spc  = thousandths of a percent (e.g. 100 000 = 100 %)
Space bef/aft = hundredths of a point

WPF pixel = EMU / 914400 * 96
```

`SlideRenderer` computes `EmuToPx` dynamically from `_model.SlideWidth / NativeW`. `PptxConverter` and `SlideRenderer` define `EmuPerInch = 914400`, `WpfDpi = 96`, `PointPerInch = 72` as local constants.

### Paragraph/run property inheritance in FlowDocument

- `ToFlowDocument` reads `a:pPr/a:defRPr` and sets properties on the `Paragraph` so new `Run`s inherit them without explicit `a:rPr`.
- `FromFlowDocument` uses `ReadLocalValue(DependencyProperty)` to write only locally-set values as explicit `a:rPr`. This prevents polluting the PPTX with spurious defaults.

## OpenXML SDK v3 Constraints

SDK v3.2.0 changed several types from enums to structs. These patterns appear throughout the codebase:

| Type | Issue | Workaround |
|------|-------|-----------|
| `TextAlignmentTypeValues` | Struct — no `InnerText`, can't use `switch` | Keep both `avEnum = pPr?.Alignment` (for `.InnerText`) and `av = avEnum?.Value` (for `==`) |
| `TextAnchoringTypeValues` | Struct | Use `==` not `switch` |
| `SchemeColorValues` | Struct | Use `==` not `switch` |
| `TextUnderlineValues` | `EnumValue<T>`, not bool | Compare via `.InnerText` (`"none"` = no underline) |
| `ModifyVerifier` | Not typed in SDK v3 | Use `OpenXmlUnknownElement` + `SetAttribute` |
| Color scheme index attrs | Inaccessible typed | Set as raw `OpenXmlAttribute` |

When adding new OpenXml property reads, check whether the type is a struct before writing a `switch` expression.

## Namespace Collision Warning

`using DocumentFormat.OpenXml.Presentation;` introduces `TextElement` and other names that collide with WPF and project types. Use fully-qualified names or aliases:

```csharp
var fontSizeProp = System.Windows.Documents.TextElement.FontSizeProperty;
```

## Localization & Theming

String resources: `Resources/Strings.ko.xaml` and `Strings.en.xaml`. Add both when introducing user-visible strings. Access via `Application.Current.TryFindResource("Key") as string`. `MainWindow` has a private helper `string S(string key, string fallback = "")`.

Themes: `Theme.Light.xaml` / `Theme.Dark.xaml`, switched at runtime by `AppSettings.ApplyTheme()`.

Toolbar/button text in XAML must use the `Content="{DynamicResource Tb_*}"` **attribute** form — inline `{DynamicResource ...}` text between tags outputs the literal string.

## EMU Range Validation

Certain PPTX files contain out-of-range EMU values (e.g. `marL = Int32.MinValue`). Before converting EMU paragraph margins to WPF `Thickness`, validate:
- `marL`: clamp to `0..51 206 400`
- `indent`: clamp to `±51 206 400`
- `spcBef` / `spcAft` pixels: only apply if `>= 0`

This prevents `ArgumentException: "X,0,0,0" is not a valid Margin value`.
