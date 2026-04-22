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

The app is a WPF PPTX editor with a clean layered architecture:

```
MainWindow (XAML/code-behind)
  ├── SlideEditorCanvas   ← interactive editor, fires events upward
  ├── SlideThumbnailPanel ← slide list sidebar
  └── Dialogs/            ← char/para/shape/doc properties, color picker, charmap
        ↓ (events)
PresentationModel         ← in-memory PPTX wrapper, owns OpenXml document
        ↕
PptxConverter             ← FlowDocument ↔ OpenXml TextBody (bidirectional)
SlideRenderer             ← SlidePart → WPF Canvas (read-only render)
PdfExporter               ← renders slides to bitmap, writes PDF
```

### Data flow for a text edit

1. Double-click on slide → `SlideEditorCanvas.StartEdit()` calls `PptxConverter.ToFlowDocument()`, places a `RichTextBox` over the shape.
2. User edits; Escape/blur fires `CommitEdit(save:true)`.
3. `CommitEdit` calls `PptxConverter.FromFlowDocument()` and fires `TextCommitted` event.
4. `MainWindow.OnEditorTextCommitted` calls `PresentationModel.UpdateShapeContent()` → mutates OpenXml tree, calls `slidePart.Slide.Save()`.
5. `EditorCanvas.Invalidate(treeIdx)` rebuilds the canvas child for that shape.

### Undo/redo

`PresentationModel` keeps two `Stack<byte[]>` (max 50 entries). Every mutating call starts with `PushUndo()`, which serializes the entire MemoryStream to a byte array. `Undo()` / `Redo()` swap stacks and call `PresentationDocument.Open(memoryStream)` to restore state.

### Unit system

All PPTX coordinates and sizes are in EMU (English Metric Units).

```
1 inch    = 914 400 EMU
1 cm      = 360 000 EMU
1 pt      = 12 700 EMU
Font size = hundredths of a point (e.g., 1800 = 18 pt)
Line spc  = thousandths of a percent (e.g., 100 000 = 100%)
Space bef/aft = hundredths of a point

WPF pixel = EMU / 914400 * 96
```

`PptxConverter` and `SlideRenderer` both define `EmuPerInch = 914400`, `WpfDpi = 96`, `PointPerInch = 72` as local constants.

### Paragraph/run property inheritance in FlowDocument

- `ToFlowDocument` reads `a:pPr/a:defRPr` (paragraph default run props) and sets them on the `Paragraph` element so new `Run`s inherit them without explicit `a:rPr`.
- `FromFlowDocument` uses `ReadLocalValue(DependencyProperty)` to distinguish locally-set values from inherited ones. Only locally-set font family, size, weight, style, and foreground are written as explicit `a:rPr` attributes. This prevents polluting the PPTX with spurious defaults.

### `SlideEditorCanvas` key details

- Shapes rendered by `SlideRenderer` are tagged with their `ShapeTree` index (`fe.Tag = treeIndex`). This index is stable across canvas rebuilds and is the key used in all event payloads.
- During edit mode the original rendered shape is hidden (`Visibility.Hidden`), a `RichTextBox` is layered on top at `ZIndex=10000`, then restored on commit.
- `CommitEdit` nulls `_editor` **before** firing `TextCommitted` to prevent re-entrancy (the event handler calls `Invalidate` which would otherwise recurse into `CommitEdit`).

## OpenXML SDK v3 Constraints

SDK v3.2.0 changed several types from enums to structs. These patterns appear throughout the codebase:

| Type | Issue | Workaround |
|------|-------|-----------|
| `TextAlignmentTypeValues` | Struct — no `InnerText`, can't use `switch` | Keep both `avEnum = pPr?.Alignment` (for `.InnerText`) and `av = avEnum?.Value` (for `==` comparisons) |
| `TextAnchoringTypeValues` | Struct | Use `==` not `switch` |
| `SchemeColorValues` | Struct | Use `==` not `switch` |
| `TextUnderlineValues` | `EnumValue<T>`, not bool | Compare via `.InnerText` (`"none"` = no underline) |
| `ModifyVerifier` | Not typed in SDK v3 | Use `OpenXmlUnknownElement` with manual `SetAttribute` |
| Color scheme index attrs | Inaccessible typed | Set as raw `OpenXmlAttribute` |

When adding new OpenXml property reads, check whether the type is a struct before writing a `switch` expression.

## Localization & Theming

String resources live in `Resources/Strings.ko.xaml` and `Strings.en.xaml`. Add both when introducing user-visible strings. Access via `Application.Current.TryFindResource("Key") as string`.

Themes are in `Theme.Light.xaml` / `Theme.Dark.xaml` and are switched at runtime by `AppSettings.ApplyTheme()` which replaces the merged dictionary entry.

## EMU Range Validation

Certain PPTX files contain out-of-range EMU values (e.g., `marL = Int32.MinValue`). Before converting EMU paragraph margins to WPF `Thickness`, validate:
- `marL`: clamp to `0..51 206 400`
- `indent`: clamp to `±51 206 400`
- `spcBef` / `spcAft` pixels: only apply if `>= 0`

This prevents `ArgumentException: "X,0,0,0" is not a valid Margin value`.
