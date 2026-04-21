using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PPEditer.Models;

/// <summary>
/// Manages an open PPTX document in memory with undo/redo support.
/// </summary>
public sealed class PresentationModel : IDisposable
{
    private const int MaxUndo = 50;

    private PresentationDocument? _doc;
    private MemoryStream?         _stream;
    private string?               _filePath;
    private bool                  _modified;

    private readonly Stack<byte[]> _undoStack = new();
    private readonly Stack<byte[]> _redoStack = new();

    // ── Properties ──────────────────────────────────────────────────────

    public bool   IsOpen    => _doc != null;
    public bool   Modified  => _modified;
    public string FilePath  => _filePath ?? string.Empty;
    public string FileName  => string.IsNullOrEmpty(_filePath)
                               ? "제목없음.pptx"
                               : Path.GetFileName(_filePath);
    public string WindowTitle => $"{(_modified ? "* " : "")}{FileName} - PPEditer";

    public bool CanUndo => _undoStack.Count > 0;
    public bool CanRedo => _redoStack.Count > 0;

    public long SlideWidth  => _doc?.PresentationPart?.Presentation.SlideSize?.Cx ?? 9144000L;
    public long SlideHeight => _doc?.PresentationPart?.Presentation.SlideSize?.Cy ?? 5143500L;

    public int SlideCount
    {
        get
        {
            var list = _doc?.PresentationPart?.Presentation.SlideIdList;
            return list?.Count() ?? 0;
        }
    }

    // ── File operations ─────────────────────────────────────────────────

    public void New()
    {
        Dispose();
        _stream = new MemoryStream();
        CreateBlankPresentation(_stream);
        _stream.Position = 0;
        _doc      = PresentationDocument.Open(_stream, true);
        _filePath = null;
        _modified = false;
        _undoStack.Clear();
        _redoStack.Clear();
    }

    public void Open(string filePath)
    {
        Dispose();
        _stream = new MemoryStream();
        using (var fs = File.OpenRead(filePath))
            fs.CopyTo(_stream);
        _stream.Position = 0;
        _doc      = PresentationDocument.Open(_stream, true);
        _filePath = filePath;
        _modified = false;
        _undoStack.Clear();
        _redoStack.Clear();
    }

    public void Save(string? filePath = null)
    {
        if (_doc is null) return;
        if (filePath != null) _filePath = filePath;
        if (string.IsNullOrEmpty(_filePath))
            throw new InvalidOperationException("저장 경로가 없습니다.");

        _doc.Save();
        _stream!.Position = 0;
        using var fs = File.Create(_filePath);
        _stream.CopyTo(fs);
        _modified = false;
    }

    // ── Slide access ────────────────────────────────────────────────────

    /// <summary>Returns the SlidePart for slide at <paramref name="index"/> (0-based).</summary>
    public SlidePart? GetSlidePart(int index)
    {
        var pp   = _doc?.PresentationPart;
        var list = pp?.Presentation.SlideIdList?.Elements<SlideId>().ToList();
        if (list is null || index < 0 || index >= list.Count) return null;
        var rId = list[index].RelationshipId?.Value;
        return rId is not null ? pp!.GetPartById(rId) as SlidePart : null;
    }

    // ── Slide mutations ─────────────────────────────────────────────────

    public int AddSlide(int afterIndex = -1)
    {
        PushUndo();
        var pp     = _doc!.PresentationPart!;
        var layout = pp.SlideLayoutParts.Skip(1).FirstOrDefault()
                     ?? pp.SlideLayoutParts.First();
        var slidePart = pp.AddNewPart<SlidePart>();
        new Slide(
            new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(
                    new DocumentFormat.OpenXml.Drawing.TransformGroup()))),
            new ColorMapOverride(new DocumentFormat.OpenXml.Drawing.MasterColorMapping()))
            .Save(slidePart);
        slidePart.AddPart(layout);

        var idList = pp.Presentation.SlideIdList ?? (pp.Presentation.SlideIdList = new SlideIdList());
        uint maxId = idList.Elements<SlideId>().Select(s => s.Id?.Value ?? 0u).DefaultIfEmpty(255u).Max();
        var newSld = new SlideId { Id = maxId + 1, RelationshipId = pp.GetIdOfPart(slidePart) };

        int insertAt = afterIndex + 1;
        var allIds   = idList.Elements<SlideId>().ToList();
        if (insertAt >= allIds.Count)
            idList.Append(newSld);
        else
            idList.InsertBefore(newSld, allIds[insertAt]);

        pp.Presentation.Save();
        _modified = true;
        return Math.Min(insertAt, SlideCount - 1);
    }

    public bool DeleteSlide(int index)
    {
        if (SlideCount <= 1) return false;
        PushUndo();
        var pp    = _doc!.PresentationPart!;
        var list  = pp.Presentation.SlideIdList!;
        var ids   = list.Elements<SlideId>().ToList();
        if (index < 0 || index >= ids.Count) return false;
        var sldId = ids[index];
        var rId   = sldId.RelationshipId?.Value;
        list.RemoveChild(sldId);
        if (rId is not null)
        {
            var part = pp.GetPartById(rId);
            pp.DeletePart(part);
        }
        pp.Presentation.Save();
        _modified = true;
        return true;
    }

    public int DuplicateSlide(int index)
    {
        PushUndo();
        var src    = GetSlidePart(index);
        if (src is null) return index;

        var pp        = _doc!.PresentationPart!;
        var newPart   = pp.AddNewPart<SlidePart>();

        // Deep-copy slide XML
        using (var srcStream = src.GetStream(FileMode.Open, FileAccess.Read))
        using (var dstStream = newPart.GetStream(FileMode.Create))
            srcStream.CopyTo(dstStream);

        // Copy relationships (images, layout, etc.)
        foreach (var rel in src.Parts)
            newPart.AddPart(rel.OpenXmlPart, rel.RelationshipId);

        var idList = pp.Presentation.SlideIdList!;
        var ids    = idList.Elements<SlideId>().ToList();
        uint maxId = ids.Select(s => s.Id?.Value ?? 0u).DefaultIfEmpty(255u).Max();
        var newId  = new SlideId { Id = maxId + 1, RelationshipId = pp.GetIdOfPart(newPart) };
        if (index + 1 < ids.Count)
            idList.InsertAfter(newId, ids[index]);
        else
            idList.Append(newId);

        pp.Presentation.Save();
        _modified = true;
        return index + 1;
    }

    public void MoveSlide(int fromIndex, int toIndex)
    {
        if (fromIndex == toIndex) return;
        PushUndo();
        var pp    = _doc!.PresentationPart!;
        var list  = pp.Presentation.SlideIdList!;
        var ids   = list.Elements<SlideId>().ToList();
        if (fromIndex < 0 || fromIndex >= ids.Count) return;
        if (toIndex   < 0 || toIndex   >= ids.Count) return;

        var moved = ids[fromIndex];
        list.RemoveChild(moved);
        ids = list.Elements<SlideId>().ToList();

        if (toIndex == 0)
            list.PrependChild(moved);
        else if (toIndex >= ids.Count)
            list.AppendChild(moved);
        else
            list.InsertBefore(moved, ids[toIndex]);

        pp.Presentation.Save();
        _modified = true;
    }

    // ── Undo / Redo ─────────────────────────────────────────────────────

    public bool Undo()
    {
        if (_undoStack.Count == 0) return false;
        _redoStack.Push(Snapshot());
        RestoreSnapshot(_undoStack.Pop());
        return true;
    }

    public bool Redo()
    {
        if (_redoStack.Count == 0) return false;
        _undoStack.Push(Snapshot());
        RestoreSnapshot(_redoStack.Pop());
        return true;
    }

    private void PushUndo()
    {
        _redoStack.Clear();
        _undoStack.Push(Snapshot());
        if (_undoStack.Count > MaxUndo)
        {
            var arr   = _undoStack.ToArray();
            _undoStack.Clear();
            foreach (var s in arr.Take(MaxUndo)) _undoStack.Push(s);
        }
    }

    private byte[] Snapshot()
    {
        _doc!.Save();
        _stream!.Position = 0;
        return _stream.ToArray();
    }

    private void RestoreSnapshot(byte[] data)
    {
        var savedPath = _filePath;
        _doc?.Dispose();
        _stream?.Dispose();
        _stream          = new MemoryStream(data) { Position = 0 };
        _doc             = PresentationDocument.Open(_stream, true);
        _filePath        = savedPath;
        _modified        = true;
    }

    // ── Blank presentation factory ───────────────────────────────────────

    private static void CreateBlankPresentation(Stream stream)
    {
        using var doc  = PresentationDocument.Create(stream, DocumentFormat.OpenXml.PresentationDocumentType.Presentation, autoSave: false);
        var pp         = doc.AddPresentationPart();

        // Minimal slide master / layout chain
        var masterPart  = pp.AddNewPart<SlideMasterPart>();
        var layoutPart  = masterPart.AddNewPart<SlideLayoutPart>();
        var themePart   = masterPart.AddNewPart<ThemePart>();
        themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
            new DocumentFormat.OpenXml.Drawing.ThemeElements(
                new DocumentFormat.OpenXml.Drawing.FontScheme(
                    new DocumentFormat.OpenXml.Drawing.MajorFont(
                        new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "맑은 고딕" }),
                    new DocumentFormat.OpenXml.Drawing.MinorFont(
                        new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "맑은 고딕" }))
                { Name = "Office" },
                new DocumentFormat.OpenXml.Drawing.FormatScheme { Name = "Office" }))
            { Name = "Office Theme" };
        themePart.Theme.Save();

        layoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(
                    new DocumentFormat.OpenXml.Drawing.TransformGroup()))))
        { Type = SlideLayoutValues.Blank };
        layoutPart.SlideLayout.Save();
        layoutPart.AddPart(masterPart);

        masterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(
                    new DocumentFormat.OpenXml.Drawing.TransformGroup()))),
            new ColorMap
            {
                Background1 = ColorSchemeIndexValues.Light1,
                Text1        = ColorSchemeIndexValues.Dark1,
                Background2  = ColorSchemeIndexValues.Light2,
                Text2        = ColorSchemeIndexValues.Dark2,
                Accent1      = ColorSchemeIndexValues.Accent1,
                Accent2      = ColorSchemeIndexValues.Accent2,
                Accent3      = ColorSchemeIndexValues.Accent3,
                Accent4      = ColorSchemeIndexValues.Accent4,
                Accent5      = ColorSchemeIndexValues.Accent5,
                Accent6      = ColorSchemeIndexValues.Accent6,
                Hyperlink    = ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink
            },
            new SlideLayoutIdList(
                new SlideLayoutId { Id = 2049, RelationshipId = masterPart.GetIdOfPart(layoutPart) }));
        masterPart.SlideMaster.Save();

        // Blank slide
        var slidePart = pp.AddNewPart<SlidePart>();
        new Slide(
            new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(
                    new DocumentFormat.OpenXml.Drawing.TransformGroup()))),
            new ColorMapOverride(new DocumentFormat.OpenXml.Drawing.MasterColorMapping()))
            .Save(slidePart);
        slidePart.AddPart(layoutPart);

        // Presentation
        pp.Presentation = new Presentation(
            new SlideSize { Cx = 9144000L, Cy = 5143500L },
            new NotesSize  { Cx = 6858000L, Cy = 9144000L },
            new SlideMasterIdList(
                new SlideMasterId { Id = 2048, RelationshipId = pp.GetIdOfPart(masterPart) }),
            new SlideIdList(
                new SlideId { Id = 256, RelationshipId = pp.GetIdOfPart(slidePart) }));
        pp.Presentation.Save();
        doc.Save();
    }

    // ── IDisposable ─────────────────────────────────────────────────────

    public void Dispose()
    {
        _doc?.Dispose();
        _stream?.Dispose();
        _doc    = null;
        _stream = null;
    }
}
