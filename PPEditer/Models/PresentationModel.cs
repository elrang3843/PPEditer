using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

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

    public bool HasWriteProtection
        => _doc?.PresentationPart?.Presentation
               .ChildElements.Any(e => e.LocalName == "modifyVerifier") == true;

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

        UpdateSaveMetadata();
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
        var masterPart = pp.SlideMasterParts.FirstOrDefault();
        var layout = masterPart?.SlideLayoutParts.Skip(1).FirstOrDefault()
                  ?? masterPart?.SlideLayoutParts.FirstOrDefault();
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
        if (layout is not null) slidePart.AddPart(layout);

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

    // ── Shape content ───────────────────────────────────────────────────

    /// <summary>
    /// Replace the text content of a shape identified by its index in the slide's ShapeTree.
    /// Preserves BodyProperties and ListStyle; rebuilds paragraphs from <paramref name="paragraphs"/>.
    /// </summary>
    public void UpdateShapeContent(int slideIndex, int shapeTreeIndex,
                                    IReadOnlyList<Services.PptxConverter.PptxParagraph> paragraphs)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return;

        var elements = slidePart.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        if (elements is null || shapeTreeIndex >= elements.Count) return;
        if (elements[shapeTreeIndex] is not Shape shape) return;
        if (shape.TextBody is null) return;

        PushUndo();

        // Preserve structural children
        var bodyProps  = shape.TextBody.GetFirstChild<A.BodyProperties>()?.CloneNode(true);
        var listStyle  = shape.TextBody.GetFirstChild<A.ListStyle>()?.CloneNode(true);
        shape.TextBody.RemoveAllChildren();
        if (bodyProps is not null) shape.TextBody.Append(bodyProps);
        if (listStyle is not null) shape.TextBody.Append(listStyle);

        foreach (var para in paragraphs)
        {
            var aPara = new A.Paragraph();

            // Paragraph alignment
            var alignVal = para.Alignment switch
            {
                System.Windows.TextAlignment.Center => A.TextAlignmentTypeValues.Center,
                System.Windows.TextAlignment.Right  => A.TextAlignmentTypeValues.Right,
                _                                   => A.TextAlignmentTypeValues.Left,
            };
            aPara.ParagraphProperties = new A.ParagraphProperties { Alignment = alignVal };

            foreach (var run in para.Runs)
            {
                var aRun  = new A.Run();
                var rProps = new A.RunProperties { Language = "ko-KR" };

                if (!string.IsNullOrEmpty(run.FontFamily))
                {
                    rProps.Append(new A.LatinFont    { Typeface = run.FontFamily });
                    rProps.Append(new A.EastAsianFont { Typeface = run.FontFamily });
                }
                if (run.FontSizePt.HasValue)
                    rProps.FontSize = (int)Math.Round(run.FontSizePt.Value * 100);
                if (run.Bold)          rProps.Bold      = true;
                if (run.Italic)        rProps.Italic    = true;
                if (run.Underline)     rProps.Underline = A.TextUnderlineValues.Single;
                if (run.Color.HasValue)
                    rProps.Append(new A.SolidFill(new A.RgbColorModelHex
                    {
                        Val = $"{run.Color.Value.R:X2}{run.Color.Value.G:X2}{run.Color.Value.B:X2}"
                    }));

                aRun.RunProperties = rProps;
                aRun.Text = new A.Text(run.Text);
                aPara.Append(aRun);
            }

            if (!para.Runs.Any())
                aPara.Append(new A.EndParagraphRunProperties { Language = "ko-KR" });

            shape.TextBody.Append(aPara);
        }

        slidePart.Slide.Save();
        _modified = true;
    }

    /// <summary>Add a new text box to the slide at the specified position (in EMU).</summary>
    public int AddTextBox(int slideIndex, long leftEmu, long topEmu,
                           long widthEmu, long heightEmu)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return -1;

        PushUndo();

        var tree  = slidePart.Slide.CommonSlideData!.ShapeTree!;
        uint maxId = tree.Elements<Shape>()
            .Select(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0u)
            .DefaultIfEmpty(1u).Max();

        var newShape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = maxId + 1, Name = $"TextBox {maxId + 1}" },
                new NonVisualShapeDrawingProperties { TextBox = true },
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset  { X = leftEmu, Y = topEmu },
                    new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                new A.PresetGeometry(new A.AdjustValueList())
                    { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill(),
                new A.Outline(new A.NoFill())),
            new TextBody(
                new A.BodyProperties
                {
                    Wrap        = A.TextWrappingValues.Square,
                    LeftInset   = 91440,
                    TopInset    = 45720,
                    RightInset  = 91440,
                    BottomInset = 45720,
                },
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(
                        new A.RunProperties { Language = "ko-KR", FontSize = 1800 },
                        new A.Text("텍스트를 입력하세요")))));

        tree.Append(newShape);
        slidePart.Slide.Save();
        _modified = true;

        // Return the index of the new element in the ShapeTree
        return tree.Elements<OpenXmlCompositeElement>().ToList().IndexOf(newShape);
    }

    /// <summary>Move a shape by <paramref name="deltaXEmu"/> / <paramref name="deltaYEmu"/> EMU.</summary>
    public void MoveShape(int slideIndex, int shapeTreeIndex, long deltaXEmu, long deltaYEmu)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return;

        var elements = slidePart.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        if (elements is null || shapeTreeIndex < 0 || shapeTreeIndex >= elements.Count) return;

        A.Transform2D? xfrm = null;
        if (elements[shapeTreeIndex] is Shape s)
            xfrm = s.ShapeProperties?.GetFirstChild<A.Transform2D>();
        else if (elements[shapeTreeIndex] is Picture p)
            xfrm = p.ShapeProperties?.GetFirstChild<A.Transform2D>();

        if (xfrm?.Offset is null) return;

        PushUndo();
        xfrm.Offset.X = (xfrm.Offset.X?.Value ?? 0L) + deltaXEmu;
        xfrm.Offset.Y = (xfrm.Offset.Y?.Value ?? 0L) + deltaYEmu;
        slidePart.Slide.Save();
        _modified = true;
    }

    /// <summary>Get position and size of a shape in EMU. Returns null if the shape has no transform.</summary>
    public (long x, long y, long cx, long cy)? GetShapeTransform(int slideIndex, int shapeTreeIndex)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return null;

        var elements = slidePart.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        if (elements is null || shapeTreeIndex < 0 || shapeTreeIndex >= elements.Count) return null;

        A.Transform2D? xfrm = elements[shapeTreeIndex] switch
        {
            Shape s   => s.ShapeProperties?.GetFirstChild<A.Transform2D>(),
            Picture p => p.ShapeProperties?.GetFirstChild<A.Transform2D>(),
            _         => null,
        };
        if (xfrm?.Offset is null || xfrm.Extents is null) return null;
        return (xfrm.Offset.X?.Value ?? 0L, xfrm.Offset.Y?.Value ?? 0L,
                xfrm.Extents.Cx?.Value ?? 0L, xfrm.Extents.Cy?.Value ?? 0L);
    }

    /// <summary>Move a shape one position toward the front of the shape tree. Returns new tree index.</summary>
    public int BringShapeForward(int slideIndex, int shapeTreeIndex)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return shapeTreeIndex;

        var tree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (tree is null) return shapeTreeIndex;

        var elements = tree.Elements<OpenXmlCompositeElement>().ToList();
        if (shapeTreeIndex < 0 || shapeTreeIndex >= elements.Count - 1) return shapeTreeIndex;

        PushUndo();
        var elem = elements[shapeTreeIndex];
        var next = elements[shapeTreeIndex + 1];
        elem.Remove();
        next.InsertAfterSelf(elem);
        slidePart.Slide.Save();
        _modified = true;
        return shapeTreeIndex + 1;
    }

    /// <summary>Move a shape one position toward the back of the shape tree. Returns new tree index.</summary>
    public int SendShapeBackward(int slideIndex, int shapeTreeIndex)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return shapeTreeIndex;

        var tree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (tree is null) return shapeTreeIndex;

        var elements = tree.Elements<OpenXmlCompositeElement>().ToList();
        if (shapeTreeIndex <= 0) return shapeTreeIndex;

        var prevElem = elements[shapeTreeIndex - 1];
        if (prevElem is not (Shape or Picture)) return shapeTreeIndex; // don't swap with structural elements

        PushUndo();
        var elem = elements[shapeTreeIndex];
        elem.Remove();
        prevElem.InsertBeforeSelf(elem);
        slidePart.Slide.Save();
        _modified = true;
        return shapeTreeIndex - 1;
    }

    /// <summary>Delete a shape identified by its shape-tree index.</summary>
    public void DeleteShape(int slideIndex, int shapeTreeIndex)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return;

        var elements = slidePart.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        if (elements is null || shapeTreeIndex < 0 || shapeTreeIndex >= elements.Count) return;

        PushUndo();
        elements[shapeTreeIndex].Remove();
        slidePart.Slide.Save();
        _modified = true;
    }

    /// <summary>Resize / reposition a shape to the given EMU coordinates.</summary>
    public void ResizeShape(int slideIndex, int shapeTreeIndex,
        long leftEmu, long topEmu, long widthEmu, long heightEmu)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return;

        var elements = slidePart.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        if (elements is null || shapeTreeIndex < 0 || shapeTreeIndex >= elements.Count) return;

        A.Transform2D? xfrm = null;
        if (elements[shapeTreeIndex] is Shape s)
            xfrm = s.ShapeProperties?.GetFirstChild<A.Transform2D>();
        else if (elements[shapeTreeIndex] is Picture p)
            xfrm = p.ShapeProperties?.GetFirstChild<A.Transform2D>();

        if (xfrm is null) return;

        PushUndo();
        if (xfrm.Offset  is null) xfrm.Offset  = new A.Offset();
        if (xfrm.Extents is null) xfrm.Extents = new A.Extents();
        xfrm.Offset.X  = leftEmu;
        xfrm.Offset.Y  = topEmu;
        xfrm.Extents.Cx = widthEmu;
        xfrm.Extents.Cy = heightEmu;
        slidePart.Slide.Save();
        _modified = true;
    }

    /// <summary>Embed an image file into the slide as a Picture element.</summary>
    public int AddImage(int slideIndex, string imagePath)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return -1;

        PushUndo();

        var ext = Path.GetExtension(imagePath).ToLowerInvariant();
        var imgType = ext switch
        {
            ".png"         => ImagePartType.Png,
            ".gif"         => ImagePartType.Gif,
            ".bmp"         => ImagePartType.Bmp,
            ".tiff" or ".tif" => ImagePartType.Tiff,
            _              => ImagePartType.Jpeg,
        };

        var imgPart = slidePart.AddImagePart(imgType);
        using (var fs = File.OpenRead(imagePath))
            imgPart.FeedData(fs);

        string rId = slidePart.GetIdOfPart(imgPart);

        // Default size: 4 × 3 inches centred on slide
        long imgW = Math.Min(3048000L, SlideWidth  / 2);
        long imgH = Math.Min(2286000L, SlideHeight / 2);
        long left = (SlideWidth  - imgW) / 2;
        long top  = (SlideHeight - imgH) / 2;

        var tree  = slidePart.Slide.CommonSlideData!.ShapeTree!;
        uint maxId = GetMaxShapeId(tree);

        var pic = new Picture(
            new NonVisualPictureProperties(
                new NonVisualDrawingProperties { Id = maxId + 1, Name = Path.GetFileName(imagePath) },
                new NonVisualPictureDrawingProperties(
                    new A.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties()),
            new BlipFill(
                new A.Blip { Embed = rId },
                new A.Stretch(new A.FillRectangle())),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset  { X = left, Y = top  },
                    new A.Extents { Cx = imgW, Cy = imgH }),
                new A.PresetGeometry(new A.AdjustValueList())
                    { Preset = A.ShapeTypeValues.Rectangle }));

        tree.Append(pic);
        slidePart.Slide.Save();
        _modified = true;
        return tree.Elements<OpenXmlCompositeElement>().ToList().IndexOf(pic);
    }

    /// <summary>Add a linked video placeholder shape.</summary>
    public int AddVideo(int slideIndex, string videoPath)
        => AddMediaShape(slideIndex, videoPath, "▶ Video", "005f87");

    /// <summary>Add a linked audio placeholder shape.</summary>
    public int AddAudio(int slideIndex, string audioPath)
        => AddMediaShape(slideIndex, audioPath, "🔊 Audio", "1b4f72");

    private int AddMediaShape(int slideIndex, string mediaPath, string labelPrefix, string hexColor)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return -1;

        PushUndo();

        long cx   = Math.Min(4572000L, SlideWidth  / 2);
        long cy   = Math.Min(914400L,  SlideHeight / 4);
        long left = (SlideWidth  - cx) / 2;
        long top  = (SlideHeight - cy) / 2;

        var tree  = slidePart.Slide.CommonSlideData!.ShapeTree!;
        uint maxId = GetMaxShapeId(tree);

        string fileName = Path.GetFileName(mediaPath);
        string label    = $"{labelPrefix}: {fileName}";

        var rProps = new A.RunProperties { Language = "ko-KR", FontSize = 1400 };
        rProps.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }));

        var shape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = maxId + 1, Name = $"Media{maxId + 1}" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset  { X = left, Y = top },
                    new A.Extents { Cx = cx,  Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList())
                    { Preset = A.ShapeTypeValues.Rectangle },
                new A.SolidFill(new A.RgbColorModelHex { Val = hexColor }),
                new A.Outline(new A.NoFill())),
            new TextBody(
                new A.BodyProperties
                {
                    Anchor    = A.TextAnchoringTypeValues.Center,
                    LeftInset = 91440, RightInset = 91440,
                    TopInset  = 45720, BottomInset = 45720,
                },
                new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Center },
                    new A.Run(rProps, new A.Text(label)))));

        tree.Append(shape);
        slidePart.Slide.Save();
        _modified = true;
        return tree.Elements<OpenXmlCompositeElement>().ToList().IndexOf(shape);
    }

    private static uint GetMaxShapeId(ShapeTree tree)
    {
        uint max = 1u;
        foreach (var el in tree.Elements<OpenXmlCompositeElement>())
        {
            uint id = el switch
            {
                Shape   s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0u,
                Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0u,
                _         => 0u,
            };
            if (id > max) max = id;
        }
        return max;
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
            BuildDefaultColorMap(),
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
            new SlideSize { Cx = 9144000, Cy = 5143500 },
            new NotesSize  { Cx = 6858000, Cy = 9144000 },
            new SlideMasterIdList(
                new SlideMasterId { Id = 2048, RelationshipId = pp.GetIdOfPart(masterPart) }),
            new SlideIdList(
                new SlideId { Id = 256, RelationshipId = pp.GetIdOfPart(slidePart) }));
        pp.Presentation.Save();
        doc.Save();
    }

    // ColorSchemeIndexValues is inaccessible in SDK v3 via using — set as raw XML attributes instead
    private static ColorMap BuildDefaultColorMap()
    {
        var cm = new ColorMap();
        foreach (var (name, val) in new[]
        {
            ("bg1","lt1"), ("tx1","dk1"), ("bg2","lt2"), ("tx2","dk2"),
            ("accent1","accent1"), ("accent2","accent2"), ("accent3","accent3"),
            ("accent4","accent4"), ("accent5","accent5"), ("accent6","accent6"),
            ("hlink","hlink"), ("folHlink","folHlink"),
        })
            cm.SetAttribute(new OpenXmlAttribute(name, string.Empty, val));
        return cm;
    }

    // ── Document info ────────────────────────────────────────────────────

    public DocProperties GetDocProperties()
    {
        var p = new DocProperties();
        if (_doc is null) return p;

        // Core properties
        var corePart = _doc.CoreFilePropertiesPart;
        if (corePart is not null)
        {
            try
            {
                XNamespace cp      = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
                XNamespace dc      = "http://purl.org/dc/elements/1.1/";
                XNamespace dcterms = "http://purl.org/dc/terms/";

                var xdoc = XDocument.Load(corePart.GetStream());
                var root = xdoc.Root!;

                p.Title          = root.Element(dc      + "title")?.Value          ?? string.Empty;
                p.Subject        = root.Element(dc      + "subject")?.Value        ?? string.Empty;
                p.Author         = root.Element(dc      + "creator")?.Value        ?? string.Empty;
                p.LastModifiedBy = root.Element(cp      + "lastModifiedBy")?.Value ?? string.Empty;

                if (int.TryParse(root.Element(cp + "revision")?.Value, out int rev))
                    p.Revision = rev;

                if (DateTime.TryParse(root.Element(dcterms + "created")?.Value,
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.RoundtripKind, out var created))
                    p.Created = created;
                if (DateTime.TryParse(root.Element(dcterms + "modified")?.Value,
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.RoundtripKind, out var modified))
                    p.Modified = modified;
            }
            catch { }
        }

        // Extended properties (Manager, Company)
        var extPart = _doc.ExtendedFilePropertiesPart;
        if (extPart is not null)
        {
            try
            {
                XNamespace ep   = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
                var xdoc        = XDocument.Load(extPart.GetStream());
                var root        = xdoc.Root!;
                p.Manager       = root.Element(ep + "Manager")?.Value ?? string.Empty;
                p.Company       = root.Element(ep + "Company")?.Value ?? string.Empty;
            }
            catch { }
        }

        return p;
    }

    public void UpdateDocInfo(DocProperties props,
        bool setProtection, string? newPassword, bool removeProtection)
    {
        if (_doc is null) return;
        PushUndo();

        SaveCorePropertiesFull(props);
        SaveExtendedProperties(props);

        var presentation = _doc.PresentationPart?.Presentation;
        if (presentation is not null)
        {
            var existing = presentation.ChildElements
                .FirstOrDefault(e => e.LocalName == "modifyVerifier");
            if (removeProtection)
            {
                if (existing is not null)
                {
                    presentation.RemoveChild(existing);
                    presentation.Save();
                }
            }
            else if (setProtection && !string.IsNullOrEmpty(newPassword))
            {
                if (existing is not null) presentation.RemoveChild(existing);
                presentation.Append(BuildModifyVerifier(newPassword!));
                presentation.Save();
            }
        }

        _modified = true;
    }

    private void SaveCorePropertiesFull(DocProperties props)
    {
        XNamespace cp      = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        XNamespace dc      = "http://purl.org/dc/elements/1.1/";
        XNamespace dcterms = "http://purl.org/dc/terms/";
        XNamespace xsi     = "http://www.w3.org/2001/XMLSchema-instance";

        var corePart = _doc!.CoreFilePropertiesPart;
        XDocument xdoc;

        if (corePart is null)
        {
            corePart = _doc.AddCoreFilePropertiesPart();
            xdoc = new XDocument(
                new XElement(cp + "coreProperties",
                    new XAttribute(XNamespace.Xmlns + "cp",      cp.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "dc",      dc.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "dcterms", dcterms.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "xsi",     xsi.NamespaceName)));
        }
        else
        {
            try   { xdoc = XDocument.Load(corePart.GetStream()); }
            catch
            {
                xdoc = new XDocument(
                    new XElement(cp + "coreProperties",
                        new XAttribute(XNamespace.Xmlns + "cp",      cp.NamespaceName),
                        new XAttribute(XNamespace.Xmlns + "dc",      dc.NamespaceName),
                        new XAttribute(XNamespace.Xmlns + "dcterms", dcterms.NamespaceName),
                        new XAttribute(XNamespace.Xmlns + "xsi",     xsi.NamespaceName)));
            }
        }

        var root = xdoc.Root!;
        var now  = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");

        XmlSetOrCreate(root, dc      + "title",          props.Title);
        XmlSetOrCreate(root, dc      + "subject",        props.Subject);
        XmlSetOrCreate(root, dc      + "creator",        props.Author);
        XmlSetOrCreate(root, cp      + "lastModifiedBy", Environment.UserName);
        XmlSetOrCreate(root, dcterms + "modified",       now, xsi + "type", "dcterms:W3CDTF");

        if (root.Element(dcterms + "created") is null)
            XmlSetOrCreate(root, dcterms + "created", now, xsi + "type", "dcterms:W3CDTF");

        var revEl = root.Element(cp + "revision");
        int rev   = revEl is not null && int.TryParse(revEl.Value, out int r) ? r + 1 : 1;
        XmlSetOrCreate(root, cp + "revision", rev.ToString());

        try
        {
            using var ws = corePart.GetStream(FileMode.Create, FileAccess.Write);
            xdoc.Save(ws);
        }
        catch { }
    }

    private void SaveExtendedProperties(DocProperties props)
    {
        XNamespace ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

        var part = _doc!.ExtendedFilePropertiesPart;
        XDocument xdoc;

        if (part is null)
        {
            part  = _doc.AddExtendedFilePropertiesPart();
            xdoc  = new XDocument(
                new XElement(ep + "Properties",
                    new XAttribute("xmlns", ep.NamespaceName)));
        }
        else
        {
            try   { xdoc = XDocument.Load(part.GetStream()); }
            catch { xdoc = new XDocument(new XElement(ep + "Properties",
                        new XAttribute("xmlns", ep.NamespaceName))); }
        }

        var root = xdoc.Root!;

        XmlSetOrRemove(root, ep + "Manager", props.Manager);
        XmlSetOrRemove(root, ep + "Company", props.Company);

        try
        {
            using var ws = part.GetStream(FileMode.Create, FileAccess.Write);
            xdoc.Save(ws);
        }
        catch { }
    }

    private void UpdateSaveMetadata()
    {
        if (_doc is null) return;

        XNamespace cp      = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        XNamespace dc      = "http://purl.org/dc/elements/1.1/";
        XNamespace dcterms = "http://purl.org/dc/terms/";
        XNamespace xsi     = "http://www.w3.org/2001/XMLSchema-instance";

        var corePart = _doc.CoreFilePropertiesPart;
        XDocument xdoc;

        if (corePart is null)
        {
            try { corePart = _doc.AddCoreFilePropertiesPart(); }
            catch { return; }
            xdoc = new XDocument(
                new XElement(cp + "coreProperties",
                    new XAttribute(XNamespace.Xmlns + "cp",      cp.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "dc",      dc.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "dcterms", dcterms.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "xsi",     xsi.NamespaceName)));
        }
        else
        {
            try   { xdoc = XDocument.Load(corePart.GetStream()); }
            catch { return; }
        }

        var root = xdoc.Root!;
        var now  = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");

        XmlSetOrCreate(root, cp      + "lastModifiedBy", Environment.UserName);
        XmlSetOrCreate(root, dcterms + "modified",       now, xsi + "type", "dcterms:W3CDTF");

        if (root.Element(dcterms + "created") is null)
            XmlSetOrCreate(root, dcterms + "created", now, xsi + "type", "dcterms:W3CDTF");

        var revEl = root.Element(cp + "revision");
        int rev   = revEl is not null && int.TryParse(revEl.Value, out int r) ? r + 1 : 1;
        XmlSetOrCreate(root, cp + "revision", rev.ToString());

        try
        {
            using var ws = corePart.GetStream(FileMode.Create, FileAccess.Write);
            xdoc.Save(ws);
        }
        catch { }
    }

    // ── XML helpers ─────────────────────────────────────────────────────

    private static void XmlSetOrCreate(XElement root, XName name, string value,
        XName? attrName = null, string? attrValue = null)
    {
        var el = root.Element(name);
        if (el is null)
        {
            el = new XElement(name, value);
            if (attrName is not null && attrValue is not null)
                el.Add(new XAttribute(attrName, attrValue));
            root.Add(el);
        }
        else
        {
            el.Value = value;
        }
    }

    private static void XmlSetOrRemove(XElement root, XName name, string value)
    {
        var el = root.Element(name);
        if (!string.IsNullOrWhiteSpace(value))
        {
            if (el is null) root.Add(new XElement(name, value));
            else            el.Value = value;
        }
        else
        {
            el?.Remove();
        }
    }

    // ── Write protection ─────────────────────────────────────────────────

    private static OpenXmlElement BuildModifyVerifier(string password)
    {
        var (salt, hash) = ComputePasswordHash(password);
        // ModifyVerifier is not accessible as a typed class in SDK v3 — use OpenXmlUnknownElement
        const string pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
        var mv = new OpenXmlUnknownElement("p", "modifyVerifier", pNs);
        mv.SetAttribute(new OpenXmlAttribute("cryptProviderType",   string.Empty, "rsaAES"));
        mv.SetAttribute(new OpenXmlAttribute("cryptAlgorithmClass", string.Empty, "hash"));
        mv.SetAttribute(new OpenXmlAttribute("cryptAlgorithmType",  string.Empty, "typeAny"));
        mv.SetAttribute(new OpenXmlAttribute("cryptAlgorithmSid",   string.Empty, "12"));
        mv.SetAttribute(new OpenXmlAttribute("spinCount",           string.Empty, "100000"));
        mv.SetAttribute(new OpenXmlAttribute("saltValue",           string.Empty, salt));
        mv.SetAttribute(new OpenXmlAttribute("hashValue",           string.Empty, hash));
        return mv;
    }

    private static (string salt, string hash) ComputePasswordHash(string password,
        int spinCount = 100000)
    {
        var saltBytes = RandomNumberGenerator.GetBytes(16);
        var pwBytes   = Encoding.Unicode.GetBytes(password);    // UTF-16LE

        using var sha = SHA256.Create();

        var init = new byte[saltBytes.Length + pwBytes.Length];
        saltBytes.CopyTo(init, 0);
        pwBytes.CopyTo(init, saltBytes.Length);
        var h = sha.ComputeHash(init);

        for (int i = 0; i < spinCount; i++)
        {
            var iter = new byte[h.Length + 4];
            h.CopyTo(iter, 0);
            BitConverter.GetBytes(i).CopyTo(iter, h.Length);
            h = sha.ComputeHash(iter);
        }

        return (Convert.ToBase64String(saltBytes), Convert.ToBase64String(h));
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
