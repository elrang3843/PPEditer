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

            // Paragraph-level properties
            var alignVal = para.Alignment switch
            {
                System.Windows.TextAlignment.Center  => A.TextAlignmentTypeValues.Center,
                System.Windows.TextAlignment.Right   => A.TextAlignmentTypeValues.Right,
                System.Windows.TextAlignment.Justify => new A.TextAlignmentTypeValues("just"),
                _                                    => A.TextAlignmentTypeValues.Left,
            };
            var pPr = new A.ParagraphProperties { Alignment = alignVal };
            if (para.MarginLeftEmu  != 0) pPr.LeftMargin = para.MarginLeftEmu;
            if (para.TextIndentEmu  != 0) pPr.Indent     = para.TextIndentEmu;
            if (para.LineSpacePct1000 > 0)
                pPr.Append(new A.LineSpacing(new A.SpacingPercent { Val = para.LineSpacePct1000 }));
            if (para.SpaceBeforePt100 > 0)
                pPr.Append(new A.SpaceBefore(new A.SpacingPoints { Val = para.SpaceBeforePt100 }));
            if (para.SpaceAfterPt100 > 0)
                pPr.Append(new A.SpaceAfter(new A.SpacingPoints { Val = para.SpaceAfterPt100 }));
            aPara.ParagraphProperties = pPr;

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

                // Extended character properties
                if (run.Strikethrough) rProps.Strike = A.TextStrikeValues.SingleStrike;
                if (run.Script == ScriptKind.Superscript) rProps.Baseline = 30000;
                else if (run.Script == ScriptKind.Subscript) rProps.Baseline = -25000;
                if (run.SpacingPt100 != 0) rProps.Spacing = run.SpacingPt100;
                if (run.BackColor.HasValue)
                    rProps.Append(new A.Highlight(new A.RgbColorModelHex
                        { Val = $"{run.BackColor.Value.R:X2}{run.BackColor.Value.G:X2}{run.BackColor.Value.B:X2}" }));
                if (run.UnderlineColor.HasValue)
                    rProps.Append(new A.UnderlineFill(new A.SolidFill(
                        new A.RgbColorModelHex { Val = $"{run.UnderlineColor.Value.R:X2}{run.UnderlineColor.Value.G:X2}{run.UnderlineColor.Value.B:X2}" })));
                if (run.HasOutline)
                    rProps.Append(new A.Outline(new A.SolidFill(
                        new A.RgbColorModelHex { Val = "000000" })));

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

    /// <summary>Update the vertical anchor of a text body shape.</summary>
    public void UpdateBodyVertAnchor(int slideIndex, int shapeTreeIndex, VertAnchor anchor)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return;
        var elements = slidePart.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        if (elements is null || shapeTreeIndex >= elements.Count) return;
        if (elements[shapeTreeIndex] is not Shape shape) return;
        if (shape.TextBody is null) return;

        PushUndo();

        var bodyPr = shape.TextBody.GetFirstChild<A.BodyProperties>();
        if (bodyPr is null) { bodyPr = new A.BodyProperties(); shape.TextBody.InsertAt(bodyPr, 0); }

        bodyPr.Anchor = anchor switch
        {
            VertAnchor.Middle => A.TextAnchoringTypeValues.Center,
            VertAnchor.Bottom => A.TextAnchoringTypeValues.Bottom,
            _                 => A.TextAnchoringTypeValues.Top,
        };

        slidePart.Slide.Save();
        _modified = true;
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

    /// <summary>Move a shape to the very front (top) of the shape tree. Returns new tree index.</summary>
    public int BringShapeToFront(int slideIndex, int shapeTreeIndex)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return shapeTreeIndex;

        var tree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (tree is null) return shapeTreeIndex;

        var elements = tree.Elements<OpenXmlCompositeElement>().ToList();
        int last = elements.Count - 1;
        if (shapeTreeIndex < 0 || shapeTreeIndex >= elements.Count || shapeTreeIndex == last)
            return shapeTreeIndex;

        PushUndo();
        var elem = elements[shapeTreeIndex];
        elem.Remove();
        tree.Append(elem);
        slidePart.Slide.Save();
        _modified = true;
        return last;
    }

    /// <summary>Move a shape to the very back (bottom content position) of the shape tree. Returns new tree index.</summary>
    public int SendShapeToBack(int slideIndex, int shapeTreeIndex)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return shapeTreeIndex;

        var tree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (tree is null) return shapeTreeIndex;

        var elements = tree.Elements<OpenXmlCompositeElement>().ToList();
        if (shapeTreeIndex < 0 || shapeTreeIndex >= elements.Count) return shapeTreeIndex;

        // Find the first actual content element — don't go before structural nodes
        int firstContent = elements.FindIndex(e => e is Shape or Picture or GroupShape);
        if (firstContent < 0 || shapeTreeIndex == firstContent) return shapeTreeIndex;

        PushUndo();
        var elem    = elements[shapeTreeIndex];
        var anchor  = elements[firstContent];
        elem.Remove();
        anchor.InsertBeforeSelf(elem);
        slidePart.Slide.Save();
        _modified = true;
        return firstContent;
    }

    // ── Shape drawing ────────────────────────────────────────────────────

    /// <summary>
    /// Add a new shape drawn by the user. xs/ys are EMU coordinates on the slide.
    /// For drag shapes: [0]=startX/Y, [1]=endX/Y.
    /// For click shapes: every vertex in order.
    /// </summary>
    public int AddDrawnShape(int slideIndex, DrawTool tool, long[] xs, long[] ys)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null || xs.Length < 2 || xs.Length != ys.Length) return -1;

        PushUndo();

        long left = xs.Min(), top = ys.Min();
        long w    = Math.Max(xs.Max() - left, 12700L);
        long h    = Math.Max(ys.Max() - top,  12700L);

        // Square / Circle → force equal sides
        if (tool is DrawTool.Square or DrawTool.EqTriangle or DrawTool.Circle)
        { long s = Math.Max(w, h); w = s; h = s; }

        long[] lx = xs.Select(x => x - left).ToArray();
        long[] ly = ys.Select(y => y - top).ToArray();

        var tree  = slidePart.Slide.CommonSlideData!.ShapeTree!;
        uint maxId = GetMaxShapeId(tree);

        bool isLineType = tool is DrawTool.Line or DrawTool.PolyLine or DrawTool.SplineLine;
        bool isFilled   = !isLineType;

        OpenXmlElement geom;
        A.Transform2D  xfrm;

        if (tool is DrawTool.Line)
        {
            xfrm = MakeXfrm(left, top, w, h);
            geom = BuildPolyPath(lx, ly, w, h, closed: false);
        }
        else if (tool is DrawTool.PolyLine)
        {
            xfrm = MakeXfrm(left, top, w, h);
            geom = BuildPolyPath(lx, ly, w, h, closed: false);
        }
        else if (tool is DrawTool.SplineLine)
        {
            xfrm = MakeXfrm(left, top, w, h);
            geom = BuildSplinePath(lx, ly, w, h, closed: false);
        }
        else if (tool is DrawTool.ScaleneTriangle or DrawTool.Polygon or DrawTool.Trapezoid)
        {
            // Trapezoid: user clicked 4 points; stored as a closed custom polygon
            xfrm = MakeXfrm(left, top, w, h);
            geom = BuildPolyPath(lx, ly, w, h, closed: true);
        }
        else if (tool is DrawTool.SplinePolygon)
        {
            xfrm = MakeXfrm(left, top, w, h);
            geom = BuildSplinePath(lx, ly, w, h, closed: true);
        }
        else if (tool is DrawTool.RightTriangle)
        {
            // Right angle is placed at the drag-start corner
            bool fH = xs[0] > xs[1]; // started from right side
            bool fV = ys[0] > ys[1]; // started from bottom
            long rax = fH ? w : 0L,  ray = fV ? h : 0L;  // right-angle vertex
            long p1x = fH ? w : 0L,  p1y = fV ? 0L : h;  // along the vertical leg
            long p2x = fH ? 0L : w,  p2y = fV ? h : 0L;  // along the horizontal leg
            xfrm = MakeXfrm(left, top, w, h);
            geom = BuildPolyPath([rax, p1x, p2x], [ray, p1y, p2y], w, h, closed: true);
        }
        else if (tool is DrawTool.Parallelogram)
        {
            // Lean direction follows horizontal drag direction
            bool fH = xs[0] > xs[1]; // started from right → lean left
            long off = w / 4;
            long[] px = fH ? [0L, w - off, w,   off]
                           : [off, w,      w - off, 0L];
            long[] py = [0L, 0L, h, h];
            xfrm = MakeXfrm(left, top, w, h);
            geom = BuildPolyPath(px, py, w, h, closed: true);
        }
        else if (tool is DrawTool.Arc)
        {
            // The arc bulges toward the drag-start corner
            bool fH = xs[0] > xs[1];
            bool fV = ys[0] > ys[1];
            // adj1 = start angle (60000ths of a degree); adj2 = swing = 90°
            long adj1 = (fH, fV) switch
            {
                (false, false) => 10800000L, // TL start → bulges TL (180°)
                (true,  false) => 16200000L, // TR start → bulges TR (270°, PPTX default)
                (false, true)  =>  5400000L, // BL start → bulges BL (90°)
                (true,  true)  =>        0L, // BR start → bulges BR (0°)
            };
            const long adj2 = 5400000L; // 90° sweep
            xfrm = MakeXfrm(left, top, w, h);
            var adjList = new A.AdjustValueList(
                new A.ShapeGuide { Name = "adj1", Formula = $"val {adj1}" },
                new A.ShapeGuide { Name = "adj2", Formula = $"val {adj2}" });
            geom = new A.PresetGeometry(adjList) { Preset = A.ShapeTypeValues.Arc };
        }
        else
        {
            xfrm = MakeXfrm(left, top, w, h);

            if (tool == DrawTool.IsoTriangle)
            {
                // "isoTri" is not a member of ShapeTypeValues in SDK v3 (struct validates known values);
                // set the prst attribute directly on the element to avoid ArgumentOutOfRangeException.
                var pg = new A.PresetGeometry(new A.AdjustValueList());
                pg.SetAttribute(new OpenXmlAttribute("prst", string.Empty, "isoTri"));
                geom = pg;
            }
            else
            {
                A.ShapeTypeValues preset = tool switch
                {
                    DrawTool.Rect or DrawTool.Square    => A.ShapeTypeValues.Rectangle,
                    DrawTool.Ellipse or DrawTool.Circle => A.ShapeTypeValues.Ellipse,
                    DrawTool.EqTriangle                 => A.ShapeTypeValues.Triangle,
                    DrawTool.Arrow                      => A.ShapeTypeValues.RightArrow,
                    _                                   => A.ShapeTypeValues.Rectangle,
                };
                geom = new A.PresetGeometry(new A.AdjustValueList()) { Preset = preset };
            }
        }

        var spPr = new ShapeProperties();
        spPr.Append(xfrm);
        spPr.Append(geom);

        if (isLineType)
            spPr.Append(new A.NoFill());
        else
            spPr.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "BDD7EE" }));

        var ol = new A.Outline { Width = 19050 }; // 1.5 pt
        ol.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "2E74B5" }));
        spPr.Append(ol);

        var shape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = maxId + 1, Name = $"Shape {maxId + 1}" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            spPr,
            new TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph()));

        tree.Append(shape);
        slidePart.Slide.Save();
        _modified = true;
        return tree.Elements<OpenXmlCompositeElement>().ToList().IndexOf(shape);
    }

    private static A.Transform2D MakeXfrm(long l, long t, long w, long h)
        => new A.Transform2D(
            new A.Offset  { X = l, Y = t },
            new A.Extents { Cx = w, Cy = h });

    private static A.CustomGeometry BuildPolyPath(long[] lx, long[] ly, long w, long h, bool closed)
    {
        var path = new A.Path
        {
            Width = w, Height = h,
            Fill  = closed ? A.PathFillModeValues.Norm : A.PathFillModeValues.None,
        };
        path.Append(new A.MoveTo(APt(lx[0], ly[0])));
        for (int i = 1; i < lx.Length; i++)
            path.Append(new A.LineTo(APt(lx[i], ly[i])));
        if (closed) path.Append(new A.CloseShapePath());
        return WrapInCustGeom(path);
    }

    private static A.CustomGeometry BuildSplinePath(long[] lx, long[] ly, long w, long h, bool closed)
    {
        int n = lx.Length;
        var path = new A.Path
        {
            Width = w, Height = h,
            Fill  = closed ? A.PathFillModeValues.Norm : A.PathFillModeValues.None,
        };
        path.Append(new A.MoveTo(APt(lx[0], ly[0])));
        for (int i = 0; i < n - 1; i++)
        {
            long p0x = lx[Math.Max(i - 1, 0)], p0y = ly[Math.Max(i - 1, 0)];
            long p1x = lx[i],                   p1y = ly[i];
            long p2x = lx[i + 1],               p2y = ly[i + 1];
            long p3x = lx[Math.Min(i + 2, n-1)],p3y = ly[Math.Min(i + 2, n-1)];
            long c1x = p1x + (p2x - p0x) / 6,   c1y = p1y + (p2y - p0y) / 6;
            long c2x = p2x - (p3x - p1x) / 6,   c2y = p2y - (p3y - p1y) / 6;
            path.Append(new A.CubicBezierCurveTo(APt(c1x, c1y), APt(c2x, c2y), APt(p2x, p2y)));
        }
        if (closed) path.Append(new A.CloseShapePath());
        return WrapInCustGeom(path);
    }

    private static A.Point APt(long x, long y) => new A.Point { X = x.ToString(), Y = y.ToString() };

    private static A.CustomGeometry WrapInCustGeom(A.Path path)
    {
        var g = new A.CustomGeometry();
        g.Append(new A.AdjustValueList());
        g.Append(new A.PathList(path));
        return g;
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

    // ── Shape style (fill / outline / position) ─────────────────────────

    /// <summary>Read the fill, outline, and transform of a shape into a ShapeStyle object.</summary>
    public ShapeStyle? GetShapeStyle(int slideIndex, int shapeTreeIndex)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return null;
        var elements = slidePart.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        if (elements is null || shapeTreeIndex < 0 || shapeTreeIndex >= elements.Count) return null;

        var elem      = elements[shapeTreeIndex];
        bool isPicture = elem is Picture;

        string name = elem switch
        {
            Shape s      => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "",
            Picture p    => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "",
            GroupShape g => g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "",
            _            => "",
        };

        ShapeProperties? spPr = elem switch
        {
            Shape s   => s.ShapeProperties,
            Picture p => p.ShapeProperties,
            _         => null,
        };

        var style = new ShapeStyle { Name = name, IsPicture = isPicture };

        if (spPr is not null)
        {
            var xfrm = spPr.GetFirstChild<A.Transform2D>();
            style.X  = xfrm?.Offset?.X?.Value  ?? 0;
            style.Y  = xfrm?.Offset?.Y?.Value  ?? 0;
            style.Cx = xfrm?.Extents?.Cx?.Value ?? 0;
            style.Cy = xfrm?.Extents?.Cy?.Value ?? 0;

            // Fill
            if (spPr.GetFirstChild<A.NoFill>() is not null)
            {
                style.FillKind = FillKind.None;
            }
            else if (spPr.GetFirstChild<A.SolidFill>() is A.SolidFill sf)
            {
                style.FillKind = FillKind.Solid;
                var hex = sf.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
                if (hex is not null) style.FillColor = RgbColor.FromHex(hex);
            }
            else
            {
                // No explicit fill element — inherited from layout/master; treat as None
                style.FillKind = FillKind.None;
            }

            // Outline
            var ol = spPr.GetFirstChild<A.Outline>();
            if (ol is null || ol.GetFirstChild<A.NoFill>() is not null)
            {
                style.OutlineKind = OutlineKind.None;
            }
            else
            {
                style.OutlineWidthPt = (ol.Width?.Value ?? 19050) / 12700.0;
                var hex = ol.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
                if (hex is not null) style.OutlineColor = RgbColor.FromHex(hex);

                var dash = ol.GetFirstChild<A.PresetDash>()?.Val?.Value;
                if      (dash == A.PresetLineDashValues.Dash)       style.OutlineKind = OutlineKind.Dash;
                else if (dash == A.PresetLineDashValues.Dot)        style.OutlineKind = OutlineKind.Dot;
                else if (dash == A.PresetLineDashValues.DashDot)    style.OutlineKind = OutlineKind.DashDot;
                else if (dash == A.PresetLineDashValues.LargeDashDotDot) style.OutlineKind = OutlineKind.DashDotDot;
                else                                                 style.OutlineKind = OutlineKind.Solid;
            }
        }
        else if (elem is GroupShape gs)
        {
            var tg   = gs.GroupShapeProperties?.GetFirstChild<A.TransformGroup>();
            style.X  = tg?.Offset?.X?.Value  ?? 0;
            style.Y  = tg?.Offset?.Y?.Value  ?? 0;
            style.Cx = tg?.Extents?.Cx?.Value ?? 0;
            style.Cy = tg?.Extents?.Cy?.Value ?? 0;
        }

        return style;
    }

    /// <summary>Apply a ShapeStyle (position/size/fill/outline) back to the shape.</summary>
    public void UpdateShapeStyle(int slideIndex, int shapeTreeIndex, ShapeStyle style)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return;
        var elements = slidePart.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        if (elements is null || shapeTreeIndex < 0 || shapeTreeIndex >= elements.Count) return;

        var elem      = elements[shapeTreeIndex];
        bool isPicture = elem is Picture;
        ShapeProperties? spPr = elem switch
        {
            Shape s   => s.ShapeProperties,
            Picture p => p.ShapeProperties,
            _         => null,
        };
        if (spPr is null) return;

        PushUndo();

        // Position & size
        var xfrm = spPr.GetFirstChild<A.Transform2D>();
        if (xfrm is null) { xfrm = new A.Transform2D(); spPr.InsertAt(xfrm, 0); }
        if (xfrm.Offset  is null) xfrm.Offset  = new A.Offset();
        if (xfrm.Extents is null) xfrm.Extents = new A.Extents();
        xfrm.Offset.X   = style.X;
        xfrm.Offset.Y   = style.Y;
        xfrm.Extents.Cx = style.Cx;
        xfrm.Extents.Cy = style.Cy;

        if (!isPicture)
        {
            spPr.RemoveAllChildren<A.NoFill>();
            spPr.RemoveAllChildren<A.SolidFill>();
            spPr.RemoveAllChildren<A.GradientFill>();
            spPr.RemoveAllChildren<A.PatternFill>();
            spPr.RemoveAllChildren<A.GroupFill>();

            if (style.FillKind == FillKind.None)
                spPr.Append(new A.NoFill());
            else
                spPr.Append(new A.SolidFill(
                    new A.RgbColorModelHex { Val = style.FillColor.ToHex() }));
        }

        spPr.RemoveAllChildren<A.Outline>();
        if (style.OutlineKind != OutlineKind.None)
        {
            int widthEmu = Math.Max(3175, (int)(style.OutlineWidthPt * 12700.0));
            var ol = new A.Outline { Width = widthEmu };
            ol.Append(new A.SolidFill(new A.RgbColorModelHex { Val = style.OutlineColor.ToHex() }));
            if (style.OutlineKind != OutlineKind.Solid)
            {
                var dv = style.OutlineKind == OutlineKind.Dash       ? A.PresetLineDashValues.Dash :
                         style.OutlineKind == OutlineKind.Dot        ? A.PresetLineDashValues.Dot  :
                         style.OutlineKind == OutlineKind.DashDot    ? A.PresetLineDashValues.DashDot :
                                                                        A.PresetLineDashValues.LargeDashDotDot;
                ol.Append(new A.PresetDash { Val = dv });
            }
            spPr.Append(ol);
        }

        slidePart.Slide.Save();
        _modified = true;
    }

    // ── Group / Ungroup ───────────────────────────────────────────────────

    /// <summary>Returns true if the element at shapeTreeIndex is a GroupShape.</summary>
    public bool IsGroupShape(int slideIndex, int shapeTreeIndex)
    {
        var elements = GetSlidePart(slideIndex)?.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        return elements is not null && shapeTreeIndex >= 0 && shapeTreeIndex < elements.Count
               && elements[shapeTreeIndex] is GroupShape;
    }

    /// <summary>Wrap the given shape-tree indices into a single GroupShape. Returns new tree index.</summary>
    public int GroupShapes(int slideIndex, int[] shapeTreeIndices)
    {
        if (shapeTreeIndices.Length < 2) return shapeTreeIndices.Length > 0 ? shapeTreeIndices[0] : -1;

        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return -1;
        var tree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (tree is null) return -1;

        var elements = tree.Elements<OpenXmlCompositeElement>().ToList();
        var sorted   = shapeTreeIndices.Where(i => i >= 0 && i < elements.Count)
                                       .OrderBy(i => i).ToArray();
        if (sorted.Length < 2) return sorted.Length > 0 ? sorted[0] : -1;

        // Bounding box
        long minX = long.MaxValue, minY = long.MaxValue, maxX = 0, maxY = 0;
        foreach (int idx in sorted)
        {
            A.Transform2D? xfrm = elements[idx] switch
            {
                Shape s   => s.ShapeProperties?.GetFirstChild<A.Transform2D>(),
                Picture p => p.ShapeProperties?.GetFirstChild<A.Transform2D>(),
                _         => null,
            };
            if (xfrm?.Offset is null || xfrm.Extents is null) continue;
            long x  = xfrm.Offset.X?.Value  ?? 0;
            long y  = xfrm.Offset.Y?.Value  ?? 0;
            long cx = xfrm.Extents.Cx?.Value ?? 0;
            long cy = xfrm.Extents.Cy?.Value ?? 0;
            if (x      < minX) minX = x;
            if (y      < minY) minY = y;
            if (x + cx > maxX) maxX = x + cx;
            if (y + cy > maxY) maxY = y + cy;
        }
        if (minX == long.MaxValue) return -1;

        PushUndo();
        uint maxId = GetMaxShapeId(tree);
        long w = maxX - minX, h = maxY - minY;

        var grp = new GroupShape();
        grp.Append(new NonVisualGroupShapeProperties(
            new NonVisualDrawingProperties { Id = maxId + 1, Name = $"Group {maxId + 1}" },
            new NonVisualGroupShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()));
        grp.Append(new GroupShapeProperties(
            new A.TransformGroup(
                new A.Offset       { X = minX, Y = minY },
                new A.Extents      { Cx = w,   Cy = h   },
                new A.ChildOffset  { X = minX, Y = minY },
                new A.ChildExtents { Cx = w,   Cy = h   })));

        foreach (int idx in sorted)
            grp.Append((OpenXmlCompositeElement)elements[idx].CloneNode(true));

        elements[sorted[0]].InsertBeforeSelf(grp);
        foreach (int idx in sorted.Reverse())
            elements[idx].Remove();

        slidePart.Slide.Save();
        _modified = true;

        return tree.Elements<OpenXmlCompositeElement>().ToList().IndexOf(grp);
    }

    /// <summary>Dissolve a GroupShape, restoring its children to the slide's top-level tree.
    /// Returns the new tree indices of the extracted children.</summary>
    public int[] UngroupShape(int slideIndex, int shapeTreeIndex)
    {
        var slidePart = GetSlidePart(slideIndex);
        if (slidePart is null) return [];
        var tree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (tree is null) return [];

        var elements = tree.Elements<OpenXmlCompositeElement>().ToList();
        if (shapeTreeIndex < 0 || shapeTreeIndex >= elements.Count) return [];
        if (elements[shapeTreeIndex] is not GroupShape grp) return [];

        PushUndo();

        var tg      = grp.GroupShapeProperties?.GetFirstChild<A.TransformGroup>();
        long gOffX  = tg?.Offset?.X?.Value       ?? 0;
        long gOffY  = tg?.Offset?.Y?.Value       ?? 0;
        long gCx    = tg?.Extents?.Cx?.Value     ?? 1;
        long gCy    = tg?.Extents?.Cy?.Value     ?? 1;
        long gChOffX = tg?.ChildOffset?.X?.Value  ?? gOffX;
        long gChOffY = tg?.ChildOffset?.Y?.Value  ?? gOffY;
        long gChCx   = tg?.ChildExtents?.Cx?.Value ?? gCx;
        long gChCy   = tg?.ChildExtents?.Cy?.Value ?? gCy;

        double scaleX = gChCx > 0 ? (double)gCx / gChCx : 1.0;
        double scaleY = gChCy > 0 ? (double)gCy / gChCy : 1.0;

        var children = grp.Elements<OpenXmlCompositeElement>()
            .Where(e => e is Shape or Picture or GroupShape)
            .Select(e => (OpenXmlCompositeElement)e.CloneNode(true))
            .ToList();

        foreach (var child in children)
        {
            A.Transform2D? xfrm = child switch
            {
                Shape s   => s.ShapeProperties?.GetFirstChild<A.Transform2D>(),
                Picture p => p.ShapeProperties?.GetFirstChild<A.Transform2D>(),
                _         => null,
            };
            if (xfrm?.Offset is not null && xfrm.Extents is not null)
            {
                long lx = xfrm.Offset.X?.Value  ?? 0;
                long ly = xfrm.Offset.Y?.Value  ?? 0;
                long cx = xfrm.Extents.Cx?.Value ?? 0;
                long cy = xfrm.Extents.Cy?.Value ?? 0;
                xfrm.Offset.X   = gOffX + (long)((lx - gChOffX) * scaleX);
                xfrm.Offset.Y   = gOffY + (long)((ly - gChOffY) * scaleY);
                xfrm.Extents.Cx = (long)(cx * scaleX);
                xfrm.Extents.Cy = (long)(cy * scaleY);
            }
        }

        foreach (var child in children)
            grp.InsertBeforeSelf(child);
        grp.Remove();

        slidePart.Slide.Save();
        _modified = true;

        var newElements = tree.Elements<OpenXmlCompositeElement>().ToList();
        return children.Select(c => newElements.IndexOf(c)).Where(i => i >= 0).ToArray();
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

    // ── Slide transitions (stored in p:transition) ──────────────────────

    public SlideTransition GetSlideTransition(int slideIndex)
    {
        var part = GetSlidePart(slideIndex);
        if (part is null) return new SlideTransition();

        var tr = part.Slide.GetFirstChild<Transition>();
        if (tr is null) return new SlideTransition { Kind = TransitionKind.None };

        TransitionKind kind;
        if      (tr.GetFirstChild<PushTransition>()     is not null) kind = TransitionKind.Push;
        else if (tr.GetFirstChild<WipeTransition>()     is not null) kind = TransitionKind.Wipe;
        else if (tr.GetFirstChild<CoverTransition>()    is not null) kind = TransitionKind.Flip;
        else if (tr.GetFirstChild<NewsflashTransition>() is not null) kind = TransitionKind.Crumple;
        else if (tr.GetFirstChild<FadeTransition>()     is not null) kind = TransitionKind.Fade;
        else                                                          kind = TransitionKind.None;

        // Recover in-memory Morph override if set
        if (_morphSlides.Contains(slideIndex)) kind = TransitionKind.Morph;

        var durAttr = tr.GetAttribute("dur", "");
        double durationMs = 700;
        if (durAttr.Value is string dv && dv != "auto" && uint.TryParse(dv, out var d))
            durationMs = d;

        return new SlideTransition { Kind = kind, DurationMs = durationMs };
    }

    public void SetSlideTransition(int slideIndex, SlideTransition tr, bool allSlides = false)
    {
        int count = allSlides ? SlideCount : 1;
        int start = allSlides ? 0 : slideIndex;

        PushUndo();
        for (int i = start; i < start + count; i++)
            ApplyTransitionToSlide(i, tr);
        _modified = true;
    }

    private void ApplyTransitionToSlide(int slideIndex, SlideTransition tr)
    {
        var part = GetSlidePart(slideIndex);
        if (part is null) return;

        // Remove existing
        var existing = part.Slide.GetFirstChild<Transition>();
        if (existing is not null) part.Slide.RemoveChild(existing);

        if (tr.Kind == TransitionKind.Morph)
        {
            // Morph: store Fade in PPTX, remember in-memory
            _morphSlides.Add(slideIndex);
            var fade = new Transition();
            fade.SetAttribute(new OpenXmlAttribute("spd", "", "med"));
            fade.Append(new FadeTransition());
            InsertTransition(part.Slide, fade);
        }
        else
        {
            _morphSlides.Remove(slideIndex);
        }

        if (tr.Kind is TransitionKind.None or TransitionKind.Morph)
        {
            part.Slide.Save();
            return;
        }

        var pTr = new Transition();
        pTr.SetAttribute(new OpenXmlAttribute("spd", "", "med"));
        pTr.SetAttribute(new OpenXmlAttribute("dur", "", ((int)tr.DurationMs).ToString()));

        OpenXmlElement child = tr.Kind switch
        {
            TransitionKind.Fade    => new FadeTransition(),
            TransitionKind.Push    => MakePushTr(),
            TransitionKind.Wipe    => MakeWipeTr(),
            TransitionKind.Flip    => MakeCoverTr(),
            TransitionKind.Crumple => new NewsflashTransition(),
            _                      => new FadeTransition(),
        };
        pTr.Append(child);
        InsertTransition(part.Slide, pTr);
        part.Slide.Save();
    }

    private static void InsertTransition(Slide slide, Transition tr)
    {
        var timing = slide.GetFirstChild<Timing>();
        if (timing is not null) slide.InsertBefore(tr, timing);
        else                    slide.Append(tr);
    }

    private static PushTransition MakePushTr()
    {
        var p = new PushTransition();
        p.SetAttribute(new OpenXmlAttribute("dir", "", "l"));
        return p;
    }

    private static WipeTransition MakeWipeTr()
    {
        var w = new WipeTransition();
        w.SetAttribute(new OpenXmlAttribute("dir", "", "l"));
        return w;
    }

    private static CoverTransition MakeCoverTr()
    {
        var c = new CoverTransition();
        c.SetAttribute(new OpenXmlAttribute("dir", "", "u"));
        return c;
    }

    private readonly HashSet<int> _morphSlides = new();

    // ── Object animations (in-memory only) ──────────────────────────────

    private readonly Dictionary<int, List<ShapeAnimation>> _slideAnimations = new();

    public List<ShapeAnimation> GetSlideAnimations(int slideIndex) =>
        _slideAnimations.TryGetValue(slideIndex, out var list)
            ? new List<ShapeAnimation>(list)
            : new List<ShapeAnimation>();

    public AnimationKind GetShapeAnimationKind(int slideIndex, int treeIndex)
    {
        if (!_slideAnimations.TryGetValue(slideIndex, out var list)) return AnimationKind.None;
        return list.FirstOrDefault(a => a.TreeIndex == treeIndex)?.Kind ?? AnimationKind.None;
    }

    public ShapeAnimation GetShapeAnimation(int slideIndex, int treeIndex)
    {
        if (_slideAnimations.TryGetValue(slideIndex, out var list))
        {
            var existing = list.FirstOrDefault(a => a.TreeIndex == treeIndex);
            if (existing is not null) return existing;
        }
        return new ShapeAnimation { TreeIndex = treeIndex };
    }

    public void SetShapeAnimation(int slideIndex, int treeIndex, AnimationKind kind,
                                   double durationMs = 500, bool autoPlay = false, int repeatCount = 1)
    {
        if (!_slideAnimations.TryGetValue(slideIndex, out var list))
            _slideAnimations[slideIndex] = list = new List<ShapeAnimation>();

        list.RemoveAll(a => a.TreeIndex == treeIndex);
        if (kind != AnimationKind.None)
            list.Add(new ShapeAnimation
            {
                TreeIndex = treeIndex, Kind = kind, DurationMs = durationMs,
                AutoPlay = autoPlay, RepeatCount = repeatCount
            });
        _modified = true;
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
