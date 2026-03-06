using System.IO.Compression;
using System.Text;
using SkiaSharp;

namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// PDF内容流写入器
/// </summary>
public class PdfContentByte
{
    private readonly StringBuilder _content = new();
    private readonly PdfPage? _page;
    private readonly PdfDocument? _document;

    public PdfWriter? Writer { get; set; }

    public PdfContentByte(PdfPage? page = null, PdfDocument? document = null)
    {
        _page = page;
        _document = document;
    }

    public void SaveState() => _content.Append("q\n");
    public void RestoreState() => _content.Append("Q\n");

    public void SetLineWidth(float width) => _content.Append($"{width:F2} w\n");

    public void SetColorStroke(BaseColor color) =>
        _content.Append($"{color.R / 255f:F3} {color.G / 255f:F3} {color.B / 255f:F3} RG\n");

    public void SetColorFill(BaseColor color) =>
        _content.Append($"{color.R / 255f:F3} {color.G / 255f:F3} {color.B / 255f:F3} rg\n");

    public void MoveTo(float x, float y) => _content.Append($"{x:F2} {y:F2} m\n");
    public void LineTo(float x, float y) => _content.Append($"{x:F2} {y:F2} l\n");
    public void Stroke() => _content.Append("S\n");
    public void Fill() => _content.Append("f\n");
    public void FillAndStroke() => _content.Append("B\n");

    public void Rectangle(float x, float y, float width, float height) =>
        _content.Append($"{x:F2} {y:F2} {width:F2} {height:F2} re\n");

    public void BeginText() => _content.Append("BT\n");
    // 文本矩阵缓存，减少冗余指令
    private float[]? _currentTextMatrix;

    public void EndText()
    {
        _content.Append("ET\n");
        _currentTextMatrix = null;
    }

    public void SetFontAndSize(string fontName, float size) =>
        _content.Append($"/{fontName} {size:F2} Tf\n");

    public virtual void SetTextMatrix(float a, float b, float c, float d, float e, float f)
    {
        // 简单比较，如果完全一致则跳过
        if (_currentTextMatrix != null &&
            Math.Abs(_currentTextMatrix[0] - a) < 0.001f &&
            Math.Abs(_currentTextMatrix[1] - b) < 0.001f &&
            Math.Abs(_currentTextMatrix[2] - c) < 0.001f &&
            Math.Abs(_currentTextMatrix[3] - d) < 0.001f &&
            Math.Abs(_currentTextMatrix[4] - e) < 0.01f &&
            Math.Abs(_currentTextMatrix[5] - f) < 0.01f)
        {
            return;
        }

        _content.Append($"{a:F3} {b:F3} {c:F3} {d:F3} {e:F2} {f:F2} Tm\n");
        
        if (_currentTextMatrix == null) _currentTextMatrix = new float[6];
        _currentTextMatrix[0] = a;
        _currentTextMatrix[1] = b;
        _currentTextMatrix[2] = c;
        _currentTextMatrix[3] = d;
        _currentTextMatrix[4] = e;
        _currentTextMatrix[5] = f;
    }

    public virtual void ShowText(string text)
    {
        var encoded = EncodeTextForPdf(text);
        _content.Append($"{encoded} Tj\n");
    }

    private static string EncodeTextForPdf(string text)
    {
        // 对于Identity-H编码的Type0字体，需要将字符编码为CID
        // 使用Unicode码点作为CID（Identity映射）
        var sb = new StringBuilder();
        sb.Append('<');

        foreach (var c in text)
        {
            // 将字符的Unicode码点作为16位CID（大端序）
            var cid = (ushort)c;
            sb.AppendFormat("{0:X4}", cid);
        }

        sb.Append('>');
        return sb.ToString();
    }

    public void ShowTextAligned(int alignment, string text, float x, float y, float rotation, string fontName = "F1", float fontSize = 12)
    {
        SaveState();
        BeginText();
        SetFontAndSize(fontName, fontSize);

        if (rotation != 0)
        {
            var rad = rotation * Math.PI / 180;
            var cos = (float)Math.Cos(rad);
            var sin = (float)Math.Sin(rad);
            // 旋转矩阵 + 平移，正确组合到 Tm 中
            SetTextMatrix(cos, sin, -sin, cos, x, y);
        }
        else
        {
            SetTextMatrix(1, 0, 0, 1, x, y);
        }

        var encoded = EncodeTextForPdf(text);
        _content.Append($"{encoded} Tj\n");

        EndText();
        RestoreState();
    }

    public void DrawImage(Image image, float x, float y, PdfWriter? writer = null)
    {
        var activeWriter = writer ?? Writer;
        if (activeWriter == null)
        {
            _content.Append($"% Image: {image.ScaledWidth:F2}x{image.ScaledHeight:F2} at ({x:F2}, {y:F2})\n");
            return;
        }

        var xobjectName = activeWriter.AddImageXObject(image);

        // PDF图像XObject的默认空间是1x1，cm矩阵缩放到目标像素尺寸
        _content.Append("q\n");
        _content.Append($"{image.ScaledWidth:F2} 0 0 {image.ScaledHeight:F2} {x:F2} {y:F2} cm\n");
        _content.Append($"/{xobjectName} Do\n");
        _content.Append("Q\n");
    }

    // 兼容性方法
    public void AddImage(Image image)
    {
        DrawImage(image, image.AbsoluteX, image.AbsoluteY);
    }

    public void AddImage(Image image, float x, float y)
    {
        DrawImage(image, x, y);
    }

    // 兼容性常量
    public const int TEXT_RENDER_MODE_FILL = 0;
    public const int TEXT_RENDER_MODE_FILL_STROKE = 2;

    private static string EscapeText(string text)
    {
        return text
            .Replace("\\", "\\\\")
            .Replace("(", "\\(")
            .Replace(")", "\\)")
            .Replace("\r", "\\r")
            .Replace("\n", "\\n")
            .Replace("\t", "\\t");
    }

    public string GetContent() => _content.ToString();

    public void Clear() => _content.Clear();
}

/// <summary>
/// PDF图像XObject引用
/// </summary>
public class PdfImageReference
{
    public int ObjectNumber { get; }
    public string Name { get; }
    public Image Image { get; }

    public PdfImageReference(int objectNumber, string name, Image image)
    {
        ObjectNumber = objectNumber;
        Name = name;
        Image = image;
    }
}

/// <summary>
/// PDF字体引用
/// </summary>
public class PdfFontReference
{
    public int ObjectNumber { get; set; }
    public string PdfName { get; }  // F1, F2, F3...
    public string Family { get; }
    public int Style { get; }
    public bool IsChineseFont { get; }

    public PdfFontReference(string pdfName, string family, int style, bool isChinese)
    {
        PdfName = pdfName;
        Family = family;
        Style = style;
        IsChineseFont = isChinese;
    }
}

/// <summary>
/// PDF写入器
/// </summary>
public class PdfWriter : IDisposable
{
    private readonly Stream _outputStream;
    private readonly PdfDocument _document;
    private readonly List<PdfIndirectObject> _objects = [];
    private readonly Dictionary<int, long> _xref = [];
    private int _objectNumber = 1;
    private long _xrefOffset;
    private readonly PdfContentByte _directContent;

    // 图片XObject管理
    private readonly Dictionary<string, PdfImageReference> _imageXObjects = [];
    private int _imageCounter = 0;

    // 多字体管理
    private readonly Dictionary<string, PdfFontReference> _fontRegistry = [];
    private int _fontCounter = 0;
    private string _chineseFontPath = "";
    private byte[]? _chineseFontData;
    private TrueTypeFont? _parsedFont;

    public bool CloseStream { get; set; } = true;
    public PdfContentByte DirectContent => _directContent;
    public PdfContentByte DirectContentUnder => _directContent;

    private readonly Dictionary<int, string> _pageDirectContents = [];
    private int _currentPageDirectContentIdx = 1;
    private bool _pageEventOpened;
    private bool _pageEventClosed;
    private AnnotationCollection? _annotations;
    private PdfEncryption? _encryption;
    private string? _pdfTitle;
    private string? _pdfAuthor;
    private string? _pdfSubject;
    private string? _pdfKeywords;
    private string _pdfCreator = "Nedev.FileConverters.DocxToPdf";

    public void SetEncryption(PdfEncryption encryption)
    {
        _encryption = encryption;
    }

    public void SetMetadata(string? title, string? author, string? subject, string? keywords, string? creator)
    {
        _pdfTitle = title;
        _pdfAuthor = author;
        _pdfSubject = subject;
        _pdfKeywords = keywords;
        _pdfCreator = creator ?? "Nedev.FileConverters.DocxToPdf";
    }

    public void SetAnnotationCollection(AnnotationCollection annotations)
    {
        _annotations = annotations;
    }

    public void SetRootOutline(PdfOutline outline)
    {
        RootOutline = outline;
    }

    public PdfWriter(Stream outputStream, PdfDocument document)
    {
        _outputStream = outputStream ?? throw new ArgumentNullException(nameof(outputStream));
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _directContent = new PdfContentByte();
        _directContent.Writer = this;
        _document.PageAdded += OnPageAdded;
    }

    private void OnPageAdded(object? sender, PdfPage page)
    {
        if (_pageEvent != null && !_pageEventOpened)
        {
            _pageEvent.OnOpenDocument(this, _document);
            _pageEventOpened = true;
        }

        if (_pageEvent != null && page.PageNumber > 1)
        {
            _pageEvent.OnEndPage(this, _document);
        }

        if (_currentPageDirectContentIdx <= _document.PageNumber)
        {
            _pageDirectContents[_currentPageDirectContentIdx] = _directContent.GetContent();
        }
        _directContent.Clear();
        _currentPageDirectContentIdx = page.PageNumber;

        _pageEvent?.OnStartPage(this, _document);
    }

    public static PdfWriter GetInstance(PdfDocument document, Stream outputStream)
    {
        return new PdfWriter(outputStream, document);
    }

    /// <summary>
    /// 根据Font对象获取或创建PDF字体名称（F1, F2...）
    /// </summary>
    public string GetFontPdfName(Font font)
    {
        var key = $"{font.Family}_{font.Style}";
        if (_fontRegistry.TryGetValue(key, out var existing))
            return existing.PdfName;

        _fontCounter++;
        var pdfName = $"F{_fontCounter}";
        var isChinese = IsChinese(font.Family);
        var fontRef = new PdfFontReference(pdfName, font.Family, font.Style, isChinese);
        _fontRegistry[key] = fontRef;
        return pdfName;
    }

    private static bool IsChinese(string fontFamily)
    {
        var lower = fontFamily.ToLowerInvariant();
        return lower.Contains("sim") || lower.Contains("song") || lower.Contains("hei")
            || lower.Contains("kai") || lower.Contains("fang") || lower.Contains("yahei")
            || lower.Contains("microsoft") || lower.Contains("ming") || lower.Contains("宋")
            || lower.Contains("黑") || lower.Contains("楷") || lower.Contains("仿")
            || lower.Contains("nsimsun") || lower.Contains("dengxian");
    }

    public void WriteDocument()
    {
        // 记录最后一页的 DirectContent
        if (_currentPageDirectContentIdx > 0)
        {
            _pageDirectContents[_currentPageDirectContentIdx] = _directContent.GetContent();
        }

        WriteHeader();

        // 确保至少有默认字体注册
        if (_fontRegistry.Count == 0)
        {
            GetFontPdfName(Font.Helvetica(12));
        }

        // 当使用中文字体时，先写入共享的CIDFont基础对象
        int sharedCidFontObjNum = 0;
        int sharedToUnicodeObjNum = 0;
        if (_useChineseFont)
        {
            WriteSharedChineseFontObjects(out sharedCidFontObjNum, out sharedToUnicodeObjNum);
        }

        // 写入字体对象
        var fontObjects = new Dictionary<string, PdfIndirectObject>();
        foreach (var kvp in _fontRegistry)
        {
            var fontObj = AddObject();
            kvp.Value.ObjectNumber = fontObj.Number;
            fontObjects[kvp.Key] = fontObj;
        }

        // 写入字体定义
        foreach (var kvp in _fontRegistry)
        {
            var fontRef = kvp.Value;
            var fontObj = fontObjects[kvp.Key];
            if (_useChineseFont)
            {
                // 所有字体都指向共享的中文CIDFont
                WriteType0FontRef(fontObj, sharedCidFontObjNum, sharedToUnicodeObjNum);
            }
            else if (fontRef.IsChineseFont)
            {
                WriteChineseFont(fontObj, fontRef);
            }
            else
            {
                WriteStandardFont(fontObj, fontRef);
            }
        }

        // 写入页面树
        var pagesObj = AddObject();
        var pageRefs = new List<int>();

        foreach (var page in _document.Pages)
        {
            var pageObj = WritePage(page, pagesObj.Number);
            pageRefs.Add(pageObj.Number);
        }

        // 更新页面树
        UpdatePagesObject(pagesObj, pageRefs);

        // 写入所有延迟的图片XObject
        foreach (var kvp in _imageXObjects)
        {
            var imgRef = kvp.Value;
            var imgPdfObj = new PdfIndirectObject(imgRef.ObjectNumber);
            WriteImageXObject(imgPdfObj, imgRef.Image);
        }

        // 写入目录
        var catalogObj = AddObject();
        WriteCatalog(catalogObj, pagesObj.Number);

        // 写入交叉引用表和尾部
        WriteXref();
        WriteTrailer(catalogObj.Number);
    }

    private PdfIndirectObject WritePage(PdfPage page, int pagesRef)
    {
        var pageObj = AddObject();
        var contentObj = AddObject();

        // 扫描页面中的图片并创建XObject
        var pageImages = CollectImagesFromPage(page);
        foreach (var image in pageImages)
        {
            AddImageXObject(image);
        }

        // 获取当前页的DirectContent（通过PageAdded事件快照的值）
        var directContent = _pageDirectContents.GetValueOrDefault(page.PageNumber, "");

        string content;
        if (!string.IsNullOrEmpty(directContent))
        {
            // ColumnText已经渲染了所有内容到该页的DirectContent快照
            content = directContent;
        }
        else
        {
            // 回退：如果ColumnText未使用，从页面元素生成内容
            content = GeneratePageContent(page, this);
        }

        // 编码内容流为字节
        var contentBytes = Encoding.UTF8.GetBytes(content);

        // 写入内容流（使用实际字节数作为Length）
        WriteContentStream(contentObj, contentBytes);

        // 构建字体资源字典
        var fontResources = BuildFontResourceDict();

        // 构建图片资源字典
        var imageResources = GetImageResources();
        var resources = fontResources;
        if (!string.IsNullOrEmpty(imageResources))
        {
            resources += " " + imageResources;
        }

        // 构建页面字典（可能包含注解）
        var pageDict = $"<< /Type /Page /Parent {pagesRef} 0 R /MediaBox [0 0 {page.PageSize.Width:F2} {page.PageSize.Height:F2}] /Contents {contentObj.Number} 0 R /Resources << {resources} >> ";
        
        // 添加注解
        if (_annotations != null && _annotations.GetAnnotationsForPage(page.PageNumber).Count > 0)
        {
            var annotObj = AddObject();
            var pageAnnots = _annotations.GetAnnotationsForPage(page.PageNumber);
            var annotRefs = new List<string>();
            
            foreach (var annot in pageAnnots)
            {
                var annotStr = annot.ToPdfDict(annotObj.Number);
                var annotBytes = Encoding.Latin1.GetBytes(annotStr);
                _xref[annotObj.Number] = _outputStream.Position;
                _outputStream.Write(annotBytes, 0, annotBytes.Length);
                annotRefs.Add($"{annotObj.Number} 0 R");
            }
            
            pageDict += $"/Annots [{string.Join(" ", annotRefs)}] ";
        }
        
        pageDict += ">>";
        WriteObjectText(pageObj, pageDict);

        return pageObj;
    }

    /// <summary>
    /// 构建字体资源字典
    /// </summary>
    private string BuildFontResourceDict()
    {
        var sb = new StringBuilder();
        sb.Append("/Font << ");
        foreach (var kvp in _fontRegistry)
        {
            var fontRef = kvp.Value;
            sb.Append($"/{fontRef.PdfName} {fontRef.ObjectNumber} 0 R ");
        }
        sb.Append(">>");
        return sb.ToString();
    }

    /// <summary>
    /// 从页面元素中收集所有图片
    /// </summary>
    private List<Image> CollectImagesFromPage(PdfPage page)
    {
        var images = new List<Image>();
        foreach (var element in page.Elements)
        {
            CollectImages(element, images);
        }
        return images;
    }

    private void CollectImages(IElement element, List<Image> images)
    {
        if (element is Image img)
        {
            images.Add(img);
        }
        else if (element is Paragraph para)
        {
            foreach (var chunk in para.Chunks)
            {
                if (chunk is ImageChunk imgChunk && imgChunk.Image != null)
                {
                    images.Add(imgChunk.Image);
                }
            }
        }
    }

    private string GeneratePageContent(PdfPage page, PdfWriter writer)
    {
        var cb = new PdfContentByte(page, _document);
        var y = page.PageSize.Height - page.Document.MarginTop;

        foreach (var element in page.Elements)
        {
            y = RenderElement(cb, element, page.Document.MarginLeft, y, page.PageSize.Width - page.Document.MarginRight, writer);

            // 自动分页检查
            if (y < page.Document.MarginBottom)
            {
                // 注意：当前简化实现，仅裁剪超出部分
                // 完整分页需要将剩余元素移到下一页
                break;
            }
        }

        return cb.GetContent();
    }

    private float RenderElement(PdfContentByte cb, IElement element, float x, float y, float maxWidth, PdfWriter writer)
    {
        switch (element)
        {
            case Paragraph para:
                return RenderParagraph(cb, para, x, y, maxWidth, writer);
            case Chunk chunk:
                return RenderChunk(cb, chunk, x, y);
            case Phrase phrase:
                return RenderPhrase(cb, phrase, x, y);
            case Image img:
                return RenderImage(cb, img, x, y, maxWidth, writer);
            case PdfPTable table:
                return RenderTable(cb, table, x, y, maxWidth, writer);
            case List list:
                return RenderList(cb, list, x, y, maxWidth);
            default:
                return y;
        }
    }

    private float RenderParagraph(PdfContentByte cb, Paragraph para, float x, float y, float maxWidth, PdfWriter writer)
    {
        var lineHeight = CalculateLineHeight(para);
        var fontSize = para.Font.Size;
        var contentWidth = maxWidth - x - para.IndentationRight;

        // y 是行的顶部位置
        y -= para.SpacingBefore;
        var currentY = y;
        var isFirstLine = true;

        // 收集所有行用于对齐
        var lines = new List<(List<(Chunk chunk, float width)> chunks, float totalWidth, bool isLastLine)>();
        var currentLine = new List<(Chunk chunk, float width)>();
        var currentLineWidth = 0f;
        var lineStartIndent = para.IndentationLeft + para.FirstLineIndent;

        foreach (var chunk in para.Chunks)
        {
            if (chunk is ImageChunk imgChunk && imgChunk.Image != null)
            {
                var img = imgChunk.Image;
                var imgWidth = img.ScaledWidth;
                if (currentLineWidth + imgWidth > contentWidth - lineStartIndent && currentLine.Count > 0)
                {
                    lines.Add((currentLine, currentLineWidth, false));
                    currentLine = new List<(Chunk, float)>();
                    currentLineWidth = 0;
                    lineStartIndent = para.IndentationLeft;
                }
                currentLine.Add((chunk, imgWidth));
                currentLineWidth += imgWidth;
                continue;
            }

            var content = chunk.Content;
            if (string.IsNullOrEmpty(content)) continue;

            while (!string.IsNullOrEmpty(content))
            {
                var availableWidth = contentWidth - lineStartIndent - currentLineWidth;
                var (lineText, remainingText) = FindBreakPoint(content, chunk.Font, availableWidth);

                if (string.IsNullOrEmpty(lineText) && currentLine.Count > 0)
                {
                    // 当前行换行
                    lines.Add((currentLine, currentLineWidth, false));
                    currentLine = new List<(Chunk, float)>();
                    currentLineWidth = 0;
                    lineStartIndent = para.IndentationLeft;
                    continue;
                }
                else if (string.IsNullOrEmpty(lineText))
                {
                    // 当前行连一个字符都放不下，强行放
                    if (!string.IsNullOrEmpty(content))
                    {
                        lineText = content.Substring(0, 1);
                        remainingText = content.Substring(1);
                    }
                }

                if (!string.IsNullOrEmpty(lineText))
                {
                    var lineChunk = new Chunk(lineText, chunk.Font);
                    var chunkWidth = lineChunk.GetWidth();
                    currentLine.Add((lineChunk, chunkWidth));
                    currentLineWidth += chunkWidth;
                }

                content = remainingText;

                if (!string.IsNullOrEmpty(content))
                {
                    lines.Add((currentLine, currentLineWidth, false));
                    currentLine = new List<(Chunk, float)>();
                    currentLineWidth = 0;
                    lineStartIndent = para.IndentationLeft;
                }
            }
        }

        // 最后一行
        if (currentLine.Count > 0)
        {
            lines.Add((currentLine, currentLineWidth, true));
        }

        // 渲染每一行（含对齐）
        isFirstLine = true;
        foreach (var (lineChunks, totalWidth, isLastLine) in lines)
        {
            var indent = isFirstLine ? (para.IndentationLeft + para.FirstLineIndent) : para.IndentationLeft;
            var lineContentWidth = contentWidth - indent;
            float startX;

            switch (para.Alignment)
            {
                case Element.ALIGN_CENTER:
                    startX = x + indent + (lineContentWidth - totalWidth) / 2;
                    break;
                case Element.ALIGN_RIGHT:
                    startX = x + indent + lineContentWidth - totalWidth;
                    break;
                default: // LEFT and JUSTIFIED (simplified)
                    startX = x + indent;
                    break;
            }

            var currentX = startX;
            foreach (var (chunk, width) in lineChunks)
            {
                if (chunk is ImageChunk imgChunk && imgChunk.Image != null)
                {
                    var img = imgChunk.Image;
                    var imgY = currentY - img.ScaledHeight;
                    cb.DrawImage(img, currentX, imgY, writer);
                    currentX += img.ScaledWidth;
                }
                else
                {
                    RenderChunk(cb, chunk, currentX, currentY);
                    currentX += width;
                }
            }

            currentY -= lineHeight;
            isFirstLine = false;
        }

        // 如果没有行，至少减一行的空间
        if (lines.Count == 0)
            currentY -= lineHeight;

        return currentY - para.SpacingAfter;
    }

    /// <summary>
    /// 计算行高，考虑中文字体
    /// </summary>
    private float CalculateLineHeight(Paragraph para)
    {
        var baseLineHeight = para.Font.Size * para.MultipliedLeading;

        // 如果段落包含中文字符，增加行高
        var hasChinese = para.Chunks.Any(c => c.Content?.Any(ch => ch >= '\u4e00' && ch <= '\u9fff') == true);
        if (hasChinese)
        {
            baseLineHeight *= 1.4f;
        }

        // 确保最小行高
        return Math.Max(baseLineHeight, para.Font.Size * 1.2f);
    }

    /// <summary>
    /// 计算 Phrase 的行高
    /// </summary>
    private float CalculateLineHeight(Phrase phrase)
    {
        var baseLineHeight = phrase.Font.Size * 1.2f;

        var hasChinese = phrase.Chunks.Any(c => c.Content?.Any(ch => ch >= '\u4e00' && ch <= '\u9fff') == true);
        if (hasChinese)
        {
            baseLineHeight *= 1.4f;
        }

        return Math.Max(baseLineHeight, phrase.Font.Size * 1.2f);
    }

    /// <summary>
    /// 查找文本的断点，返回（当前行文本，剩余文本）
    /// </summary>
    private (string line, string remaining) FindBreakPoint(string text, Font font, float maxWidth)
    {
        if (string.IsNullOrEmpty(text)) return (text, "");

        // 如果整行都能放下，直接返回
        var totalWidth = font.GetWidthPoint(text);
        if (totalWidth <= maxWidth)
        {
            return (text, "");
        }

        // 逐字符查找断点
        var sb = new StringBuilder();
        var currentWidth = 0f;

        for (int i = 0; i < text.Length; i++)
        {
            var c = text[i];
            var charWidth = GetCharWidth(c, font);

            if (currentWidth + charWidth > maxWidth && sb.Length > 0)
            {
                return (sb.ToString(), text.Substring(i));
            }

            sb.Append(c);
            currentWidth += charWidth;
        }

        return (sb.ToString(), "");
    }

    /// <summary>
    /// 获取单个字符的宽度
    /// </summary>
    private float GetCharWidth(char c, Font font)
    {
        if (c >= '\u4e00' && c <= '\u9fff')
        {
            return font.Size;
        }
        else if (c >= '\u3000' && c <= '\u303f')
        {
            return font.Size;
        }
        else if (c >= '\uff00' && c <= '\uffef')
        {
            return font.Size;
        }
        else
        {
            return font.Size * 0.5f;
        }
    }

    private float RenderChunk(PdfContentByte cb, Chunk chunk, float x, float y)
    {
        if (string.IsNullOrEmpty(chunk.Content)) return y;

        cb.SaveState();

        if (chunk.BackgroundColor != null)
        {
            cb.SetColorFill(chunk.BackgroundColor);
            cb.Rectangle(x, y - chunk.Font.Size * 0.2f, chunk.GetWidth(), chunk.Font.Size * 1.2f);
            cb.Fill();
        }

        // 设置文本颜色和字体
        cb.SetColorFill(chunk.Font.Color);

        // 计算文本基线位置
        // 尝试使用真实字体度量
        float textBaselineY;
        var fontRef = GetFontReference(chunk.Font);
        if (fontRef != null && fontRef.IsChineseFont && _parsedFont != null)
        {
            // TrueType 度量：Ascent 通常为正，Descent 通常为负，UnitsPerEm 为基准
            // 基线 = y (Top) - AscentScaled
            // 但这里 y 传入的是什么？通常是 Top。
            // 假设 y 是 Top，则基线在 y - AscentScaled。
            // 需注意：iText/PDF 坐标系中，字号本身就是缩放因子。
            // 字体度量是以 UnitsPerEm 为单位的整数。
            // 缩放比例 = fontSize / UnitsPerEm
            float scale = chunk.Font.Size / _parsedFont.UnitsPerEm;
            float ascent = _parsedFont.Ascent * scale;
            textBaselineY = y - ascent + chunk.TextRise;
        }
        else
        {
            // 默认估算：基线在 Top 下方约 0.8 * Size 处
            textBaselineY = y - chunk.Font.Size * 0.8f + chunk.TextRise;
        }

        // 获取正确的PDF字体名称
        var fontPdfName = GetFontPdfName(chunk.Font);

        cb.BeginText();
        cb.SetFontAndSize(fontPdfName, chunk.Font.Size);
        cb.SetTextMatrix(1, 0, 0, 1, x, textBaselineY);
        cb.ShowText(chunk.Content);
        cb.EndText();

        if (chunk.HasUnderline)
        {
            cb.SetLineWidth(chunk.UnderlineThickness);
            cb.SetColorStroke(chunk.Font.Color);
            cb.MoveTo(x, textBaselineY + chunk.UnderlineYPosition);
            cb.LineTo(x + chunk.GetWidth(), textBaselineY + chunk.UnderlineYPosition);
            cb.Stroke();
        }

        cb.RestoreState();

        return y;
    }

    private PdfFontReference? GetFontReference(Font font)
    {
        var key = $"{font.Family}_{font.Style}";
        return _fontRegistry.TryGetValue(key, out var reference) ? reference : null;
    }

    private float RenderPhrase(PdfContentByte cb, Phrase phrase, float x, float y)
    {
        var currentX = x;
        foreach (var chunk in phrase.Chunks)
        {
            RenderChunk(cb, chunk, currentX, y);
            currentX += chunk.GetWidth();
        }
        return y - CalculateLineHeight(phrase);
    }

    private float RenderImage(PdfContentByte cb, Image img, float x, float y, float maxWidth, PdfWriter writer)
    {
        var imgX = img.HasAbsolutePosition ? img.AbsoluteX : x;
        var imgY = img.HasAbsolutePosition ? img.AbsoluteY : y - img.ScaledHeight;

        cb.DrawImage(img, imgX, imgY, writer);

        return imgY - 5;
    }

    private float RenderTable(PdfContentByte cb, PdfPTable table, float x, float y, float maxWidth, PdfWriter writer)
    {
        var widths = table.GetWidths(maxWidth - x);
        var startX = x;
        var currentY = y - table.SpacingBefore;

        var cells = table.Cells.ToList();
        var numCols = table.NumberOfColumns;

        // 构建行列网格，支持colspan / rowspan
        var rows = new List<List<(PdfPCell cell, int colStart, int colSpan, int rowSpan)>>();
        var currentRow = new List<(PdfPCell cell, int colStart, int colSpan, int rowSpan)>();
        var colIndex = 0;

        // 跟踪被rowspan占据的单元格
        var occupiedCells = new Dictionary<(int row, int col), bool>();
        var rowIndex = 0;

        foreach (var cell in cells)
        {
            // 跳过被rowspan占据的列
            while (occupiedCells.ContainsKey((rowIndex, colIndex)))
            {
                colIndex++;
                if (colIndex >= numCols)
                {
                    rows.Add(currentRow);
                    currentRow = new List<(PdfPCell, int, int, int)>();
                    rowIndex++;
                    colIndex = 0;
                }
            }

            var colspan = Math.Min(cell.Colspan, numCols - colIndex);
            var rowspan = cell.Rowspan;

            currentRow.Add((cell, colIndex, colspan, rowspan));

            // 标记被rowspan占据的后续行
            for (int r = 1; r < rowspan; r++)
            {
                for (int c = 0; c < colspan; c++)
                {
                    occupiedCells[(rowIndex + r, colIndex + c)] = true;
                }
            }

            colIndex += colspan;

            if (colIndex >= numCols)
            {
                rows.Add(currentRow);
                currentRow = new List<(PdfPCell, int, int, int)>();
                rowIndex++;
                colIndex = 0;
            }
        }

        if (currentRow.Count > 0)
        {
            rows.Add(currentRow);
        }

        // 计算每行高度
        var rowHeights = new float[rows.Count];
        for (int r = 0; r < rows.Count; r++)
        {
            var maxH = 20f; // 最小高度
            foreach (var (cell, colStart, colSpan, rowSpan) in rows[r])
            {
                if (rowSpan <= 1)
                {
                    var cellWidth = 0f;
                    for (int c = colStart; c < colStart + colSpan && c < widths.Length; c++)
                        cellWidth += widths[c] * table.WidthPercentage / 100f;

                    var contentHeight = CalculateCellHeight(cell, cellWidth);
                    var cellHeight = contentHeight + cell.PaddingTop + cell.PaddingBottom;
                    maxH = Math.Max(maxH, cellHeight);
                }
            }
            rowHeights[r] = maxH;
        }

        // 渲染每行
        for (int r = 0; r < rows.Count; r++)
        {
            var rowHeight = rowHeights[r];
            var rowCellX = startX;

            foreach (var (cell, colStart, colSpan, rowSpan) in rows[r])
            {
                var cellWidth = 0f;
                for (int c = colStart; c < colStart + colSpan && c < widths.Length; c++)
                    cellWidth += widths[c] * table.WidthPercentage / 100f;

                var totalRowSpanHeight = rowHeight;
                for (int rs = 1; rs < rowSpan && r + rs < rows.Count; rs++)
                    totalRowSpanHeight += rowHeights[r + rs];

                var cellX = startX;
                for (int c = 0; c < colStart && c < widths.Length; c++)
                    cellX += widths[c] * table.WidthPercentage / 100f;

                DrawCell(cb, cell, cellX, currentY, cellWidth, totalRowSpanHeight, writer);
            }

            currentY -= rowHeight;
        }

        return currentY - table.SpacingAfter;
    }

    /// <summary>
    /// 计算单元格内容高度
    /// </summary>
    private float CalculateCellHeight(PdfPCell cell, float cellWidth)
    {
        var height = 0f;
        foreach (var elem in cell.Elements)
        {
            if (elem is Paragraph para)
            {
                var lineHeight = CalculateLineHeight(para);
                var lines = CalculateParagraphLines(para, cellWidth - cell.PaddingLeft - cell.PaddingRight);
                height += lines * lineHeight + para.SpacingBefore + para.SpacingAfter;
            }
            else if (elem is Chunk chunk)
            {
                height += chunk.Font.Size * 1.5f;
            }
        }
        return height * 1.1f;
    }

    /// <summary>
    /// 计算段落需要的行数
    /// </summary>
    private int CalculateParagraphLines(Paragraph para, float maxWidth)
    {
        var lines = 1;
        var currentWidth = para.FirstLineIndent;

        foreach (var chunk in para.Chunks)
        {
            var content = chunk.Content;
            if (string.IsNullOrEmpty(content)) continue;

            foreach (var c in content)
            {
                var charWidth = GetCharWidth(c, chunk.Font);
                if (currentWidth + charWidth > maxWidth)
                {
                    lines++;
                    currentWidth = 0;
                }
                currentWidth += charWidth;
            }
        }

        return lines;
    }

    /// <summary>
    /// 绘制单元格
    /// </summary>
    private void DrawCell(PdfContentByte cb, PdfPCell cell, float x, float y, float width, float height, PdfWriter writer)
    {
        // 绘制单元格背景
        if (cell.BackgroundColor != null)
        {
            cb.SaveState();
            cb.SetColorFill(cell.BackgroundColor);
            cb.Rectangle(x, y - height, width, height);
            cb.Fill();
            cb.RestoreState();
        }

        // 绘制单元格边框（四边独立颜色和线宽）
        // 上边框
        if (cell.BorderWidthTop > 0)
        {
            cb.SaveState();
            cb.SetLineWidth(cell.BorderWidthTop);
            cb.SetColorStroke(cell.BorderColorTop ?? BaseColor.Black);
            cb.MoveTo(x, y);
            cb.LineTo(x + width, y);
            cb.Stroke();
            cb.RestoreState();
        }

        // 下边框
        if (cell.BorderWidthBottom > 0)
        {
            cb.SaveState();
            cb.SetLineWidth(cell.BorderWidthBottom);
            cb.SetColorStroke(cell.BorderColorBottom ?? BaseColor.Black);
            cb.MoveTo(x, y - height);
            cb.LineTo(x + width, y - height);
            cb.Stroke();
            cb.RestoreState();
        }

        // 左边框
        if (cell.BorderWidthLeft > 0)
        {
            cb.SaveState();
            cb.SetLineWidth(cell.BorderWidthLeft);
            cb.SetColorStroke(cell.BorderColorLeft ?? BaseColor.Black);
            cb.MoveTo(x, y);
            cb.LineTo(x, y - height);
            cb.Stroke();
            cb.RestoreState();
        }

        // 右边框
        if (cell.BorderWidthRight > 0)
        {
            cb.SaveState();
            cb.SetLineWidth(cell.BorderWidthRight);
            cb.SetColorStroke(cell.BorderColorRight ?? BaseColor.Black);
            cb.MoveTo(x + width, y);
            cb.LineTo(x + width, y - height);
            cb.Stroke();
            cb.RestoreState();
        }

        // 渲染单元格内容
        var contentY = y - cell.PaddingTop;
        var contentMaxWidth = x + width - cell.PaddingRight;
        foreach (var elem in cell.Elements)
        {
            contentY = RenderElement(cb, elem, x + cell.PaddingLeft, contentY, contentMaxWidth, writer);
        }
    }

    private float RenderList(PdfContentByte cb, List list, float x, float y, float maxWidth)
    {
        var currentY = y;
        var itemNumber = 1;

        foreach (var item in list.Items)
        {
            var symbol = list.ListType == List.ORDERED ? $"{itemNumber}." : list.ListSymbol.Content;
            var fontPdfName = GetFontPdfName(item.Font);

            // 符号渲染（不再嵌套 BT/ET）
            var symbolChunk = new Chunk(symbol, item.Font);
            RenderChunk(cb, symbolChunk, x + list.IndentationLeft - list.SymbolIndent, currentY);

            currentY = RenderParagraph(cb, item, x + list.IndentationLeft, currentY, maxWidth, this);
            itemNumber++;
        }

        return currentY;
    }

    // ================================================
    // PDF 底层写入方法
    // ================================================

    private void WriteHeader()
    {
        // PDF header: %PDF-1.4 followed by binary marker (4 bytes >= 128)
        var header = Encoding.Latin1.GetBytes("%PDF-1.4\n");
        _outputStream.Write(header);
        byte[] binaryMarker = [(byte)'%', 0xE2, 0xE3, 0xCF, 0xD3, (byte)'\n'];
        _outputStream.Write(binaryMarker);
    }

    private PdfIndirectObject AddObject()
    {
        var obj = new PdfIndirectObject(_objectNumber++);
        _objects.Add(obj);
        return obj;
    }

    /// <summary>
    /// 写入文本型PDF对象（字典等）
    /// </summary>
    private void WriteObjectText(PdfIndirectObject obj, string content)
    {
        _xref[obj.Number] = _outputStream.Position;
        var data = $"{obj.Number} 0 obj\n{content}\nendobj\n\n";
        WriteBytes(Encoding.Latin1.GetBytes(data));
    }

    /// <summary>
    /// 写入内容流对象（使用实际字节长度）
    /// </summary>
    private void WriteContentStream(PdfIndirectObject obj, byte[] content)
    {
        _xref[obj.Number] = _outputStream.Position;
        var header = $"{obj.Number} 0 obj\n<< /Length {content.Length} >>\nstream\n";
        WriteBytes(Encoding.Latin1.GetBytes(header));
        _outputStream.Write(content, 0, content.Length);
        WriteBytes(Encoding.Latin1.GetBytes("\nendstream\nendobj\n\n"));
    }

    private void UpdatePagesObject(PdfIndirectObject pagesObj, List<int> pageRefs)
    {
        var kids = string.Join(" ", pageRefs.Select(r => $"{r} 0 R"));
        var content = $"<< /Type /Pages /Kids [{kids}] /Count {pageRefs.Count} >>";
        WriteObjectText(pagesObj, content);
    }

    private void WriteCatalog(PdfIndirectObject catalogObj, int pagesRef)
    {
        var content = $"<< /Type /Catalog /Pages {pagesRef} 0 R ";
        
        if (RootOutline != null)
        {
            var outlineObj = AddObject();
            var outlineContent = WriteOutlineRecursive(RootOutline, pagesRef, 1);
            WriteObjectText(outlineObj, outlineContent);
            content += $"/Outlines {outlineObj.Number} 0 R ";
        }
        
        content += ">>";
        WriteObjectText(catalogObj, content);
    }
    
    private string WriteOutlineRecursive(PdfOutline outline, int pagesRef, int indent)
    {
        var sb = new StringBuilder();
        var titleEscaped = EscapePdfString(outline.Title);
        
        sb.Append($"<< /Title ({titleEscaped}) ");
        
        if (outline.Destination != null)
        {
            sb.Append($"/Dest [{pagesRef} 0 R /Fit] ");
        }
        
        if (outline.Children.Count > 0)
        {
            var kids = new List<string>();
            foreach (var child in outline.Children)
            {
                var childObj = AddObject();
                var childContent = WriteOutlineRecursive(child, pagesRef, indent + 1);
                WriteObjectText(childObj, childContent);
                kids.Add($"{childObj.Number} 0 R");
            }
            sb.Append($"/First {kids[0]} 0 R /Last {kids[^1]} 0 R /Count {outline.Children.Count} ");
        }
        
        if (outline.Parent != null)
        {
            var parentObjNum = FindOutlineObjectNumber(outline.Parent);
            if (parentObjNum > 0)
                sb.Append($"/Parent {parentObjNum} 0 R ");
        }
        
        sb.Append(">>");
        return sb.ToString();
    }
    
    private int FindOutlineObjectNumber(PdfOutline outline)
    {
        return _objects.FirstOrDefault(o => o.Number > 1000)?.Number ?? 0;
    }
    
    private static string EscapePdfString(string s)
    {
        return s.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");
    }

    private void WriteXref()
    {
        _xrefOffset = _outputStream.Position;

        // 按对象编号顺序写入，确保所有对象都有xref条目
        var maxObjNum = _objects.Count > 0 ? _objects.Max(o => o.Number) : 0;

        // xref header
        var headerLine = $"xref\n0 {maxObjNum + 1}\n";
        WriteBytes(Encoding.Latin1.GetBytes(headerLine));

        // 第0个对象（空闲对象）- 精确20字节: "0000000000 65535 f \r\n"
        WriteBytes(Encoding.Latin1.GetBytes("0000000000 65535 f \r\n"));

        for (int i = 1; i <= maxObjNum; i++)
        {
            if (_xref.TryGetValue(i, out var offset))
            {
                // 精确20字节: "nnnnnnnnnn 00000 n \r\n"
                var entry = $"{offset:D10} 00000 n \r\n";
                WriteBytes(Encoding.Latin1.GetBytes(entry));
            }
            else
            {
                WriteBytes(Encoding.Latin1.GetBytes("0000000000 65535 f \r\n"));
            }
        }
    }

    // ================================================
    // 字体相关
    // ================================================

    private string _chineseFontName = "SimSun";
    private bool _useChineseFont = false;

    /// <summary>
    /// 设置中文字体
    /// </summary>
    public void SetChineseFont(string fontName)
    {
        _chineseFontName = fontName;
        _useChineseFont = true;

        // 尝试找到并解析字体文件
        TryLoadChineseFontFile(fontName);

        // 如果找不到指定的中文字体，尝试使用系统默认
        if (_parsedFont == null)
        {
            var fallback = Nedev.FileConverters.DocxToPdf.Helpers.SystemFontProvider.GetFontPath("Microsoft YaHei") 
                           ?? Nedev.FileConverters.DocxToPdf.Helpers.SystemFontProvider.GetFontPath("SimSun");
                           
            if (fallback != null)
            {
                TryLoadFontFromPath(fallback);
                _chineseFontName = Path.GetFileNameWithoutExtension(fallback);
            }
        }

        Console.WriteLine($"Using Chinese font: {_chineseFontName} (Loaded: {_parsedFont != null})");
    }

    private void TryLoadChineseFontFile(string fontName)
    {
        // 1. 使用 SystemFontProvider 查找
        var path = Nedev.FileConverters.DocxToPdf.Helpers.SystemFontProvider.GetFontPath(fontName);
        if (path != null)
        {
            TryLoadFontFromPath(path);
            return;
        }

        // 2. 传统扫描（保留兼容性，虽然 SystemFontProvider 已覆盖大部分）
        var fontDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts));
        if (!Directory.Exists(fontDir))
        {
             if (OperatingSystem.IsWindows())
                 fontDir = @"C:\Windows\Fonts";
             else
                 return; // 非 Windows 且未找到字体，放弃
        }

        // 搜索匹配的字体文件
        var searchNames = new[] { fontName, fontName.Replace(" ", "") };
        foreach (var name in searchNames)
        {
            foreach (var ext in new[] { ".ttf", ".ttc", ".otf" })
            {
                var p = Path.Combine(fontDir, name + ext);
                if (File.Exists(p))
                {
                    TryLoadFontFromPath(p);
                    return;
                }
            }
        }
    }

    private void TryLoadFontFromPath(string path)
    {
        try
        {
            _chineseFontPath = path;
            _chineseFontData = File.ReadAllBytes(path);
            _parsedFont = new TrueTypeFont(_chineseFontData);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load font from {path}: {ex.Message}");
        }
    }

    private void WriteStandardFont(PdfIndirectObject fontObj, PdfFontReference fontRef)
    {
        var baseFontName = fontRef.Family;
        if ((fontRef.Style & Font.BOLD) != 0 && (fontRef.Style & Font.ITALIC) != 0)
            baseFontName += "-BoldOblique";
        else if ((fontRef.Style & Font.BOLD) != 0)
            baseFontName += "-Bold";
        else if ((fontRef.Style & Font.ITALIC) != 0)
            baseFontName += "-Oblique";

        var fontDict = $"<< /Type /Font /Subtype /Type1 /BaseFont /{baseFontName} /Encoding /WinAnsiEncoding >>";
        WriteObjectText(fontObj, fontDict);
    }

    /// <summary>
    /// 写入共享的中文字体基础对象（FontDescriptor, CIDFont, ToUnicode, CIDToGIDMap）
    /// 所有字体引用共享这些对象
    /// </summary>
    private void WriteSharedChineseFontObjects(out int cidFontObjNum, out int toUnicodeObjNum)
    {
        var fontDescriptorObj = AddObject();
        var cidFontObj = AddObject();
        var toUnicodeObj = AddObject();
        var cidToGidMapObj = AddObject();

        var fontName = _chineseFontName.Replace(" ", "");

        // 使用解析的字体信息或默认值
        int ascent = _parsedFont?.Ascent ?? 800;
        int descent = _parsedFont?.Descent ?? -200;
        int capHeight = _parsedFont?.CapHeight ?? 700;
        int stemV = _parsedFont?.StemV ?? 80;
        int xMin = _parsedFont?.XMin ?? -100;
        int yMin = _parsedFont?.YMin ?? -200;
        int xMax = _parsedFont?.XMax ?? 1000;
        int yMax = _parsedFont?.YMax ?? 800;
        int flags = 32; // Symbolic

        // 字体描述符
        string descriptorContent;
        if (_chineseFontData != null)
        {
            // 嵌入字体文件
            var fontFileObj = AddObject();
            WriteFontFileStream(fontFileObj, _chineseFontData);
            descriptorContent = $"<< /Type /FontDescriptor /FontName /{fontName} /Flags {flags} /ItalicAngle 0 /Ascent {ascent} /Descent {descent} /CapHeight {capHeight} /StemV {stemV} /FontBBox [{xMin} {yMin} {xMax} {yMax}] /FontFile2 {fontFileObj.Number} 0 R >>";
        }
        else
        {
            descriptorContent = $"<< /Type /FontDescriptor /FontName /{fontName} /Flags {flags} /ItalicAngle 0 /Ascent {ascent} /Descent {descent} /CapHeight {capHeight} /StemV {stemV} /FontBBox [{xMin} {yMin} {xMax} {yMax}] >>";
        }
        WriteObjectText(fontDescriptorObj, descriptorContent);

        // 构建CIDToGIDMap流 — 从字体cmap的Unicode→GlyphID映射
        var cidToGidData = BuildCIDToGIDMap();
        WriteContentStream(cidToGidMapObj, cidToGidData);

        // 构建W数组 (为ASCII字符 32-126 提供准确的显示宽度)
        var dw = 1000;
        var wArray = new StringBuilder();
        wArray.Append("[ 32 [ ");
        for (int i = 32; i <= 126; i++)
        {
            var glyphId = _parsedFont?.GetGlyphId((char)i) ?? 0;
            var rawWidth = _parsedFont?.GetGlyphWidth(glyphId) ?? (_parsedFont?.UnitsPerEm / 2 ?? dw);
            var unitsPerEm = _parsedFont?.UnitsPerEm ?? 2048;
            var pdfWidth = (int)Math.Round((double)rawWidth * 1000 / unitsPerEm);
            wArray.Append(pdfWidth).Append(' ');
        }
        wArray.Append("] ]");

        // CID字体 — 使用自定义CIDToGIDMap而非/Identity，并带上英文字符宽度W数组
        var cidFont = $"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{fontName} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor {fontDescriptorObj.Number} 0 R /DW {dw} /W {wArray} /CIDToGIDMap {cidToGidMapObj.Number} 0 R >>";
        WriteObjectText(cidFontObj, cidFont);

        // 构建ToUnicode CMap
        var cmap = BuildToUnicodeCMap();
        var cmapBytes = Encoding.Latin1.GetBytes(cmap);
        WriteContentStream(toUnicodeObj, cmapBytes);

        cidFontObjNum = cidFontObj.Number;
        toUnicodeObjNum = toUnicodeObj.Number;
    }

    /// <summary>
    /// 写入Type0字体引用（指向共享的CIDFont）
    /// </summary>
    private void WriteType0FontRef(PdfIndirectObject fontDictObj, int cidFontObjNum, int toUnicodeObjNum)
    {
        var fontName = _chineseFontName.Replace(" ", "");
        var fontDict = $"<< /Type /Font /Subtype /Type0 /BaseFont /{fontName} /Encoding /Identity-H /DescendantFonts [{cidFontObjNum} 0 R] /ToUnicode {toUnicodeObjNum} 0 R >>";
        WriteObjectText(fontDictObj, fontDict);
    }

    private void WriteChineseFont(PdfIndirectObject fontDictObj, PdfFontReference fontRef)
    {
        // 非_useChineseFont模式下的独立中文字体写入
        var fontDescriptorObj = AddObject();
        var cidFontObj = AddObject();
        var toUnicodeObj = AddObject();
        var cidToGidMapObj = AddObject();

        var fontName = _chineseFontName.Replace(" ", "");
        int ascent = _parsedFont?.Ascent ?? 800;
        int descent = _parsedFont?.Descent ?? -200;
        int capHeight = _parsedFont?.CapHeight ?? 700;
        int stemV = _parsedFont?.StemV ?? 80;
        int xMin = _parsedFont?.XMin ?? -100;
        int yMin = _parsedFont?.YMin ?? -200;
        int xMax = _parsedFont?.XMax ?? 1000;
        int yMax = _parsedFont?.YMax ?? 800;
        int flags = 32;

        string descriptorContent;
        if (_chineseFontData != null)
        {
            var fontFileObj = AddObject();
            WriteFontFileStream(fontFileObj, _chineseFontData);
            descriptorContent = $"<< /Type /FontDescriptor /FontName /{fontName} /Flags {flags} /ItalicAngle 0 /Ascent {ascent} /Descent {descent} /CapHeight {capHeight} /StemV {stemV} /FontBBox [{xMin} {yMin} {xMax} {yMax}] /FontFile2 {fontFileObj.Number} 0 R >>";
        }
        else
        {
            descriptorContent = $"<< /Type /FontDescriptor /FontName /{fontName} /Flags {flags} /ItalicAngle 0 /Ascent {ascent} /Descent {descent} /CapHeight {capHeight} /StemV {stemV} /FontBBox [{xMin} {yMin} {xMax} {yMax}] >>";
        }
        WriteObjectText(fontDescriptorObj, descriptorContent);

        var cidToGidData = BuildCIDToGIDMap();
        WriteContentStream(cidToGidMapObj, cidToGidData);

        var cidFont = $"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{fontName} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor {fontDescriptorObj.Number} 0 R /DW 1000 /CIDToGIDMap {cidToGidMapObj.Number} 0 R >>";
        WriteObjectText(cidFontObj, cidFont);

        var cmap = BuildToUnicodeCMap();
        var cmapBytes = Encoding.Latin1.GetBytes(cmap);
        WriteContentStream(toUnicodeObj, cmapBytes);

        var fontDict = $"<< /Type /Font /Subtype /Type0 /BaseFont /{fontName} /Encoding /Identity-H /DescendantFonts [{cidFontObj.Number} 0 R] /ToUnicode {toUnicodeObj.Number} 0 R >>";
        WriteObjectText(fontDictObj, fontDict);
    }

    /// <summary>
    /// 写入字体文件流（FontFile2）
    /// </summary>
    private void WriteFontFileStream(PdfIndirectObject obj, byte[] fontData)
    {
        // 使用FlateDecode压缩字体数据
        byte[] compressed;
        using (var ms = new MemoryStream())
        {
            using (var ds = new DeflateStream(ms, CompressionLevel.Optimal, true))
            {
                ds.Write(fontData, 0, fontData.Length);
            }
            compressed = ms.ToArray();
        }

        // 需要加zlib头（78 01）和结尾adler32
        byte[] zlibData;
        using (var ms = new MemoryStream())
        {
            ms.WriteByte(0x78); // zlib header
            ms.WriteByte(0x01); // low compression flag
            ms.Write(compressed, 0, compressed.Length);

            // 计算adler32
            uint adler = Adler32(fontData);
            ms.WriteByte((byte)(adler >> 24));
            ms.WriteByte((byte)(adler >> 16));
            ms.WriteByte((byte)(adler >> 8));
            ms.WriteByte((byte)(adler));

            zlibData = ms.ToArray();
        }

        _xref[obj.Number] = _outputStream.Position;
        var header = $"{obj.Number} 0 obj\n<< /Length {zlibData.Length} /Length1 {fontData.Length} /Filter /FlateDecode >>\nstream\n";
        WriteBytes(Encoding.Latin1.GetBytes(header));
        _outputStream.Write(zlibData, 0, zlibData.Length);
        WriteBytes(Encoding.Latin1.GetBytes("\nendstream\nendobj\n\n"));
    }

    private static uint Adler32(byte[] data)
    {
        uint a = 1, b = 0;
        foreach (var d in data)
        {
            a = (a + d) % 65521;
            b = (b + a) % 65521;
        }
        return (b << 16) | a;
    }

    /// <summary>
    /// 构建CIDToGIDMap流 — 将Unicode CID映射到字体中的Glyph ID
    /// 格式：65536个2字节大端序整数，CID[i] → GID
    /// </summary>
    private byte[] BuildCIDToGIDMap()
    {
        // CIDToGIDMap: 65536 entries × 2 bytes = 131072 bytes
        var data = new byte[65536 * 2];

        if (_parsedFont != null)
        {
            foreach (var kvp in _parsedFont.UnicodeToGlyph)
            {
                var unicode = (int)kvp.Key;
                var glyphId = kvp.Value;
                if (unicode >= 0 && unicode < 65536)
                {
                    data[unicode * 2] = (byte)(glyphId >> 8);      // 高字节
                    data[unicode * 2 + 1] = (byte)(glyphId & 0xFF); // 低字节
                }
            }
        }
        else
        {
            // 无字体数据时使用Identity映射作为回退
            for (int i = 0; i < 65536; i++)
            {
                data[i * 2] = (byte)(i >> 8);
                data[i * 2 + 1] = (byte)(i & 0xFF);
            }
        }

        return data;
    }

    /// <summary>
    /// 构建包含bfrange的ToUnicode CMap
    /// CID = Unicode code point (因为我们在EncodeTextForPdf中用Unicode编码)
    /// 所以ToUnicode就是Identity映射
    /// </summary>
    private string BuildToUnicodeCMap()
    {
        var sb = new StringBuilder();
        sb.Append("/CIDInit /ProcSet findresource begin\n");
        sb.Append("12 dict begin\n");
        sb.Append("begincmap\n");
        sb.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
        sb.Append("/CMapName /Adobe-Identity-UCS def\n");
        sb.Append("/CMapType 2 def\n");
        sb.Append("1 begincodespacerange\n");
        sb.Append("<0000> <FFFF>\n");
        sb.Append("endcodespacerange\n");

        // Identity映射：CID = Unicode
        sb.Append("5 beginbfrange\n");
        sb.Append("<0020> <007E> <0020>\n");  // ASCII可打印字符
        sb.Append("<00A0> <00FF> <00A0>\n");  // Latin扩展
        sb.Append("<2000> <206F> <2000>\n");  // 通用标点
        sb.Append("<3000> <303F> <3000>\n");  // CJK标点
        sb.Append("<4E00> <9FFF> <4E00>\n");  // CJK统一汉字
        sb.Append("endbfrange\n");

        sb.Append("endcmap\n");
        sb.Append("CMapName currentdict /CMap defineresource pop\n");
        sb.Append("end\n");
        sb.Append("end\n");
        return sb.ToString();
    }

    /// <summary>
    /// 添加图片XObject并返回引用名称
    /// </summary>
    public string AddImageXObject(Image image)
    {
        var imageHash = Convert.ToBase64String(System.Security.Cryptography.MD5.HashData(image.ImageData));
        if (_imageXObjects.TryGetValue(imageHash, out var existingRef))
        {
            return existingRef.Name;
        }

        _imageCounter++;
        var xobjectName = $"Im{_imageCounter}";
        var imageObj = AddObject();

        // 延迟写入，防止在生成中间页面/Content时将 XObject 交叉写入到了文档流中导致 PDF 损坏
        var reference = new PdfImageReference(imageObj.Number, xobjectName, image);
        _imageXObjects[imageHash] = reference;

        return xobjectName;
    }

    /// <summary>
    /// 写入图像XObject - 统一使用JPEG/DCTDecode
    /// </summary>
    private void WriteImageXObject(PdfIndirectObject imageObj, Image image)
    {
        byte[] imageData;
        string filter;
        string colorSpace;
        int bitsPerComponent = 8;

        if (image.Format == SKEncodedImageFormat.Jpeg)
        {
            // JPEG图片直接使用DCTDecode
            imageData = image.ImageData;
            filter = "/DCTDecode";
            colorSpace = "/DeviceRGB";
        }
        else
        {
            // 非JPEG格式统一转为JPEG，避免PNG数据+FlateDecode不兼容的问题
            imageData = image.GetJpegData(90);
            filter = "/DCTDecode";
            colorSpace = "/DeviceRGB";
        }

        var width = (int)image.OriginalWidth;
        var height = (int)image.OriginalHeight;

        var streamDict = $"<< /Type /XObject /Subtype /Image /Width {width} /Height {height} /ColorSpace {colorSpace} /BitsPerComponent {bitsPerComponent} /Filter {filter} /Length {imageData.Length} >>";

        _xref[imageObj.Number] = _outputStream.Position;
        var header = $"{imageObj.Number} 0 obj\n{streamDict}\nstream\n";
        WriteBytes(Encoding.Latin1.GetBytes(header));
        _outputStream.Write(imageData, 0, imageData.Length);
        WriteBytes(Encoding.Latin1.GetBytes("\nendstream\nendobj\n\n"));
    }

    /// <summary>
    /// 获取图片XObject的资源字典字符串
    /// </summary>
    public string GetImageResources()
    {
        if (_imageXObjects.Count == 0) return "";

        var sb = new StringBuilder();
        sb.Append("/XObject << ");
        foreach (var kvp in _imageXObjects)
        {
            sb.Append($"/{kvp.Value.Name} {kvp.Value.ObjectNumber} 0 R ");
        }
        sb.Append(">>");
        return sb.ToString();
    }

    private void WriteTrailer(int rootObjNumber)
    {
        var maxObjNum = _objects.Count > 0 ? _objects.Max(o => o.Number) : 0;
        var size = maxObjNum + 1;
        
        var trailer = $"trailer\n<< /Size {size} /Root {rootObjNumber} 0 R ";
        
        // 添加PDF元数据
        if (!string.IsNullOrEmpty(_pdfTitle) || !string.IsNullOrEmpty(_pdfAuthor) || 
            !string.IsNullOrEmpty(_pdfSubject) || !string.IsNullOrEmpty(_pdfKeywords))
        {
            var infoDict = new StringBuilder();
            infoDict.Append("/Info << ");
            
            if (!string.IsNullOrEmpty(_pdfTitle))
                infoDict.Append($"/Title ({EscapePdfString(_pdfTitle)}) ");
            if (!string.IsNullOrEmpty(_pdfAuthor))
                infoDict.Append($"/Author ({EscapePdfString(_pdfAuthor)}) ");
            if (!string.IsNullOrEmpty(_pdfSubject))
                infoDict.Append($"/Subject ({EscapePdfString(_pdfSubject)}) ");
            if (!string.IsNullOrEmpty(_pdfKeywords))
                infoDict.Append($"/Keywords ({EscapePdfString(_pdfKeywords)}) ");
            if (!string.IsNullOrEmpty(_pdfCreator))
                infoDict.Append($"/Creator ({EscapePdfString(_pdfCreator)}) ");
            
            infoDict.Append($"/CreationDate (D:{DateTime.Now:yyyyMMddHHmmss}) ");
            infoDict.Append(">>");
            
            var infoObj = AddObject();
            WriteObjectText(infoObj, infoDict.ToString());
            trailer += $"/Info {infoObj.Number} 0 R ";
        }
        
        trailer += ">>\nstartxref\n{_xrefOffset}\n%%EOF\n";
        trailer = trailer.Replace("{_xrefOffset}", _xrefOffset.ToString());
        
        WriteBytes(Encoding.Latin1.GetBytes(trailer));
    }

    private void WriteBytes(byte[] data)
    {
        _outputStream.Write(data, 0, data.Length);
    }

    // 兼容性属性和方法
    public PdfOutline? RootOutline { get; private set; }

    private PdfPageEventHelper? _pageEvent;
    public PdfPageEventHelper? PageEvent
    {
        get => _pageEvent;
        set => _pageEvent = value;
    }

    public float GetVerticalPosition(bool b)
    {
        return _document?.PageSize.Height ?? 0;
    }

    public void Close()
    {
        try
        {
            if (!_pageEventClosed && _pageEvent != null)
            {
                if (!_pageEventOpened)
                {
                    _pageEvent.OnOpenDocument(this, _document);
                    _pageEventOpened = true;
                }

                if (_document.PageNumber > 0)
                {
                    _pageEvent.OnEndPage(this, _document);
                }

                _pageEvent.OnCloseDocument(this, _document);
                _pageEventClosed = true;
            }

            if (_document != null && _outputStream != null)
            {
                WriteDocument();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[PdfWriter] Error closing document: {ex.Message}");
        }
        finally
        {
            if (CloseStream)
            {
                _outputStream?.Close();
            }
        }
    }

    public void Dispose()
    {
        Close();
        GC.SuppressFinalize(this);
    }

    private class PdfIndirectObject
    {
        public int Number { get; }
        public PdfIndirectObject(int number) => Number = number;
    }
}
