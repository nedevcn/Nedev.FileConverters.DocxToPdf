namespace Nedev.DocxToPdf.PdfEngine;

/// <summary>
/// 元素对齐常量
/// </summary>
public static class Element
{
    public const int ALIGN_LEFT = 0;
    public const int ALIGN_CENTER = 1;
    public const int ALIGN_RIGHT = 2;
    public const int ALIGN_JUSTIFIED = 3;
    public const int ALIGN_TOP = 4;
    public const int ALIGN_MIDDLE = 5;
    public const int ALIGN_BOTTOM = 6;
    public const int ALIGN_BASELINE = 7;

    public const int NO_BORDER = 0;
    public const int TOP = 1;
    public const int BOTTOM = 2;
    public const int LEFT = 4;
    public const int RIGHT = 8;
}

/// <summary>
/// PDF元素接口
/// </summary>
public interface IElement
{
    int Type { get; }
    bool IsContent();
    bool IsNestable();
}

/// <summary>
/// 文本块
/// </summary>
public class Chunk : IElement
{
    public const int STANDARD = 0;
    public const int NEXTPAGE = 1;
    public const int PAGEBREAK = 2;

    public string Content { get; set; }
    public Font Font { get; set; }
    public BaseColor? BackgroundColor { get; set; }
    public float TextRise { get; set; }
    public string? Anchor { get; set; }
    public bool HasUnderline { get; set; }
    public float UnderlineThickness { get; set; } = 0.1f;
    public float UnderlineYPosition { get; set; } = -1f;

    public int Type => STANDARD;
    public bool IsContent() => true;
    public bool IsNestable() => false;

    public Chunk(string content, Font? font = null)
    {
        Content = content ?? "";
        Font = font ?? Font.Helvetica(12);
    }

    public Chunk SetBackground(BaseColor color)
    {
        BackgroundColor = color;
        return this;
    }

    public Chunk SetTextRise(float rise)
    {
        TextRise = rise;
        return this;
    }

    public Chunk SetAnchor(string anchor)
    {
        Anchor = anchor;
        return this;
    }

    public Chunk SetUnderline(float thickness = 0.1f, float yPosition = -1f)
    {
        HasUnderline = true;
        UnderlineThickness = thickness;
        UnderlineYPosition = yPosition;
        return this;
    }

    // 兼容性方法
    public Chunk SetTextRenderMode(int mode, float strokeWidth, BaseColor strokeColor)
    {
        // 简化实现，仅记录状态
        return this;
    }

    public float GetWidth() => Font.GetWidthPoint(Content);
}

/// <summary>
/// 短语（多个Chunk的集合）
/// </summary>
public class Phrase : IElement
{
    private readonly List<Chunk> _chunks = [];

    public int Type => 1;
    public bool IsContent() => true;
    public bool IsNestable() => false;

    public Font Font { get; set; }
    public float Leading { get; set; }
    public IList<Chunk> Chunks => _chunks;

    public Phrase(string? text = null, Font? font = null)
    {
        Font = font ?? Font.Helvetica(12);
        if (!string.IsNullOrEmpty(text))
        {
            Add(new Chunk(text, Font));
        }
    }

    public void Add(Chunk chunk)
    {
        _chunks.Add(chunk);
    }

    public void Add(string text)
    {
        _chunks.Add(new Chunk(text, Font));
    }

    public string GetContent() => string.Join("", _chunks.Select(c => c.Content));
}

/// <summary>
/// 图片块 - 用于在段落中嵌入图片
/// </summary>
public class ImageChunk : Chunk
{
    public Image Image { get; }

    public ImageChunk(Image image) : base("", null)
    {
        Image = image;
    }

    public new float GetWidth() => Image?.ScaledWidth ?? 0;
}

/// <summary>
/// 段落
/// </summary>
public class Paragraph : Phrase
{
    public int Alignment { get; set; } = Element.ALIGN_LEFT;
    public float IndentationLeft { get; set; }
    public float IndentationRight { get; set; }
    public float FirstLineIndent { get; set; }
    public float SpacingBefore { get; set; }
    public float SpacingAfter { get; set; }
    public float MultipliedLeading { get; set; } = 1.2f;
    public bool KeepTogether { get; set; }

    public new int Type => 2;

    public Paragraph(string? text = null, Font? font = null) : base(text, font) { }

    public void SetLeading(float fixedLeading, float multipliedLeading)
    {
        Leading = fixedLeading;
        MultipliedLeading = multipliedLeading;
    }

    /// <summary>
    /// 添加图片到段落
    /// </summary>
    public void Add(Image image)
    {
        Add(new ImageChunk(image));
    }
}
