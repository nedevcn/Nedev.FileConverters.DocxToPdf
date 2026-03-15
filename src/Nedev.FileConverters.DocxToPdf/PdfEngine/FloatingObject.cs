namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// 浮动对象环绕样式
/// </summary>
public enum WrappingStyle
{
    Inline,
    InFrontOfText,
    BehindText,
    TopAndBottom,
    Square,
    Tight,
    Through
}

/// <summary>
/// 浮动对象
/// </summary>
public class FloatingObject : IElement
{
    public Image Image { get; }
    public WrappingStyle Wrapping { get; set; }
    public float Left { get; set; }
    public float Top { get; set; }
    public bool PositionIsAbsolute { get; set; }

    /// <summary>
    /// Distance from text (points). Applies equally on all sides; used to pad exclusion rectangles.
    /// </summary>
    public float TextDistance { get; set; }

    /// <summary>
    /// 浮动对象宽度（来自 Image.ScaledWidth）
    /// </summary>
    public float Width => Image?.ScaledWidth ?? 0;

    /// <summary>
    /// 浮动对象高度（来自 Image.ScaledHeight）
    /// </summary>
    public float Height => Image?.ScaledHeight ?? 0;

    // IElement 实现
    public int Type => -100; // Custom type
    public bool IsContent() => true;
    public bool IsNestable() => false;

    public FloatingObject(Image image)
    {
        Image = image ?? throw new ArgumentNullException(nameof(image));
        Wrapping = WrappingStyle.Inline;
    }
}
