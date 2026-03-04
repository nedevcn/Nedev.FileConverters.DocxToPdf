namespace Nedev.DocxToPdf.PdfEngine;

/// <summary>
/// 环绕样式
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

    public float Width => Image.ScaledWidth;
    public float Height => Image.ScaledHeight;

    public int Type => 100;
    public bool IsContent() => true;
    public bool IsNestable() => false;

    public FloatingObject(Image image)
    {
        Image = image ?? throw new ArgumentNullException(nameof(image));
        Wrapping = WrappingStyle.Inline;
    }
}
