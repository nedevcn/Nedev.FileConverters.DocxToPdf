namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// ????
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
/// ????
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

    public FloatingObject(Image image)
    {
        Image = image ?? throw new ArgumentNullException(nameof(image));
        Wrapping = WrappingStyle.Inline;
    }
}
