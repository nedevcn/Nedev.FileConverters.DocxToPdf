namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// PDF??/?????
/// </summary>
public class Rectangle
{
    public float Left { get; set; }
    public float Bottom { get; set; }
    public float Right { get; set; }
    public float Top { get; set; }

    public float Width => Right - Left;
    public float Height => Top - Bottom;

    public static readonly Rectangle A4 = new(0, 0, 595, 842);
    public static readonly Rectangle A3 = new(0, 0, 842, 1191);
    public static readonly Rectangle A5 = new(0, 0, 420, 595);
    public static readonly Rectangle Letter = new(0, 0, 612, 792);
    public static readonly Rectangle Legal = new(0, 0, 612, 1008);

    public Rectangle(float left, float bottom, float right, float top)
    {
        Left = left;
        Bottom = bottom;
        Right = right;
        Top = top;
    }

    public Rectangle(float width, float height)
    {
        Left = 0;
        Bottom = 0;
        Right = width;
        Top = height;
    }

    public Rectangle Rotate() => new(Bottom, Left, Top, Right);

    public Rectangle Clone() => new(Left, Bottom, Right, Top);

    public override string ToString() => $"Rectangle[({Left}, {Bottom}) - ({Right}, {Top})]";

    // ?????
    public const int NO_BORDER = 0;
}
