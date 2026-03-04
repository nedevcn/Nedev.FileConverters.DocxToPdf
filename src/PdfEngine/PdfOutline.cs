namespace Nedev.DocxToPdf.PdfEngine;

/// <summary>
/// PDF书签目标
/// </summary>
public class PdfDestination
{
    public const int XYZ = 1;
    public const int FIT = 2;
    public const int FITH = 3;
    public const int FITV = 4;
    public const int FITR = 5;
    public const int FITB = 6;
    public const int FITBH = 7;
    public const int FITBV = 8;

    public int Type { get; }
    public float Left { get; }
    public float Top { get; }
    public float Zoom { get; }

    public PdfDestination(int type, float left = 0, float top = 0, float zoom = 0)
    {
        Type = type;
        Left = left;
        Top = top;
        Zoom = zoom;
    }
}

/// <summary>
/// PDF书签
/// </summary>
public class PdfOutline
{
    public string Title { get; }
    public PdfDestination Destination { get; }
    public PdfOutline? Parent { get; }
    public List<PdfOutline> Children { get; } = [];
    public bool Open { get; set; }

    public PdfOutline(PdfOutline? parent, PdfDestination destination, string title, bool open = false)
    {
        Parent = parent;
        Destination = destination;
        Title = title;
        Open = open;

        parent?.Children.Add(this);
    }
}
