using System.Text;

namespace Nedev.DocxToPdf.PdfEngine;

public class PdfAnnotation
{
    public int PageNumber { get; set; }
    public float X { get; set; }
    public float Y { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
    public string Action { get; set; } = "";
    public string? Dest { get; set; }
    public string? Uri { get; set; }

    public PdfAnnotation(int pageNumber, float x, float y, float width, float height)
    {
        PageNumber = pageNumber;
        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    public string ToPdfDict(int objectNumber)
    {
        var sb = new StringBuilder();
        sb.Append($"{objectNumber} 0 obj\n");
        sb.Append("<<\n");
        sb.Append("/Type /Annot\n");
        sb.Append("/Subtype /Link\n");
        sb.Append($"/Rect [{X:F2} {Y:F2} {X + Width:F2} {Y + Height:F2}]\n");

        if (!string.IsNullOrEmpty(Uri))
        {
            sb.Append("/A <<\n");
            sb.Append("/Type /Action\n");
            sb.Append("/S /URI\n");
            sb.Append($"/URI ({Uri})\n");
            sb.Append(">>\n");
        }
        else if (!string.IsNullOrEmpty(Dest))
        {
            sb.Append($"/Dest [{PageNumber} 0 R /Fit]\n");
        }

        sb.Append(">>\n");
        sb.Append("endobj\n");
        return sb.ToString();
    }
}

public class AnnotationCollection
{
    private readonly List<PdfAnnotation> _annotations = [];
    private readonly PdfDocument? _document;

    public AnnotationCollection(PdfDocument? document = null)
    {
        _document = document;
    }

    public void AddAnnotation(PdfAnnotation annotation)
    {
        _annotations.Add(annotation);
    }

    public void AddLink(int pageNumber, float x, float y, float width, float height, string uri)
    {
        var annotation = new PdfAnnotation(pageNumber, x, y, width, height)
        {
            Uri = uri
        };
        _annotations.Add(annotation);
    }

    public void AddInternalLink(int pageNumber, float x, float y, float width, float height, string dest)
    {
        var annotation = new PdfAnnotation(pageNumber, x, y, width, height)
        {
            Dest = dest
        };
        _annotations.Add(annotation);
    }

    public List<PdfAnnotation> GetAnnotationsForPage(int pageNumber)
    {
        return _annotations.Where(a => a.PageNumber == pageNumber).ToList();
    }

    public int Count => _annotations.Count;
}
