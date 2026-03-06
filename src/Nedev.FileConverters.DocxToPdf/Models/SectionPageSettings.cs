using Nedev.FileConverters.DocxToPdf.PdfEngine;

namespace Nedev.FileConverters.DocxToPdf.Models;

public class SectionPageSettings
{
    public Rectangle PageSize { get; set; } = Rectangle.A4;
    public float MarginLeft { get; set; }
    public float MarginRight { get; set; }
    public float MarginTop { get; set; }
    public float MarginBottom { get; set; }
    public float HeaderDistance { get; set; }
    public float FooterDistance { get; set; }
}
