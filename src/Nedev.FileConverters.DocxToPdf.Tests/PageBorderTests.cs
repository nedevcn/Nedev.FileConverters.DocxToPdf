using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf;
using Nedev.FileConverters.DocxToPdf.Models;
using Xunit;

namespace Nedev.FileConverters.DocxToPdf.Tests;

public class PageBorderTests
{
    [Fact]
    public void PageBorder_Parsing_Works()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Hello Page Borders"))),
                new SectionProperties(
                    new PageBorders(
                        new TopBorder { Val = BorderValues.Double, Size = 24, Color = "FF0000", Space = 4 },
                        new BottomBorder { Val = BorderValues.Double, Size = 24, Color = "00FF00", Space = 4 },
                        new LeftBorder { Val = BorderValues.Double, Size = 24, Color = "0000FF", Space = 4 },
                        new RightBorder { Val = BorderValues.Double, Size = 24, Color = "FFFF00", Space = 4 }
                    ) { OffsetFrom = PageBorderOffsetValues.Page }
                )
            ));
        }

        ms.Position = 0;
        var converter = new DocxToPdfConverter();
        using var output = new MemoryStream();
        converter.Convert(ms, output);
        
        Assert.True(output.Length > 0);
        // We can't easily verify the drawing commands in the PDF without a heavy PDF parser,
        // but successful conversion means no crashes during border rendering.
    }

    [Fact]
    public void AnchoredTextbox_ConvertedAsText()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());

            // create a paragraph containing a drawing anchor with textual content
            var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
            var anchor = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor();
            // minimal required children: extent and graphic/a:graphicData/shape with text
            var extent = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 914400, Cy = 457200 };
            anchor.Append(extent);

            var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
            var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
            var shape = new DocumentFormat.OpenXml.Drawing.Shape();
            var txBody = new DocumentFormat.OpenXml.Drawing.TextBody();
            txBody.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());
            var paraText = new DocumentFormat.OpenXml.Drawing.Paragraph();
            var run = new DocumentFormat.OpenXml.Drawing.Run();
            run.Append(new DocumentFormat.OpenXml.Drawing.Text("Anchored text"));
            paraText.Append(run);
            txBody.Append(paraText);
            shape.Append(txBody);
            graphicData.Append(shape);
            graphic.Append(graphicData);
            anchor.Append(graphic);
            drawing.Append(anchor);

            main.Document.Body.Append(new Paragraph(new Run(drawing)));
            main.Document.Save();
        }

        ms.Position = 0;
        var converter = new DocxToPdfConverter();
        using var output = new MemoryStream();
        converter.Convert(ms, output);

        var pdfString = System.Text.Encoding.ASCII.GetString(output.ToArray());
        Assert.Contains("Anchored text", pdfString);
    }
}
