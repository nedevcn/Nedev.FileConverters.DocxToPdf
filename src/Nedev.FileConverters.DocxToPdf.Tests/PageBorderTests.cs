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
}
