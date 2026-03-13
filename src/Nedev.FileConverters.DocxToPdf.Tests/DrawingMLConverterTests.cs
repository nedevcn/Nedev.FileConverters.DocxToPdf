using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Converters;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.Models;
using System.IO;
using Xunit;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class DrawingMLConverterTests
    {
        [Fact]
        public void ConvertDrawing_WithGraphicData_ReturnsElement()
        {
            // 创建测试文档
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // 添加图片部分
                var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                // 创建简单的 1x1 像素 PNG
                var pngBytes = new byte[] {
                    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
                    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
                    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
                    0x54, 0x08, 0x99, 0x63, 0xF8, 0x0F, 0x00, 0x00,
                    0x01, 0x01, 0x00, 0x05, 0x18, 0xD8, 0x4E, 0x00,
                    0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
                    0x42, 0x60, 0x82
                };
                using (var imageStream = new MemoryStream(pngBytes))
                {
                    imagePart.FeedData(imageStream);
                }

                // 创建 DrawingML Inline 图片
                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                var inline = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
                var extent = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 914400, Cy = 914400 }; // 1 inch in EMUs
                inline.Append(extent);

                var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
                var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };
                var pic = new DocumentFormat.OpenXml.Drawing.Picture();
                var blipFill = new DocumentFormat.OpenXml.Drawing.BlipFill();
                var blip = new DocumentFormat.OpenXml.Drawing.Blip { Embed = mainPart.GetIdOfPart(imagePart) };
                blipFill.Append(blip);
                pic.Append(blipFill);
                graphicData.Append(pic);
                graphic.Append(graphicData);
                inline.Append(graphic);
                drawing.Append(inline);

                mainPart.Document.Save();

                // 测试转换
                var options = new ConvertOptions();
                var fontHelper = new FontHelper(options);
                var converter = new DrawingMLConverter(doc, fontHelper);
                var result = converter.ConvertDrawing(drawing, 500f);

                Assert.NotNull(result);
            }
        }

        [Fact]
        public void ConvertDrawing_EmptyDrawing_ReturnsNull()
        {
            // 创建测试文档
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // 创建空的 Drawing
                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();

                mainPart.Document.Save();

                // 测试转换
                var options = new ConvertOptions();
                var fontHelper = new FontHelper(options);
                var converter = new DrawingMLConverter(doc, fontHelper);
                var result = converter.ConvertDrawing(drawing, 500f);

                Assert.Null(result);
            }
        }

        [Fact]
        public void ConvertDrawing_ShapeWithText_ReturnsParagraphElement()
        {
            // 创建测试文档
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // 创建 DrawingML 形状
                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                var inline = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
                var extent = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 1371600, Cy = 457200 }; // 1.5 x 0.5 inch in EMUs
                inline.Append(extent);

                var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
                var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
                var shape = new DocumentFormat.OpenXml.Drawing.Shape();
                var txBody = new DocumentFormat.OpenXml.Drawing.TextBody();
                var bodyPr = new DocumentFormat.OpenXml.Drawing.BodyProperties();
                txBody.Append(bodyPr);

                var para = new DocumentFormat.OpenXml.Drawing.Paragraph();
                var run = new DocumentFormat.OpenXml.Drawing.Run();
                var text = new DocumentFormat.OpenXml.Drawing.Text("Hello DrawingML");
                run.Append(text);
                para.Append(run);
                txBody.Append(para);
                shape.Append(txBody);
                graphicData.Append(shape);
                graphic.Append(graphicData);
                inline.Append(graphic);
                drawing.Append(inline);

                mainPart.Document.Save();

                // 测试转换
                var options = new ConvertOptions();
                var fontHelper = new FontHelper(options);
                var converter = new DrawingMLConverter(doc, fontHelper);
                var result = converter.ConvertDrawing(drawing, 500f);

                Assert.NotNull(result);
            }
        }

        [Fact]
        public void ConvertDrawing_AnchoredShapeWithText_ReturnsParagraphElement()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // create a drawing that is anchored rather than inline
                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                var anchor = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor();
                var extent = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 1371600, Cy = 457200 };
                anchor.Append(extent);

                var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
                var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
                var shape = new DocumentFormat.OpenXml.Drawing.Shape();
                var txBody = new DocumentFormat.OpenXml.Drawing.TextBody();
                txBody.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());

                var para = new DocumentFormat.OpenXml.Drawing.Paragraph();
                var run = new DocumentFormat.OpenXml.Drawing.Run();
                run.Append(new DocumentFormat.OpenXml.Drawing.Text("Anchored text"));
                para.Append(run);
                txBody.Append(para);
                shape.Append(txBody);
                graphicData.Append(shape);
                graphic.Append(graphicData);
                anchor.Append(graphic);
                drawing.Append(anchor);

                mainPart.Document.Save();

                var options = new ConvertOptions();
                var fontHelper = new FontHelper(options);
                var converter = new DrawingMLConverter(doc, fontHelper);
                var result = converter.ConvertDrawing(drawing, 500f);
                Assert.NotNull(result);
                // anchored element should not be floating object
                Assert.IsNotType<FloatingObject>(result);
            }
        }

        [Fact]
        public void ConvertDrawing_StyledText_RespectsFontAndAlignment()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body());

                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                var inline = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
                var extent = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 914400, Cy = 228600 };
                inline.Append(extent);

                var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
                var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
                var shape = new DocumentFormat.OpenXml.Drawing.Shape();
                var txBody = new DocumentFormat.OpenXml.Drawing.TextBody();
                txBody.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());

                var para = new DocumentFormat.OpenXml.Drawing.Paragraph();
                para.Append(new DocumentFormat.OpenXml.Drawing.ParagraphProperties(
                    new DocumentFormat.OpenXml.Drawing.Alignment { Val = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center }
                ));
                
                // first run: bold green 24pt
                var run1 = new DocumentFormat.OpenXml.Drawing.Run();
                var runPr1 = new DocumentFormat.OpenXml.Drawing.RunProperties();
                runPr1.FontSize = new DocumentFormat.OpenXml.Drawing.FontSize { Val = 2400 };
                runPr1.Bold = new DocumentFormat.OpenXml.Drawing.Bold();
                runPr1.Append(new DocumentFormat.OpenXml.Drawing.SolidFill(
                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "00FF00" }
                ));
                run1.Append(runPr1);
                run1.Append(new DocumentFormat.OpenXml.Drawing.Text("Styled"));
                para.Append(run1);

                // second run: italic, underline, SimSun font
                var run2 = new DocumentFormat.OpenXml.Drawing.Run();
                var runPr2 = new DocumentFormat.OpenXml.Drawing.RunProperties();
                runPr2.Italic = new DocumentFormat.OpenXml.Drawing.Italic();
                runPr2.Underline = new DocumentFormat.OpenXml.Drawing.TextUnderline { Val = DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single };
                var latin = new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "SimSun" };
                runPr2.Append(latin);
                run2.Append(runPr2);
                run2.Append(new DocumentFormat.OpenXml.Drawing.Text(" ItalicUnder"));
                para.Append(run2);

                txBody.Append(para);
                shape.Append(txBody);
                graphicData.Append(shape);
                graphic.Append(graphicData);
                inline.Append(graphic);
                drawing.Append(inline);

                main.Document.Save();

                var options = new ConvertOptions();
                var fontHelper = new FontHelper(options);
                var converter = new DrawingMLConverter(doc, fontHelper);
                var result = converter.ConvertDrawing(drawing, 500f);

                Assert.IsType<iTextParagraph>(result);
                var pdfPara = (iTextParagraph)result;
                Assert.Equal(Element.ALIGN_CENTER, pdfPara.Alignment);
                Assert.Equal(2, pdfPara.Chunks.Count);

                var c1 = pdfPara.Chunks[0];
                Assert.Equal(24f, c1.Font.Size, 2);
                Assert.True((c1.Font.Style & iTextFont.BOLD) != 0);
                Assert.Equal(BaseColor.GREEN, c1.Font.Color);

                var c2 = pdfPara.Chunks[1];
                Assert.True((c2.Font.Style & iTextFont.ITALIC) != 0);
                Assert.True(c2.HasUnderline);
                Assert.Contains("SimSun", c2.Font.FamilyName, StringComparison.OrdinalIgnoreCase);

                // also check pdf stream contains italic indicator and SimSun font name
                using var outStream = new MemoryStream();
                var writer = new PdfWriter(outStream);
                var ct = new ColumnText(writer.DirectContent);
                ct.SetSimpleColumn(0, 0, 500, 500);
                ct.AddElement(pdfPara);
                ct.Go();
                var pdfText = System.Text.Encoding.ASCII.GetString(outStream.ToArray());
                Assert.Contains("SimSun", pdfText, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Italic", pdfText, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("underline", pdfText, StringComparison.OrdinalIgnoreCase);

                // ensure text operators exist (a lightweight check)
                Assert.Matches(@"\d+\s+Tf", pdfText);
                Assert.Matches(@"TJ|Tj", pdfText);
            }
        }

        [Fact]
        public void ConvertDrawing_GradientAndSmallCaps_AreRespected()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body());

                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                var inline = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
                var extent = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 914400, Cy = 228600 };
                inline.Append(extent);
                var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
                var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
                var shape = new DocumentFormat.OpenXml.Drawing.Shape();
                var txBody = new DocumentFormat.OpenXml.Drawing.TextBody();
                txBody.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());

                var para = new DocumentFormat.OpenXml.Drawing.Paragraph();
                var run = new DocumentFormat.OpenXml.Drawing.Run();
                var runPr = new DocumentFormat.OpenXml.Drawing.RunProperties();
                // gradient fill with two stops
                var grad = new DocumentFormat.OpenXml.Drawing.GradientFill();
                grad.Append(new DocumentFormat.OpenXml.Drawing.GradientStop(
                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FF0000" }
                ));
                grad.Append(new DocumentFormat.OpenXml.Drawing.GradientStop(
                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "0000FF" }
                ));
                runPr.Append(grad);
                runPr.SmallCaps = new DocumentFormat.OpenXml.Drawing.SmallCaps();
                run.Append(runPr);
                run.Append(new DocumentFormat.OpenXml.Drawing.Text("smallcaps"));
                para.Append(run);
                txBody.Append(para);
                shape.Append(txBody);
                graphicData.Append(shape);
                graphic.Append(graphicData);
                inline.Append(graphic);
                drawing.Append(inline);

                main.Document.Save();

                var options = new ConvertOptions();
                var fontHelper = new FontHelper(options);
                var converter = new DrawingMLConverter(doc, fontHelper);
                var result = converter.ConvertDrawing(drawing, 500f);

                var pdfPara = (iTextParagraph)result;
                var chunk = pdfPara.Chunks[0];
                Assert.Equal("SMALLCAPS", chunk.Content);
                // size should be reduced
                Assert.True(chunk.Font.Size < options.DefaultFontSize);
                Assert.Equal(new BaseColor(0x7F,0x00,0x7F), chunk.Font.Color); // average of red/blue
            }
        }

        [Fact]
        public void ConvertDrawing_CharSpacing_BidiAndField_AreApplied()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body());
                // declare a document variable
                var settings = main.AddNewPart<DocumentSettingsPart>();
                settings.Settings = new Settings(
                    new DocumentVariables(
                        new DocumentVariable { Name="Foo", Val="Bar" }
                    )
                );

                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                var inline = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
                var extent = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 914400, Cy = 228600 };
                inline.Append(extent);
                var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
                var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
                var shape = new DocumentFormat.OpenXml.Drawing.Shape();
                var txBody = new DocumentFormat.OpenXml.Drawing.TextBody();
                txBody.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());

                var para = new DocumentFormat.OpenXml.Drawing.Paragraph();
                var run = new DocumentFormat.OpenXml.Drawing.Run();
                var runPr = new DocumentFormat.OpenXml.Drawing.RunProperties();
                // set char spacing 2000 units (2pt assumed)
                runPr.CharacterSpacing = new DocumentFormat.OpenXml.Drawing.CharacterSpacing { Val = 2000 };
                runPr.Language = new DocumentFormat.OpenXml.Drawing.Language { Val = "ar-SA" };
                run.Append(runPr);
                run.Append(new DocumentFormat.OpenXml.Drawing.Text("Foo"));
                para.Append(run);
                txBody.Append(para);
                shape.Append(txBody);
                graphicData.Append(shape);
                graphic.Append(graphicData);
                inline.Append(graphic);
                drawing.Append(inline);

                main.Document.Save();

                var options = new ConvertOptions();
                var fontHelper = new FontHelper(options);
                var converter = new DrawingMLConverter(doc, fontHelper);
                var result = converter.ConvertDrawing(drawing, 500f);

                var pdfPara = (iTextParagraph)result;
                var chunk = pdfPara.Chunks[0];
                // bidi reversed 'Foo'->'ooF'
                Assert.Equal("ooF", chunk.Content);
                // char spacing set properly
                Assert.InRange(chunk.CharSpacing, 1.9f, 2.1f);

                // verify field resolution: since chunk.Text started as Foo (variable name)
                // ResolveField should have replaced it earlier; value now 'Bar'
                Assert.Contains("Bar", chunk.Content);
            }
        }
    }
}
