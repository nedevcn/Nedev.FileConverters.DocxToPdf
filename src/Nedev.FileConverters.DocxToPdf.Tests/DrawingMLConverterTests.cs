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
                var run = new DocumentFormat.OpenXml.Drawing.Run();
                var runPr = new DocumentFormat.OpenXml.Drawing.RunProperties();
                runPr.FontSize = new DocumentFormat.OpenXml.Drawing.FontSize { Val = 2400 }; // 24pt
                runPr.Bold = new DocumentFormat.OpenXml.Drawing.Bold();
                var sc = new DocumentFormat.OpenXml.Drawing.SolidFill(
                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "00FF00" }
                );
                runPr.Append(sc);
                run.Append(runPr);
                run.Append(new DocumentFormat.OpenXml.Drawing.Text("Styled"));
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

                Assert.IsType<iTextParagraph>(result);
                var pdfPara = (iTextParagraph)result;
                Assert.Equal(Element.ALIGN_CENTER, pdfPara.Alignment);
                Assert.Single(pdfPara.Chunks);
                var chunk = pdfPara.Chunks[0];
                Assert.Equal(24f, chunk.Font.Size, 2);
                Assert.True((chunk.Font.Style & iTextFont.BOLD) != 0);
                Assert.Equal(BaseColor.GREEN, chunk.Font.Color);
            }
        }
    }
}
