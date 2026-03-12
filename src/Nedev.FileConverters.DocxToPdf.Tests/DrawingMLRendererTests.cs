using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Rendering;
using Nedev.FileConverters.DocxToPdf.Models;
using System.IO;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class DrawingMLRendererTests
    {
        [Fact]
        public void RenderToPng_ShapeWithSolidFill_ReturnsPngBytes()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // 创建 DrawingML 形状
                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                var inline = new DW.Inline();
                var extent = new DW.Extent { Cx = 914400, Cy = 914400 }; // 1x1 inch
                inline.Append(extent);

                var graphic = new A.Graphic();
                var graphicData = new A.GraphicData { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
                var shape = new A.Shape();
                var spPr = new A.ShapeProperties();

                // 设置变换
                var xfrm = new A.Transform2D();
                xfrm.Append(new A.Offset { X = 0, Y = 0 });
                xfrm.Append(new A.Extents { Cx = 914400, Cy = 914400 });
                spPr.Append(xfrm);

                // 设置预设几何形状（矩形）
                var prstGeom = new A.PresetGeometry();
                prstGeom.Append(new A.AdjustValueList());
                prstGeom.Preset = new EnumValue<A.ShapeTypeValues>(A.ShapeTypeValues.Rectangle);
                spPr.Append(prstGeom);

                // 设置填充颜色
                var solidFill = new A.SolidFill();
                solidFill.Append(new A.RgbColorModelHex { Val = "4472C4" }); // 蓝色
                spPr.Append(solidFill);

                shape.Append(spPr);
                graphicData.Append(shape);
                graphic.Append(graphicData);
                inline.Append(graphic);
                drawing.Append(inline);

                mainPart.Document.Save();

                // 测试渲染
                var options = new ConvertOptions();
                var renderer = new DrawingMLRenderer(doc, options);
                var pngBytes = renderer.RenderToPng(drawing, 96, 96);

                Assert.NotNull(pngBytes);
                Assert.True(pngBytes.Length > 0);
                // 验证 PNG 文件头
                Assert.Equal(0x89, pngBytes[0]);
                Assert.Equal(0x50, pngBytes[1]);
                Assert.Equal(0x4E, pngBytes[2]);
                Assert.Equal(0x47, pngBytes[3]);
            }
        }

        [Fact]
        public void RenderToPng_EllipseShape_ReturnsPngBytes()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                var inline = new DW.Inline();
                var extent = new DW.Extent { Cx = 914400, Cy = 914400 };
                inline.Append(extent);

                var graphic = new A.Graphic();
                var graphicData = new A.GraphicData { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
                var shape = new A.Shape();
                var spPr = new A.ShapeProperties();

                var xfrm = new A.Transform2D();
                xfrm.Append(new A.Offset { X = 0, Y = 0 });
                xfrm.Append(new A.Extents { Cx = 914400, Cy = 914400 });
                spPr.Append(xfrm);

                // 椭圆形状
                var prstGeom = new A.PresetGeometry();
                prstGeom.Append(new A.AdjustValueList());
                prstGeom.Preset = new EnumValue<A.ShapeTypeValues>(A.ShapeTypeValues.Ellipse);
                spPr.Append(prstGeom);

                var solidFill = new A.SolidFill();
                solidFill.Append(new A.RgbColorModelHex { Val = "E7E6E6" });
                spPr.Append(solidFill);

                shape.Append(spPr);
                graphicData.Append(shape);
                graphic.Append(graphicData);
                inline.Append(graphic);
                drawing.Append(inline);

                mainPart.Document.Save();

                var options = new ConvertOptions();
                var renderer = new DrawingMLRenderer(doc, options);
                var pngBytes = renderer.RenderToPng(drawing, 96, 96);

                Assert.NotNull(pngBytes);
                Assert.True(pngBytes.Length > 0);
            }
        }

        [Fact]
        public void RenderToPng_ShapeWithText_ReturnsPngBytes()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                var inline = new DW.Inline();
                var extent = new DW.Extent { Cx = 1371600, Cy = 457200 }; // 1.5 x 0.5 inch
                inline.Append(extent);

                var graphic = new A.Graphic();
                var graphicData = new A.GraphicData { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
                var shape = new A.Shape();

                // 形状属性
                var spPr = new A.ShapeProperties();
                var xfrm = new A.Transform2D();
                xfrm.Append(new A.Offset { X = 0, Y = 0 });
                xfrm.Append(new A.Extents { Cx = 1371600, Cy = 457200 });
                spPr.Append(xfrm);

                var prstGeom = new A.PresetGeometry();
                prstGeom.Append(new A.AdjustValueList());
                prstGeom.Preset = new EnumValue<A.ShapeTypeValues>(A.ShapeTypeValues.Rectangle);
                spPr.Append(prstGeom);

                var solidFill = new A.SolidFill();
                solidFill.Append(new A.RgbColorModelHex { Val = "FFFFFF" });
                spPr.Append(solidFill);

                shape.Append(spPr);

                // 文本体
                var txBody = new A.TextBody();
                txBody.Append(new A.BodyProperties());
                txBody.Append(new A.ListStyle());

                var para = new A.Paragraph();
                var run = new A.Run();
                run.Append(new A.Text("Test Text"));
                para.Append(run);
                txBody.Append(para);

                shape.Append(txBody);
                graphicData.Append(shape);
                graphic.Append(graphicData);
                inline.Append(graphic);
                drawing.Append(inline);

                mainPart.Document.Save();

                var options = new ConvertOptions();
                var renderer = new DrawingMLRenderer(doc, options);
                var pngBytes = renderer.RenderToPng(drawing, 144, 48);

                Assert.NotNull(pngBytes);
                Assert.True(pngBytes.Length > 0);
            }
        }

        [Fact]
        public void RenderToPng_EmptyDrawing_ReturnsNull()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                mainPart.Document.Save();

                var options = new ConvertOptions();
                var renderer = new DrawingMLRenderer(doc, options);
                var pngBytes = renderer.RenderToPng(drawing, 96, 96);

                Assert.Null(pngBytes);
            }
        }

        [Fact]
        public void EMU_ToPixels_ConvertsCorrectly()
        {
            // 914400 EMU = 1 inch = 96 pixels (at 96 DPI)
            var pixels = EMU.ToPixels(914400);
            Assert.Equal(96f, pixels);

            // 0 EMU = 0 pixels
            var zeroPixels = EMU.ToPixels(0);
            Assert.Equal(0f, zeroPixels);
        }

        [Fact]
        public void EMU_FromPixels_ConvertsCorrectly()
        {
            // 96 pixels = 914400 EMU (at 96 DPI)
            var emu = EMU.FromPixels(96f);
            Assert.Equal(914400L, emu);

            // 0 pixels = 0 EMU
            var zeroEmu = EMU.FromPixels(0f);
            Assert.Equal(0L, zeroEmu);
        }
    }
}
