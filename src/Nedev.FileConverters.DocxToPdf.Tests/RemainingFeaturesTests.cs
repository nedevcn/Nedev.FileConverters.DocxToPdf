using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf;
using Nedev.FileConverters.DocxToPdf.Converters;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.Models;
using Nedev.FileConverters.DocxToPdf.Rendering;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class RemainingFeaturesTests
    {
        #region SHAPEDOG Field Tests

        [Theory]
        [InlineData("SHAPEDOG \\s", "1")]      // 页码
        [InlineData("SHAPEDOG \\p", "1")]      // 段落号
        [InlineData("SHAPEDOG \\r 5", "5")]    // 相对页码
        [InlineData("SHAPEDOG", "1")]          // 默认
        public void ResolveField_Shapedog_ReturnsExpectedValue(string instruction, string expected)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                var result = converter.ResolveField(instruction, doc, "C:\\test.docx");

                Assert.Equal(expected, result);
            }
        }

        [Fact]
        public void ResolveField_Shapedog_EmptyInstruction_ReturnsDefaultValue()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                var result = converter.ResolveField("SHAPEDOG   ", doc, "C:\\test.docx");

                // SHAPEDOG 默认返回页码 "1"
                Assert.Equal("1", result);
            }
        }

        #endregion

        #region LISTNUM Field Tests

        [Theory]
        [InlineData("LISTNUM", "1.")]
        [InlineData("LISTNUM \\l 0", "1.")]
        [InlineData("LISTNUM \\l 1", "  a.")]
        [InlineData("LISTNUM \\l 2", "    i.")]
        [InlineData("LISTNUM \\s 5", "5.")]
        [InlineData("LISTNUM MyList \\l 0 \\s 3", "3.")]
        public void ResolveField_Listnum_ReturnsExpectedValue(string instruction, string expected)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                var result = converter.ResolveField(instruction, doc, "C:\\test.docx");

                Assert.Equal(expected, result);
            }
        }

        [Fact]
        public void ResolveField_Listnum_EmptyInstruction_ReturnsDefaultValue()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                var result = converter.ResolveField("LISTNUM   ", doc, "C:\\test.docx");

                // LISTNUM 默认返回 "1."
                Assert.Equal("1.", result);
            }
        }

        [Fact]
        public void ResolveField_Listnum_InvalidLevel_ReturnsDefault()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                var result = converter.ResolveField("LISTNUM \\l 10", doc, "C:\\test.docx");

                // 对于无效级别，应该返回默认格式
                Assert.NotNull(result);
                Assert.True(result.EndsWith("."));
            }
        }

        [Fact]
        public void ResolveField_IfField_EvaluatesTrueAndFalse()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                Assert.Equal("yes", converter.ResolveField("IF 1=1 yes no", doc));
                Assert.Equal("no", converter.ResolveField("IF 2=1 yes no", doc));
                Assert.Equal("Bar", converter.ResolveField("IF Foo=Foo yes no", doc));
            }
        }

        #endregion

        #region SmartArt Renderer Tests

        [Fact]
        public void SmartArtData_DefaultValues_AreCorrect()
        {
            var data = new SmartArtData();

            Assert.Equal(string.Empty, data.Title);
            Assert.Equal(SmartArtLayoutType.Other, data.LayoutType);
            Assert.Empty(data.Nodes);
        }

        [Fact]
        public void SmartArtNode_DefaultValues_AreCorrect()
        {
            var node = new SmartArtNode();

            Assert.Equal(string.Empty, node.Id);
            Assert.Equal(string.Empty, node.Text);
            Assert.Equal(0, node.Level);
            Assert.Equal(string.Empty, node.ParentId);
        }

        [Fact]
        public void SmartArtLayoutType_Enum_HasExpectedValues()
        {
            Assert.Equal(0, (int)SmartArtLayoutType.Hierarchy);
            Assert.Equal(1, (int)SmartArtLayoutType.Process);
            Assert.Equal(2, (int)SmartArtLayoutType.Cycle);
            Assert.Equal(3, (int)SmartArtLayoutType.Matrix);
            Assert.Equal(4, (int)SmartArtLayoutType.Pyramid);
            Assert.Equal(5, (int)SmartArtLayoutType.Other);
        }

        [Fact]
        public void SmartArtData_WithNodes_WorksCorrectly()
        {
            var data = new SmartArtData
            {
                Title = "Organization Chart",
                LayoutType = SmartArtLayoutType.Hierarchy
            };

            data.Nodes.Add(new SmartArtNode
            {
                Id = "1",
                Text = "CEO",
                Level = 0
            });

            data.Nodes.Add(new SmartArtNode
            {
                Id = "2",
                Text = "Manager",
                Level = 1,
                ParentId = "1"
            });

            Assert.Equal("Organization Chart", data.Title);
            Assert.Equal(SmartArtLayoutType.Hierarchy, data.LayoutType);
            Assert.Equal(2, data.Nodes.Count);
            Assert.Equal("CEO", data.Nodes[0].Text);
            Assert.Equal("Manager", data.Nodes[1].Text);
        }

        [Fact]
        public void SmartArtRenderer_Constructor_InitializesCorrectly()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var options = new ConvertOptions();
                var renderer = new SmartArtRenderer(doc, options);

                Assert.NotNull(renderer);
            }
        }

        #endregion

        #region Integration Tests

        [Fact]
        public void AllFieldTypes_AreResolvable()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // 添加 DocumentSettingsPart 用于 DOCVARIABLE 测试
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(
                    new DocumentVariables(
                        new DocumentVariable { Name = "TestVar", Val = "TestValue" }
                    )
                );

                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();

                // 测试所有字段类型
                // PAGE 字段返回 null（需要运行时动态获取，由 PageNumberProvider 处理）
                Assert.Null(converter.ResolveField("PAGE", doc, "C:\\test.docx"));
                // DATE 字段返回当前日期
                Assert.NotNull(converter.ResolveField("DATE", doc, "C:\\test.docx"));
                // AUTHOR 可能返回 null（如果文档没有设置作者）
                var authorResult = converter.ResolveField("AUTHOR", doc, "C:\\test.docx");
                Assert.True(authorResult == null || authorResult is string);
                // TITLE 可能返回 null
                var titleResult = converter.ResolveField("TITLE", doc, "C:\\test.docx");
                Assert.True(titleResult == null || titleResult is string);
                // FILENAME 应该返回文件名
                Assert.Equal("test", converter.ResolveField("FILENAME", doc, "C:\\test.docx"));
                // DOCVARIABLE 应该返回变量值
                Assert.Equal("TestValue", converter.ResolveField("DOCVARIABLE TestVar", doc, "C:\\test.docx"));
                // SHAPEDOG 默认返回页码
                Assert.Equal("1", converter.ResolveField("SHAPEDOG", doc, "C:\\test.docx"));
                // LISTNUM 默认返回 "1."
                Assert.Equal("1.", converter.ResolveField("LISTNUM", doc, "C:\\test.docx"));
                // SYMBOL 65 = 'A'
                Assert.Equal("A", converter.ResolveField("SYMBOL 65", doc, "C:\\test.docx"));
                // SECTION 字段返回 null（需要运行时动态获取）
                Assert.Null(converter.ResolveField("SECTION", doc, "C:\\test.docx"));
                // SECTIONPAGES 字段返回 null（需要运行时动态获取）
                Assert.Null(converter.ResolveField("SECTIONPAGES", doc, "C:\\test.docx"));
            }
        }

        #endregion

        #region Page Number Field Tests

        [Fact]
        public void PageNumberProvider_ReturnsCorrectValues()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                var options = new ConvertOptions();
                var fontHelper = new FontHelper(options);
                var paragraphConverter = new ParagraphConverter(fontHelper);

                // 设置页码提供者
                paragraphConverter.PageNumberProvider = () => (5, 10);

                // 验证提供者返回正确的值
                var pageInfo = paragraphConverter.PageNumberProvider?.Invoke();
                Assert.True(pageInfo.HasValue);
                Assert.Equal(5, pageInfo.Value.Current);
                Assert.Equal(10, pageInfo.Value.Total);
            }
        }

        [Fact]
        public void SectionInfoProvider_ReturnsCorrectValues()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                var options = new ConvertOptions();
                var fontHelper = new FontHelper(options);
                var paragraphConverter = new ParagraphConverter(fontHelper);

                // 设置节信息提供者
                paragraphConverter.SectionInfoProvider = () => (1, 3, 5);

                // 验证提供者返回正确的值
                var sectionInfo = paragraphConverter.SectionInfoProvider?.Invoke();
                Assert.True(sectionInfo.HasValue);
                Assert.Equal(1, sectionInfo.Value.SectionIndex);
                Assert.Equal(3, sectionInfo.Value.PageInSection);
                Assert.Equal(5, sectionInfo.Value.TotalPagesInSection);
            }
        }

        [Theory]
        [InlineData("PAGE")]
        [InlineData("NUMPAGES")]
        public void ResolveField_PageFields_ReturnsNull_ForRuntimeProcessing(string fieldCode)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                // 直接测试 ResolveField 方法，它应该返回 null（因为需要运行时处理）
                var result = converter.ResolveField(fieldCode, doc, "C:\\test.docx");
                Assert.Null(result);
            }
        }

        [Theory]
        [InlineData("SECTION")]
        [InlineData("SECTIONPAGES")]
        public void ResolveField_SectionFields_ReturnsNull_ForRuntimeProcessing(string fieldCode)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                // 直接测试 ResolveField 方法，它应该返回 null（因为需要运行时处理）
                var result = converter.ResolveField(fieldCode, doc, "C:\\test.docx");
                Assert.Null(result);
            }
        }

        #endregion

        #region New Feature Tests - Remaining 2%

        [Fact]
        public void ChartRenderer_InitializesCorrectly()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var options = new ConvertOptions();
                var renderer = new ChartRenderer(doc, options);

                Assert.NotNull(renderer);
            }
        }

        [Fact]
        public void ChartType_Enum_HasExpectedValues()
        {
            Assert.Equal(0, (int)ChartType.Bar);
            Assert.Equal(1, (int)ChartType.Column);
            Assert.Equal(2, (int)ChartType.Line);
            Assert.Equal(3, (int)ChartType.Pie);
            Assert.Equal(4, (int)ChartType.Area);
            Assert.Equal(5, (int)ChartType.Scatter);
            Assert.Equal(6, (int)ChartType.Radar);
        }

        [Fact]
        public void SmartArtRenderer_InitializesCorrectly()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var options = new ConvertOptions();
                var renderer = new SmartArtRenderer(doc, options);

                Assert.NotNull(renderer);
            }
        }

        [Fact]
        public void SmartArtLayoutType_Pyramid_IsDefined()
        {
            Assert.Equal(4, (int)SmartArtLayoutType.Pyramid);
        }

        [Fact]
        public void OMMLRenderer_InitializesCorrectly()
        {
            var renderer = new OMMLRenderer(16f);
            Assert.NotNull(renderer);
        }

        #endregion
    }
}
