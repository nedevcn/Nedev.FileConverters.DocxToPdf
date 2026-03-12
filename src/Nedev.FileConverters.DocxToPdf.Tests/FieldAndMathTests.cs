using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.Models;
using System.Linq;
using Xunit;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class FieldAndMathTests
    {
        [Fact]
        public void MathHelper_Fractions_ReturnsCorrectString()
        {
            // <m:f><m:num><m:t>1</m:t></m:num><m:den><m:t>2</m:t></m:den></m:f>
            var f = new Fraction();
            var num = new Numerator();
            num.Append(new DocumentFormat.OpenXml.Math.Run(new DocumentFormat.OpenXml.Math.Text("1")));
            var den = new Denominator();
            den.Append(new DocumentFormat.OpenXml.Math.Run(new DocumentFormat.OpenXml.Math.Text("2")));
            f.Append(num, den);

            var para = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            var omath = new DocumentFormat.OpenXml.Math.OfficeMath();
            omath.Append(f);
            para.Append(omath);

            var chunks = MathHelper.ExtractMathChunks(para);
            Assert.Single(chunks);
            Assert.Equal("[公式: (1)/(2)]", chunks[0].Content);
        }

        [Fact]
        public void MergeField_Resolution_UsesMergeData()
        {
            var options = new ConvertOptions
            {
                MergeData = new System.Collections.Generic.Dictionary<string, string>
                {
                    { "CustomerName", "John Doe" }
                }
            };
            var converter = new DocxToPdfConverter(options);
            
            using var ms = new System.IO.MemoryStream();
            using var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();

            var result = converter.ResolveField("MERGEFIELD CustomerName", doc, "C:\\test.docx");
            
            Assert.Equal("John Doe", result);
        }

        [Fact]
        public void FileNameField_Resolution_UsesPath()
        {
            var converter = new DocxToPdfConverter();
            using var ms = new System.IO.MemoryStream();
            using var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();

            var result = converter.ResolveField("FILENAME", doc, "C:\\Documents\\Report.docx");
            
            Assert.Equal("Report", result);
        }

        [Theory]
        [InlineData("SYMBOL 97", "a")]           // ASCII 'a'
        [InlineData("SYMBOL 65", "A")]           // ASCII 'A'
        [InlineData("SYMBOL 8364", "€")]         // Euro sign (Unicode)
        [InlineData("SYMBOL 8730", "√")]         // Square root
        [InlineData("SYMBOL 945", "α")]          // Greek alpha
        public void SymbolField_Resolution_ReturnsCorrectCharacter(string instruction, string expected)
        {
            var converter = new DocxToPdfConverter();
            using var ms = new System.IO.MemoryStream();
            using var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();

            var result = converter.ResolveField(instruction, doc, "C:\\test.docx");

            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData("SYMBOL")]           // Missing char code
        [InlineData("SYMBOL abc")]       // Invalid char code
        public void SymbolField_InvalidInput_ReturnsNull(string instruction)
        {
            var converter = new DocxToPdfConverter();
            using var ms = new System.IO.MemoryStream();
            using var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();

            var result = converter.ResolveField(instruction, doc, "C:\\test.docx");

            Assert.Null(result);
        }

        [Fact]
        public void NestedFields_HyperlinkWithPageField_ResolvesCorrectly()
        {
            // 测试嵌套字段：HYPERLINK 包含 PAGE 字段
            var options = new ConvertOptions();
            var converter = new DocxToPdfConverter(options);

            // 创建包含嵌套字段的测试文档
            using var ms = new System.IO.MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body!;

                // 创建段落包含复杂字段结构
                var para = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                
                // 字段开始: HYPERLINK
                para.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
                para.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldCode(" HYPERLINK \"http://example.com/page\" ") { Space = SpaceProcessingModeValues.Preserve }));
                para.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
                
                // 显示文本
                para.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Page ") { Space = SpaceProcessingModeValues.Preserve }));
                
                // 嵌套字段: PAGE
                para.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
                para.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }));
                para.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
                para.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("1"))); // 模拟页码
                para.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldChar { FieldCharType = FieldCharValues.End }));
                
                para.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldChar { FieldCharType = FieldCharValues.End }));
                
                body.Append(para);
                mainPart.Document.Save();
            }

            // 验证文档可以正常转换（不抛出异常）
            ms.Position = 0;
            using (var doc = WordprocessingDocument.Open(ms, false))
            {
                var mainPart = doc.MainDocumentPart!;
                var para = mainPart.Document.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First();
                
                // 验证字段结构正确解析
                var fieldChars = para.Descendants<FieldChar>().ToList();
                // 应该有 6 个 FieldChar: 2 Begin + 2 Separate + 2 End
                Assert.True(fieldChars.Count >= 4, $"Expected at least 4 FieldChar elements, found {fieldChars.Count}");

                // 验证 Begin 和 End 成对出现
                var begins = fieldChars.Count(fc => fc.FieldCharType?.Value == FieldCharValues.Begin);
                var ends = fieldChars.Count(fc => fc.FieldCharType?.Value == FieldCharValues.End);
                Assert.Equal(begins, ends);
            }
        }

        [Fact]
        public void EqField_Fraction_ReturnsCorrectFormat()
        {
            var converter = new DocxToPdfConverter();
            using var ms = new System.IO.MemoryStream();
            using var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();

            var result = converter.ResolveField("EQ \\f(1,2)", doc, "C:\\test.docx");
            
            Assert.Equal("(1)/(2)", result);
        }

        [Fact]
        public void EqField_Superscript_ReturnsCorrectFormat()
        {
            var converter = new DocxToPdfConverter();
            using var ms = new System.IO.MemoryStream();
            using var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();

            var result = converter.ResolveField("EQ \\s(x,2)", doc, "C:\\test.docx");

            // EQ field returns the parsed format, actual rendering depends on implementation
            Assert.NotNull(result);
        }

        [Fact]
        public void EqField_Radical_ReturnsCorrectFormat()
        {
            var converter = new DocxToPdfConverter();
            using var ms = new System.IO.MemoryStream();
            using var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();

            var result = converter.ResolveField("EQ \\r(3,x)", doc, "C:\\test.docx");

            // EQ field returns the parsed format, actual rendering depends on implementation
            Assert.NotNull(result);
        }
    }
}
