using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf;
using System.IO;
using Xunit;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class DocVariableFieldTests
    {
        [Fact]
        public void ResolveField_DocVariable_ReturnsVariableValue()
        {
            // 创建测试文档
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // 添加 DocumentSettingsPart 并设置变量
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(
                    new DocumentVariables(
                        new DocumentVariable { Name = "CompanyName", Val = "Nedev Corporation" },
                        new DocumentVariable { Name = "DocumentVersion", Val = "1.0" }
                    )
                );

                mainPart.Document.Save();

                // 测试 DOCVARIABLE 字段解析
                var converter = new DocxToPdfConverter();
                var result = converter.ResolveField("DOCVARIABLE CompanyName", doc, "C:\\test.docx");

                Assert.Equal("Nedev Corporation", result);
            }
        }

        [Fact]
        public void ResolveField_DocVariable_NotFound_ReturnsPlaceholder()
        {
            // 创建测试文档
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // 添加空的 DocumentSettingsPart
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();

                mainPart.Document.Save();

                // 测试不存在的变量
                var converter = new DocxToPdfConverter();
                var result = converter.ResolveField("DOCVARIABLE NonExistent", doc, "C:\\test.docx");

                Assert.Equal("«NonExistent»", result);
            }
        }

        [Fact]
        public void ResolveField_DocVariable_EmptyInstruction_ReturnsNull()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                var result = converter.ResolveField("DOCVARIABLE", doc, "C:\\test.docx");

                Assert.Null(result);
            }
        }

        [Fact]
        public void ResolveField_DocVariable_CaseInsensitive()
        {
            // 创建测试文档
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(
                    new DocumentVariables(
                        new DocumentVariable { Name = "TestVar", Val = "Test Value" }
                    )
                );

                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                
                // 测试大小写不敏感
                var result1 = converter.ResolveField("DOCVARIABLE testvar", doc, "C:\\test.docx");
                var result2 = converter.ResolveField("DOCVARIABLE TESTVAR", doc, "C:\\test.docx");
                var result3 = converter.ResolveField("DOCVARIABLE TestVar", doc, "C:\\test.docx");

                Assert.Equal("Test Value", result1);
                Assert.Equal("Test Value", result2);
                Assert.Equal("Test Value", result3);
            }
        }

        [Fact]
        public void ResolveField_DocVariable_MultipleVariables()
        {
            // 创建测试文档
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(
                    new DocumentVariables(
                        new DocumentVariable { Name = "Var1", Val = "Value1" },
                        new DocumentVariable { Name = "Var2", Val = "Value2" },
                        new DocumentVariable { Name = "Var3", Val = "Value3" }
                    )
                );

                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();

                Assert.Equal("Value1", converter.ResolveField("DOCVARIABLE Var1", doc, "C:\\test.docx"));
                Assert.Equal("Value2", converter.ResolveField("DOCVARIABLE Var2", doc, "C:\\test.docx"));
                Assert.Equal("Value3", converter.ResolveField("DOCVARIABLE Var3", doc, "C:\\test.docx"));
            }
        }

        [Fact]
        public void ResolveField_DocVariable_SpecialCharacters()
        {
            // 创建测试文档
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(
                    new DocumentVariables(
                        new DocumentVariable { Name = "Special", Val = "Hello, World! @#$%" }
                    )
                );

                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                var result = converter.ResolveField("DOCVARIABLE Special", doc, "C:\\test.docx");

                Assert.Equal("Hello, World! @#$%", result);
            }
        }

        [Fact]
        public void ResolveField_DocVariable_EmptyValue()
        {
            // 创建测试文档
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(
                    new DocumentVariables(
                        new DocumentVariable { Name = "EmptyVar", Val = "" }
                    )
                );

                mainPart.Document.Save();

                var converter = new DocxToPdfConverter();
                var result = converter.ResolveField("DOCVARIABLE EmptyVar", doc, "C:\\test.docx");

                Assert.Equal("", result);
            }
        }
    }
}
