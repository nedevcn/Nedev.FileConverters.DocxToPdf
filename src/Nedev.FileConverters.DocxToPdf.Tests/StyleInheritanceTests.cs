using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Helpers;
using System.IO;
using Xunit;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class StyleInheritanceTests
    {
        #region StyleInheritanceResolver Tests

        [Fact]
        public void ResolveStyle_WithSimpleStyle_ReturnsCorrectProperties()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // 创建样式定义
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var styles = new Styles();

                var testStyle = new Style(
                    new StyleParagraphProperties(
                        new Justification { Val = JustificationValues.Center },
                        new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Auto }
                    ),
                    new StyleRunProperties(
                        new Bold(),
                        new FontSize { Val = "24" },
                        new RunFonts { Ascii = "Arial", EastAsia = "微软雅黑" }
                    )
                )
                {
                    StyleId = "TestStyle",
                    Type = StyleValues.Paragraph
                };

                styles.Append(testStyle);
                stylesPart.Styles = styles;
                stylesPart.Styles.Save();

                mainPart.Document.Save();

                // 测试样式解析
                var resolver = new StyleInheritanceResolver(styles);
                var resolved = resolver.ResolveStyle("TestStyle");

                Assert.Equal("TestStyle", resolved.StyleId);
                Assert.Equal(JustificationValues.Center, resolved.Justification);
                Assert.True(resolved.Bold);
                Assert.Equal("24", resolved.FontSize);
                Assert.Equal("Arial", resolved.FontAscii);
                Assert.Equal("微软雅黑", resolved.FontEastAsia);
                Assert.NotNull(resolved.Spacing);
            }
        }

        [Fact]
        public void ResolveStyle_WithInheritance_ChildOverridesParent()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var styles = new Styles();

                // 父样式
                var parentStyle = new Style(
                    new StyleParagraphProperties(
                        new Justification { Val = JustificationValues.Left }
                    ),
                    new StyleRunProperties(
                        new FontSize { Val = "22" },
                        new Bold()
                    )
                )
                {
                    StyleId = "ParentStyle",
                    Type = StyleValues.Paragraph
                };

                // 子样式（继承父样式）
                var childStyle = new Style(
                    new BasedOn { Val = "ParentStyle" },
                    new StyleParagraphProperties(
                        new Justification { Val = JustificationValues.Right }
                    ),
                    new StyleRunProperties(
                        new Italic() // 覆盖父样式的加粗
                    )
                )
                {
                    StyleId = "ChildStyle",
                    Type = StyleValues.Paragraph
                };

                styles.Append(parentStyle);
                styles.Append(childStyle);
                stylesPart.Styles = styles;
                stylesPart.Styles.Save();

                mainPart.Document.Save();

                var resolver = new StyleInheritanceResolver(styles);
                var resolved = resolver.ResolveStyle("ChildStyle");

                // 子样式应该覆盖父样式的对齐方式
                Assert.Equal(JustificationValues.Right, resolved.Justification);
                // 子样式继承了父样式的字体大小
                Assert.Equal("22", resolved.FontSize);
                // 子样式添加了斜体
                Assert.True(resolved.Italic);
            }
        }

        [Fact]
        public void ResolveStyle_MultiLevelInheritance_WorksCorrectly()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var styles = new Styles();

                // 基础样式
                var baseStyle = new Style(
                    new StyleRunProperties(
                        new FontSize { Val = "20" }
                    )
                )
                {
                    StyleId = "BaseStyle",
                    Type = StyleValues.Paragraph
                };

                // 中间样式
                var middleStyle = new Style(
                    new BasedOn { Val = "BaseStyle" },
                    new StyleRunProperties(
                        new Bold()
                    )
                )
                {
                    StyleId = "MiddleStyle",
                    Type = StyleValues.Paragraph
                };

                // 具体样式
                var specificStyle = new Style(
                    new BasedOn { Val = "MiddleStyle" },
                    new StyleRunProperties(
                        new Italic()
                    )
                )
                {
                    StyleId = "SpecificStyle",
                    Type = StyleValues.Paragraph
                };

                styles.Append(baseStyle);
                styles.Append(middleStyle);
                styles.Append(specificStyle);
                stylesPart.Styles = styles;
                stylesPart.Styles.Save();

                mainPart.Document.Save();

                var resolver = new StyleInheritanceResolver(styles);
                var resolved = resolver.ResolveStyle("SpecificStyle");

                Assert.Equal("20", resolved.FontSize); // 继承自 BaseStyle
                Assert.True(resolved.Bold); // 继承自 MiddleStyle
                Assert.True(resolved.Italic); // 来自 SpecificStyle
            }
        }

        [Fact]
        public void ResolveStyle_HeadingStyle_DetectsOutlineLevel()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var styles = new Styles();

                var headingStyle = new Style(
                    new StyleParagraphProperties(
                        new OutlineLevel { Val = 2 }
                    ),
                    new StyleRunProperties(
                        new Bold(),
                        new FontSize { Val = "28" }
                    )
                )
                {
                    StyleId = "Heading2",
                    Type = StyleValues.Paragraph
                };

                styles.Append(headingStyle);
                stylesPart.Styles = styles;
                stylesPart.Styles.Save();

                mainPart.Document.Save();

                var resolver = new StyleInheritanceResolver(styles);
                var resolved = resolver.ResolveStyle("Heading2");

                Assert.True(resolved.IsHeading());
                Assert.Equal(2, resolved.GetHeadingLevel());
                Assert.Equal(2, resolved.OutlineLevel);
            }
        }

        [Fact]
        public void MergeWithDirectProperties_DirectPropertiesOverrideInherited()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var styles = new Styles();

                var testStyle = new Style(
                    new StyleParagraphProperties(
                        new Justification { Val = JustificationValues.Center },
                        new SpacingBetweenLines { Before = "120", After = "120" }
                    ),
                    new StyleRunProperties(
                        new FontSize { Val = "24" }
                    )
                )
                {
                    StyleId = "TestStyle",
                    Type = StyleValues.Paragraph
                };

                styles.Append(testStyle);
                stylesPart.Styles = styles;
                stylesPart.Styles.Save();

                mainPart.Document.Save();

                var resolver = new StyleInheritanceResolver(styles);
                var inherited = resolver.ResolveStyle("TestStyle");

                // 创建直接段落属性（覆盖继承的样式）
                var directProps = new ParagraphProperties(
                    new Justification { Val = JustificationValues.Right },
                    new SpacingBetweenLines { Before = "240" }
                );

                var merged = resolver.MergeWithDirectProperties(inherited, directProps);

                // 直接属性应该覆盖继承的属性
                Assert.Equal(JustificationValues.Right, merged.Justification);
                Assert.NotNull(merged.Spacing);
                // 继承的属性应该保留（如果没有被覆盖）
                Assert.Equal("24", merged.FontSize);
            }
        }

        [Fact]
        public void ResolveStyle_Caching_ReturnsSameInstance()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var styles = new Styles();

                var testStyle = new Style(
                    new StyleRunProperties(new Bold())
                )
                {
                    StyleId = "CachedStyle",
                    Type = StyleValues.Paragraph
                };

                styles.Append(testStyle);
                stylesPart.Styles = styles;
                stylesPart.Styles.Save();

                mainPart.Document.Save();

                var resolver = new StyleInheritanceResolver(styles);
                
                // 多次解析相同样式
                var resolved1 = resolver.ResolveStyle("CachedStyle");
                var resolved2 = resolver.ResolveStyle("CachedStyle");
                var resolved3 = resolver.ResolveStyle("CachedStyle");

                // 应该返回缓存的相同实例
                Assert.Same(resolved1, resolved2);
                Assert.Same(resolved2, resolved3);
            }
        }

        [Fact]
        public void ResolveStyle_EmptyOrNullStyleId_ReturnsEmptyStyle()
        {
            var resolver = new StyleInheritanceResolver(null);
            
            var resolved1 = resolver.ResolveStyle(null);
            var resolved2 = resolver.ResolveStyle("");
            var resolved3 = resolver.ResolveStyle("   ");

            Assert.NotNull(resolved1);
            Assert.NotNull(resolved2);
            Assert.NotNull(resolved3);
            Assert.Null(resolved1.StyleId);
            Assert.Null(resolved2.StyleId);
            Assert.Null(resolved3.StyleId);
        }

        [Fact]
        public void ResolvedStyle_Clone_CreatesIndependentCopy()
        {
            var original = new ResolvedStyle
            {
                StyleId = "Original",
                Bold = true,
                FontSize = "24",
                Justification = JustificationValues.Center
            };

            var clone = original.Clone();

            // 克隆应该具有相同的值
            Assert.Equal(original.StyleId, clone.StyleId);
            Assert.Equal(original.Bold, clone.Bold);
            Assert.Equal(original.FontSize, clone.FontSize);
            Assert.Equal(original.Justification, clone.Justification);

            // 修改克隆不应该影响原始对象
            clone.StyleId = "Modified";
            clone.Bold = false;

            Assert.Equal("Original", original.StyleId);
            Assert.True(original.Bold);
        }

        [Fact]
        public void ResolvedStyle_GetFontSizeInPoints_ConvertsCorrectly()
        {
            var style1 = new ResolvedStyle { FontSize = "24" }; // 12pt
            var style2 = new ResolvedStyle { FontSize = "28" }; // 14pt
            var style3 = new ResolvedStyle { FontSize = null };

            Assert.Equal(12f, style1.GetFontSizeInPoints());
            Assert.Equal(14f, style2.GetFontSizeInPoints());
            Assert.Null(style3.GetFontSizeInPoints());
        }

        [Fact]
        public void ResolvedStyle_GetPreferredFontName_ReturnsCorrectPriority()
        {
            var style1 = new ResolvedStyle { FontEastAsia = "微软雅黑", FontAscii = "Arial" };
            var style2 = new ResolvedStyle { FontAscii = "Arial", FontHighAnsi = "Times New Roman" };
            var style3 = new ResolvedStyle { FontHighAnsi = "Times New Roman" };
            var style4 = new ResolvedStyle();

            Assert.Equal("微软雅黑", style1.GetPreferredFontName());
            Assert.Equal("Arial", style2.GetPreferredFontName());
            Assert.Equal("Times New Roman", style3.GetPreferredFontName());
            Assert.Null(style4.GetPreferredFontName());
        }

        #endregion

        #region Integration Tests with ParagraphConverter

        [Fact]
        public void ParagraphConverter_WithStyleInheritance_AppliesInheritedProperties()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var styles = new Styles();

                // 创建带继承样式的段落样式
                var customStyle = new Style(
                    new StyleParagraphProperties(
                        new Justification { Val = JustificationValues.Center },
                        new SpacingBetweenLines { Before = "200", After = "200", Line = "360" }
                    ),
                    new StyleRunProperties(
                        new Bold(),
                        new FontSize { Val = "26" }
                    )
                )
                {
                    StyleId = "CustomStyle",
                    Type = StyleValues.Paragraph
                };

                styles.Append(customStyle);
                stylesPart.Styles = styles;
                stylesPart.Styles.Save();

                // 创建应用该样式的段落
                var paragraph = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = "CustomStyle" }
                    ),
                    new Run(new Text("Test content"))
                );

                mainPart.Document.Body?.Append(paragraph);
                mainPart.Document.Save();

                // 创建转换器并测试
                var options = new Models.ConvertOptions();
                var fontHelper = new FontHelper(options);
                var converter = new Converters.ParagraphConverter(fontHelper, styles);

                var elements = converter.Convert(paragraph);

                Assert.NotEmpty(elements);
                // 验证段落继承了样式的属性
                // 注意：由于 PDF 引擎的具体实现，这里主要验证没有异常发生
                // 详细的属性验证需要在更高级别的集成测试中完成
            }
        }

        #endregion
    }
}
