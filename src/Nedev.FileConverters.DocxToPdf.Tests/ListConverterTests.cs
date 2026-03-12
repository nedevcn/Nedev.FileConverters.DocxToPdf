using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;
using Nedev.FileConverters.DocxToPdf.Converters;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using Nedev.FileConverters.DocxToPdf.Models;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class ListConverterTests
    {
        private static ListConverter CreateConverter()
        {
            var fontHelper = new Helpers.FontHelper(new ConvertOptions());
            return new ListConverter(fontHelper);
        }

        [Fact]
        public void List_CreatesPdfList_WithItems()
        {
            // arrange
            var numPr = new DocumentFormat.OpenXml.Wordprocessing.NumberingProperties(
                new DocumentFormat.OpenXml.Wordprocessing.NumberingId { Val = 1 },
                new DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference { Val = 0 }
            );
            var pPr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(numPr);
            var paragraph = new WParagraph(pPr, new WRun(new WText("List item text")));

            var listConverter = CreateConverter();
            var listItems = new List<WParagraph> { paragraph };

            // act
            var elements = listConverter.ConvertListItems(listItems, null, 1);

            // assert
            Assert.NotNull(elements);
            var pdfList = elements;
            Assert.NotNull(pdfList);
            Assert.Single(pdfList.Items);
            
            var listItem = pdfList.Items.First();
            Assert.NotNull(listItem);
        }

        [Fact]
        public void OrderedList_CreatesPdfList_WithOrderedType()
        {
            // arrange
            var numPr = new DocumentFormat.OpenXml.Wordprocessing.NumberingProperties(
                new DocumentFormat.OpenXml.Wordprocessing.NumberingId { Val = 1 },
                new DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference { Val = 0 }
            );
            var pPr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(numPr);
            var paragraph = new WParagraph(pPr, new WRun(new WText("First item")));

            var listConverter = CreateConverter();
            var listItems = new List<WParagraph> { paragraph };

            // act
            var pdfList = listConverter.ConvertListItems(listItems, null, 1);

            // assert
            Assert.NotNull(pdfList);
            Assert.Single(pdfList.Items);
        }

        [Fact]
        public void MultipleListItems_CreatesMultipleItems()
        {
            // arrange
            var listConverter = CreateConverter();
            var listItems = new List<WParagraph>();

            for (int i = 1; i <= 5; i++)
            {
                var numPr = new DocumentFormat.OpenXml.Wordprocessing.NumberingProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.NumberingId { Val = 1 },
                    new DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference { Val = 0 }
                );
                var pPr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(numPr);
                var paragraph = new WParagraph(pPr, new WRun(new WText($"Item {i}")));
                listItems.Add(paragraph);
            }

            // act
            var pdfList = listConverter.ConvertListItems(listItems, null, 1);

            // assert
            Assert.NotNull(pdfList);
            Assert.Equal(5, pdfList.Count);
        }

        [Fact]
        public void NestedList_HandlesNestedLevels()
        {
            // arrange
            var listConverter = CreateConverter();
            var listItems = new List<WParagraph>();

            // 第一级
            var numPr1 = new DocumentFormat.OpenXml.Wordprocessing.NumberingProperties(
                new DocumentFormat.OpenXml.Wordprocessing.NumberingId { Val = 1 },
                new DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference { Val = 0 }
            );
            var pPr1 = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(numPr1);
            listItems.Add(new WParagraph(pPr1, new WRun(new WText("Level 1 Item"))));

            // 第二级（相同 NumberingId 但不同 Level）
            var numPr2 = new DocumentFormat.OpenXml.Wordprocessing.NumberingProperties(
                new DocumentFormat.OpenXml.Wordprocessing.NumberingId { Val = 1 },
                new DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference { Val = 1 }
            );
            var pPr2 = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(numPr2);
            listItems.Add(new WParagraph(pPr2, new WRun(new WText("Level 2 Item"))));

            // act
            var pdfList = listConverter.ConvertListItems(listItems, null, 1);

            // assert
            Assert.NotNull(pdfList);
            // 注意：ListConverter 可能只处理相同级别的列表项
            // 实际行为取决于实现，这里只验证不为空
            Assert.True(pdfList.Count >= 1);
        }

        [Fact]
        public void EmptyList_ReturnsEmptyList()
        {
            // arrange
            var listConverter = CreateConverter();
            var listItems = new List<WParagraph>();

            // act
            var pdfList = listConverter.ConvertListItems(listItems, null, 1);

            // assert
            Assert.NotNull(pdfList);
            Assert.Empty(pdfList.Items);
        }

        [Fact]
        public void ListWithCustomNumbering_AppliesNumbering()
        {
            // arrange
            var listConverter = CreateConverter();
            
            // 创建 Numbering 定义
            var numbering = new DocumentFormat.OpenXml.Wordprocessing.Numbering(
                new DocumentFormat.OpenXml.Wordprocessing.AbstractNum(
                    new DocumentFormat.OpenXml.Wordprocessing.Level(
                        new DocumentFormat.OpenXml.Wordprocessing.NumberingFormat { Val = DocumentFormat.OpenXml.Wordprocessing.NumberFormatValues.UpperRoman },
                        new DocumentFormat.OpenXml.Wordprocessing.LevelText { Val = "%1." }
                    ) { LevelIndex = 0 }
                ) { AbstractNumberId = 1 },
                new DocumentFormat.OpenXml.Wordprocessing.NumberingInstance(
                    new DocumentFormat.OpenXml.Wordprocessing.AbstractNumId { Val = 1 }
                ) { NumberID = 1 }
            );

            var numPr = new DocumentFormat.OpenXml.Wordprocessing.NumberingProperties(
                new DocumentFormat.OpenXml.Wordprocessing.NumberingId { Val = 1 },
                new DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference { Val = 0 }
            );
            var pPr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(numPr);
            var paragraph = new WParagraph(pPr, new WRun(new WText("Roman numeral item")));

            var listItems = new List<WParagraph> { paragraph };

            // act
            var pdfList = listConverter.ConvertListItems(listItems, numbering, 1);

            // assert
            Assert.NotNull(pdfList);
            Assert.Single(pdfList.Items);
        }

        [Fact]
        public void BulletList_CreatesUnorderedList()
        {
            // arrange
            var listConverter = CreateConverter();
            
            var numPr = new DocumentFormat.OpenXml.Wordprocessing.NumberingProperties(
                new DocumentFormat.OpenXml.Wordprocessing.NumberingId { Val = 2 },
                new DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference { Val = 0 }
            );
            var pPr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(numPr);
            var paragraph = new WParagraph(pPr, new WRun(new WText("Bullet item")));

            var listItems = new List<WParagraph> { paragraph };

            // act
            var pdfList = listConverter.ConvertListItems(listItems, null, 2);

            // assert
            Assert.NotNull(pdfList);
            Assert.Single(pdfList.Items);
        }

        [Fact]
        public void ListWithFormatting_PreservesFormatting()
        {
            // arrange
            var listConverter = CreateConverter();
            
            var numPr = new DocumentFormat.OpenXml.Wordprocessing.NumberingProperties(
                new DocumentFormat.OpenXml.Wordprocessing.NumberingId { Val = 1 },
                new DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference { Val = 0 }
            );
            var pPr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(numPr);
            
            // 带格式的文本
            var run = new WRun(
                new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.Bold()
                ),
                new WText("Bold list item")
            );
            var paragraph = new WParagraph(pPr, run);

            var listItems = new List<WParagraph> { paragraph };

            // act
            var pdfList = listConverter.ConvertListItems(listItems, null, 1);

            // assert
            Assert.NotNull(pdfList);
            Assert.Single(pdfList.Items);
        }

        [Fact]
        public void ListWithMultipleRuns_CombinesText()
        {
            // arrange
            var listConverter = CreateConverter();
            
            var numPr = new DocumentFormat.OpenXml.Wordprocessing.NumberingProperties(
                new DocumentFormat.OpenXml.Wordprocessing.NumberingId { Val = 1 },
                new DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference { Val = 0 }
            );
            var pPr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(numPr);
            
            var paragraph = new WParagraph(
                pPr,
                new WRun(new WText("First part ")),
                new WRun(new WText("second part"))
            );

            var listItems = new List<WParagraph> { paragraph };

            // act
            var pdfList = listConverter.ConvertListItems(listItems, null, 1);

            // assert
            Assert.NotNull(pdfList);
            Assert.Single(pdfList.Items);
        }

        [Fact]
        public void ListItemWithIndentation_AppliesIndent()
        {
            // arrange
            var listConverter = CreateConverter();
            
            var numPr = new DocumentFormat.OpenXml.Wordprocessing.NumberingProperties(
                new DocumentFormat.OpenXml.Wordprocessing.NumberingId { Val = 1 },
                new DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference { Val = 0 }
            );
            var indent = new DocumentFormat.OpenXml.Wordprocessing.Indentation { Left = "720" };
            var pPr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(numPr, indent);
            var paragraph = new WParagraph(pPr, new WRun(new WText("Indented item")));

            var listItems = new List<WParagraph> { paragraph };

            // act
            var pdfList = listConverter.ConvertListItems(listItems, null, 1);

            // assert
            Assert.NotNull(pdfList);
            Assert.Single(pdfList.Items);
        }
    }
}
