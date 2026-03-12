using System;
using System.Linq;
using Xunit;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;
using Nedev.FileConverters.DocxToPdf.Converters;
using Nedev.FileConverters.DocxToPdf.Models;
using Nedev.FileConverters.DocxToPdf.PdfEngine;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class TableConverterTests
    {
        private static TableConverter CreateConverter()
        {
            var fontHelper = new Helpers.FontHelper(new ConvertOptions());
            var paragraphConverter = new ParagraphConverter(fontHelper);
            return new TableConverter(fontHelper, paragraphConverter);
        }

        [Fact]
        public void EmptyTable_ReturnsNull()
        {
            // arrange
            var table = new WTable();
            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.Null(pdfTable);
        }

        [Fact]
        public void TableWithColors_AppliesBackground()
        {
            // arrange
            var tc = new WTableCell(new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
                new DocumentFormat.OpenXml.Wordprocessing.Shading { Fill = "FF0000" }
            ));
            var tr = new WTableRow(tc);
            var table = new WTable(new DocumentFormat.OpenXml.Wordprocessing.TableGrid(new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "500" }), tr);
            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.Single(pdfTable.RowsList);
            
            var cells = pdfTable.RowsList[0].Cells;
            Assert.Single(cells);
            
            var firstCell = cells[0];
            Assert.NotNull(firstCell);
            Assert.NotNull(firstCell.BackgroundColor);
            Assert.Equal(255, firstCell.BackgroundColor.R);
            Assert.Equal(0, firstCell.BackgroundColor.G);
            Assert.Equal(0, firstCell.BackgroundColor.B);
        }

        [Fact]
        public void StandardTable_ReturnsPopulatedPdfPTable()
        {
            // arrange
            var grid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "2000" },
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "2000" }
            );
            
            var row1 = new WTableRow(
                new WTableCell(new WParagraph(new WRun(new WText("R1C1")))),
                new WTableCell(new WParagraph(new WRun(new WText("R1C2"))))
            );
            var row2 = new WTableRow(
                new WTableCell(new WParagraph(new WRun(new WText("R2C1")))),
                new WTableCell(new WParagraph(new WRun(new WText("R2C2"))))
            );
            
            var table = new WTable(grid, row1, row2);
            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.Equal(2, pdfTable.RowsList.Count);
            Assert.Equal(2, pdfTable.RowsList[0].Cells.Count);
        }

        [Fact]
        public void TableWithMergedCells_HandlesGridSpan()
        {
            // arrange
            var grid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "1000" },
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "1000" },
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "1000" }
            );

            // 第一行：合并前两个单元格
            var cellProps = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
                new DocumentFormat.OpenXml.Wordprocessing.GridSpan { Val = 2 }
            );
            var row1 = new WTableRow(
                new WTableCell(cellProps, new WParagraph(new WRun(new WText("Merged")))),
                new WTableCell(new WParagraph(new WRun(new WText("Single"))))
            );

            // 第二行：三个独立单元格
            var row2 = new WTableRow(
                new WTableCell(new WParagraph(new WRun(new WText("A")))),
                new WTableCell(new WParagraph(new WRun(new WText("B")))),
                new WTableCell(new WParagraph(new WRun(new WText("C"))))
            );

            var table = new WTable(grid, row1, row2);
            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.Equal(2, pdfTable.RowsList.Count);
        }

        [Fact]
        public void TableWithBorders_AppliesBorderStyles()
        {
            // arrange
            var tableProps = new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                new DocumentFormat.OpenXml.Wordprocessing.TableBorders(
                    new DocumentFormat.OpenXml.Wordprocessing.TopBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 12, Color = "000000" },
                    new DocumentFormat.OpenXml.Wordprocessing.BottomBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 12, Color = "000000" },
                    new DocumentFormat.OpenXml.Wordprocessing.LeftBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 12, Color = "000000" },
                    new DocumentFormat.OpenXml.Wordprocessing.RightBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 12, Color = "000000" }
                )
            );

            var grid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "2000" }
            );

            var row = new WTableRow(
                new WTableCell(new WParagraph(new WRun(new WText("Cell with borders"))))
            );

            var table = new WTable(tableProps, grid, row);
            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.Single(pdfTable.RowsList);
        }

        [Fact]
        public void TableWithWidthPercentage_SetsWidthPercentage()
        {
            // arrange
            var tableProps = new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                new DocumentFormat.OpenXml.Wordprocessing.TableWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Pct, Width = "5000" } // 50%
            );

            var grid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "2000" }
            );

            var row = new WTableRow(
                new WTableCell(new WParagraph(new WRun(new WText("Width test"))))
            );

            var table = new WTable(tableProps, grid, row);
            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.True(pdfTable.WidthPercentage > 0);
        }

        [Fact]
        public void TableWithCellMargins_AppliesMargins()
        {
            // arrange
            var cellProps = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
                new DocumentFormat.OpenXml.Wordprocessing.TableCellMargin(
                    new DocumentFormat.OpenXml.Wordprocessing.TopMargin { Width = "100", Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa },
                    new DocumentFormat.OpenXml.Wordprocessing.BottomMargin { Width = "100", Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa },
                    new DocumentFormat.OpenXml.Wordprocessing.LeftMargin { Width = "100", Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa },
                    new DocumentFormat.OpenXml.Wordprocessing.RightMargin { Width = "100", Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa }
                )
            );

            var grid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "2000" }
            );

            var row = new WTableRow(
                new WTableCell(cellProps, new WParagraph(new WRun(new WText("Cell with margins"))))
            );

            var table = new WTable(grid, row);
            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.Single(pdfTable.RowsList);
        }

        [Fact]
        public void TableWithVerticalMerge_HandlesMerge()
        {
            // arrange
            var grid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "1000" },
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "1000" }
            );

            // 第一行：第一个单元格开始纵向合并
            var cell1Props = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
                new DocumentFormat.OpenXml.Wordprocessing.VerticalMerge { Val = DocumentFormat.OpenXml.Wordprocessing.MergedCellValues.Restart }
            );
            var row1 = new WTableRow(
                new WTableCell(cell1Props, new WParagraph(new WRun(new WText("Merged Start")))),
                new WTableCell(new WParagraph(new WRun(new WText("Normal"))))
            );

            // 第二行：第一个单元格继续纵向合并
            var cell2Props = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
                new DocumentFormat.OpenXml.Wordprocessing.VerticalMerge()
            );
            var row2 = new WTableRow(
                new WTableCell(cell2Props, new WParagraph()),
                new WTableCell(new WParagraph(new WRun(new WText("Normal"))))
            );

            var table = new WTable(grid, row1, row2);
            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.Equal(2, pdfTable.RowsList.Count);
        }

        [Fact]
        public void TableWithNestedContent_ConvertsCorrectly()
        {
            // arrange
            var grid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "3000" }
            );

            // 包含多个段落的单元格
            var cell = new WTableCell(
                new WParagraph(new WRun(new WText("Paragraph 1"))),
                new WParagraph(new WRun(new WText("Paragraph 2")))
            );

            var row = new WTableRow(cell);
            var table = new WTable(grid, row);
            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.Single(pdfTable.RowsList);
        }

        [Fact]
        public void TableWithAlignment_AppliesAlignment()
        {
            // arrange
            var tableProps = new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                new DocumentFormat.OpenXml.Wordprocessing.TableJustification { Val = DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues.Center }
            );

            var grid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "2000" }
            );

            var row = new WTableRow(
                new WTableCell(new WParagraph(new WRun(new WText("Centered table"))))
            );

            var table = new WTable(tableProps, grid, row);
            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
        }

        [Fact]
        public void LargeTable_HandlesManyRows()
        {
            // arrange
            var grid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "1000" }
            );

            var table = new WTable(grid);
            
            // 添加 10 行
            for (int i = 0; i < 10; i++)
            {
                table.Append(new WTableRow(
                    new WTableCell(new WParagraph(new WRun(new WText($"Row {i + 1}"))))
                ));
            }

            var converter = CreateConverter();

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.Equal(10, pdfTable.RowsList.Count);
        }
    }
}
