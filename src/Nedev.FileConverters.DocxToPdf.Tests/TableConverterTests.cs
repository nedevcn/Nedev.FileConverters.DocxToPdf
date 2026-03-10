using System;
using System.Linq;
using Xunit;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using Nedev.FileConverters.DocxToPdf.Converters;
using Nedev.FileConverters.DocxToPdf.Models;
using Nedev.FileConverters.DocxToPdf.PdfEngine;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class TableConverterTests
    {
        [Fact]
        public void EmptyTable_ReturnsNull()
        {
            // arrange
            var table = new WTable();

            var fontHelper = new Helpers.FontHelper(new ConvertOptions());
            var paragraphConverter = new ParagraphConverter(fontHelper);
            var converter = new TableConverter(fontHelper, paragraphConverter);

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

            var fontHelper = new Helpers.FontHelper(new ConvertOptions());
            var paragraphConverter = new ParagraphConverter(fontHelper);
            var converter = new TableConverter(fontHelper, paragraphConverter);

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.Equal(1, pdfTable.RowsList.Count);
            
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
                new WTableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("R1C1")))),
                new WTableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("R1C2"))))
            );
            var row2 = new WTableRow(
                new WTableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("R2C1")))),
                new WTableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("R2C2"))))
            );
            
            var table = new WTable(grid, row1, row2);

            var fontHelper = new Helpers.FontHelper(new ConvertOptions());
            var paragraphConverter = new ParagraphConverter(fontHelper);
            var converter = new TableConverter(fontHelper, paragraphConverter);

            // act
            var pdfTable = converter.Convert(table, 500f);

            // assert
            Assert.NotNull(pdfTable);
            Assert.Equal(2, pdfTable.RowsList.Count);
            Assert.Equal(2, pdfTable.RowsList[0].Cells.Count);
        }
    }
}
