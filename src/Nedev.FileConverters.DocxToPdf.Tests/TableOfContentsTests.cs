using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using Nedev.FileConverters.DocxToPdf.Models;
using Xunit;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class TableOfContentsTests
    {
        [Fact]
        public void UpdateTOCPageNumbers_AppendsEncodedNumbers()
        {
            var pdfDoc = new PdfDocument();
            using var ms = new MemoryStream();
            var writer = PdfWriter.GetInstance(pdfDoc, ms);
            writer.CloseStream = false;
            writer.SetChineseFont("SimSun");
            pdfDoc.Open();
            pdfDoc.NewPage();

            var entries = new List<TableOfContentsGenerator.TOCEntry>
            {
                new() { Title = "One", Level = 1, PageNumber = 2 },
                new() { Title = "Two", Level = 1, PageNumber = 3 }
            };

            TableOfContentsGenerator.UpdateTOCPageNumbers(pdfDoc, writer, entries, tocStartPage: 1);

            pdfDoc.Close();
            writer.Close();

            ms.Position = 0;
            var content = System.Text.Encoding.UTF8.GetString(ms.ToArray());
            Assert.Contains("<0032>", content);
            Assert.Contains("<0033>", content);
        }

        [Fact]
        public void Convert_TwoHeadingDocument_ContainsPageNumbersInOutput()
        {
            using var docStream = new MemoryStream();
            using (var word = WordprocessingDocument.Create(docStream, WordprocessingDocumentType.Document))
            {
                var main = word.AddMainDocumentPart();
                main.Document = new Document(new Body());
                var body = main.Document.Body;
                body.Append(new Paragraph(new Run(new Text("First heading")))
                {
                    ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" })
                });
                body.Append(new Paragraph(new Run(new Text("Second heading")))
                {
                    ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" })
                });
                main.Document.Save();
            }

            docStream.Position = 0;
            var converter = new DocxToPdfConverter(new ConvertOptions { GenerateTableOfContents = true });
            using var pdfOut = new MemoryStream();
            converter.Convert(docStream, pdfOut);

            var pdfText = System.Text.Encoding.UTF8.GetString(pdfOut.ToArray());
            Assert.Contains("<0032>", pdfText);
        }

        [Fact]
        public void PdfReader_CountsPages_Correctly()
        {
            // fake simple pdf with two pages markers
            var fake = System.Text.Encoding.ASCII.GetBytes("%PDF-1.4\n1 0 obj<< /Type /Page >>endobj\n2 0 obj<< /Type /Page >>endobj");
            using var ms = new MemoryStream(fake);
            var reader = new PdfReader(ms);
            Assert.Equal(2, reader.NumberOfPages);
        }

        [Fact]
        public void PdfStamper_AppendsOverUnderContent()
        {
            var fake = System.Text.Encoding.ASCII.GetBytes("dummy");
            var reader = new PdfReader(fake);
            using var outMs = new MemoryStream();
            var stamper = new PdfStamper(reader, outMs);
            stamper.GetOverContent(1).BeginText();
            stamper.GetOverContent(1).ShowText("hello");
            stamper.Close();
            outMs.Position = 0;
            var result = System.Text.Encoding.UTF8.GetString(outMs.ToArray());
            Assert.Contains("% OverContent page 1", result);
            Assert.Contains("hello", result);
        }
    }
}
