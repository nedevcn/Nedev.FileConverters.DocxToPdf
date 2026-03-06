using System;
using System.Linq;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Converters;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using Nedev.FileConverters.DocxToPdf.Models;
using Xunit;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class ParagraphConverterTests
    {
        [Fact]
        public void ComplexHyperlinkField_GeneratesAnchor()
        {
            // arrange: paragraph containing a complex hyperlink field
            // { HYPERLINK "http://example.com" }Click here
            var paragraph = new WParagraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" HYPERLINK \"http://example.com\"")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("Click here")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })
            );

            // debug: inspect paragraph composition
            foreach (var elem in paragraph.ChildElements)
            {
                Console.WriteLine($"Child: {elem.GetType().Name} - {elem.InnerXml}");
                if (elem is Run r)
                {
                    foreach (var nested in r.ChildElements)
                        Console.WriteLine($"  Nested: {nested.GetType().Name} - {nested.InnerXml}");
                }
            }

            var converter = new ParagraphConverter(new Helpers.FontHelper(new Models.ConvertOptions()))
            {
                // provide a resolver that simply returns the URL for hyperlink fields
                FieldResolver = instr =>
                {
                    if (instr.TrimStart().ToUpperInvariant().StartsWith("HYPERLINK"))
                    {
                        var m = System.Text.RegularExpressions.Regex.Match(instr, "HYPERLINK\\s+\"([^\"]+)\"");
                        return m.Success ? m.Groups[1].Value : null;
                    }
                    return null;
                }
            };

            // act
            // clear any previous log
            ParagraphConverter.DebugLog.Clear();
            var elements = converter.Convert(paragraph);

            // assert
            var pdfPara = Assert.Single(elements) as PdfEngine.Paragraph;
            Assert.NotNull(pdfPara);
            var chunks = pdfPara.Chunks.OfType<Chunk>().ToList();
            Assert.NotEmpty(chunks);

            // find any chunk with anchor
            var anchorChunk = chunks.FirstOrDefault(c => !string.IsNullOrEmpty(c.Anchor));
            if (anchorChunk == null)
            {
                // fail with diagnostic information
                var details = string.Join(" | ", chunks.Select(c => $"[{c.Content}]('{c.Anchor}')"));
                var debugLog = ParagraphConverter.DebugLog.ToString();
                throw new Xunit.Sdk.XunitException("No anchored chunk found; chunks: " + details + "\nLog:\n" + debugLog);
            }
            var log = ParagraphConverter.DebugLog.ToString();
            Assert.True(anchorChunk.Anchor == "http://example.com",
                $"expected anchor 'http://example.com' but was '{anchorChunk.Anchor}'\nLog:\n{log}");
            Assert.Contains("Click", anchorChunk.Content);
        }
    }
}
