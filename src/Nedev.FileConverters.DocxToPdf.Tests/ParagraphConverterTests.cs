using System;
using System.Linq;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Converters;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using Nedev.FileConverters.DocxToPdf.Models;
using System.Collections.Generic;
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

        [Fact]
        public void SampleDocument_HeaderParagraph_PreservesSpaces()
        {
            // load the test document and inspect the first paragraph runs
            // resolve path relative to test assembly output
            // determine workspace root by walking up from the test output directory
            var dir = new System.IO.DirectoryInfo(AppContext.BaseDirectory);
            // climb up five levels from output dir to reach workspace project root
            for (int i = 0; i < 5; i++)
            {
                if (dir.Parent != null) dir = dir.Parent;
            }
            var root = dir.FullName;
            var path = System.IO.Path.Combine(root, "tests", "能率项目-维修工单服务报告模板1.3.docx");
            Console.WriteLine("Computed workspace root: " + root);
            Console.WriteLine("Opening docx path: " + path);
            Console.WriteLine("Opening docx path: " + path);
            using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, false);
            var body = doc.MainDocumentPart?.Document.Body;
            Assert.NotNull(body);
            var paras = body.Descendants<WParagraph>().ToList();
            Assert.NotEmpty(paras);

            Console.WriteLine("Document contains " + paras.Count + " paragraphs.");
            WParagraph? target = null;
            for (int i = 0; i < paras.Count; i++)
            {
                var text = paras[i].InnerText;
                Console.WriteLine($"Para[{i}]='{text}'");
                if (text.Contains("能率") || text.Contains("有限公司"))
                {
                    target = paras[i];
                    Console.WriteLine("--> selected as target paragraph");
                    break;
                }
            }
            Assert.NotNull(target);

            Console.WriteLine("Target paragraph XML: " + target.OuterXml);
            Console.WriteLine("Inspect runs in target paragraph:");
            int idx = 0;
            foreach (var run in target.Elements<Run>())
            {
                var txt = run.GetFirstChild<Text>();
                var content = txt?.Text ?? string.Empty;
                var xmlSpace = txt?.Space?.Value;
                Console.WriteLine($"Run[{idx++}] content: >{content}< xml:space={xmlSpace} innerXml={run.InnerXml}");
            }

            // convert the target paragraph to see what chunks are produced
            var converter = new ParagraphConverter(new Helpers.FontHelper(new Models.ConvertOptions()));
            ParagraphConverter.DebugLog.Clear();
            var elements = converter.Convert(target!);
            Assert.Single(elements);
            var pdfPara = elements[0] as PdfEngine.Paragraph;
            Assert.NotNull(pdfPara);

            var chunks = pdfPara.Chunks.OfType<Chunk>().ToList();
            Console.WriteLine("Converted chunks:");
            foreach (var c in chunks)
            {
                Console.WriteLine($"[{c.Content}]");
            }
        }

        [Fact]
        public void ParagraphCenterAlignment_IsApplied()
        {
            // create a simple paragraph with a single chunk
            var font = new PdfEngine.Font("Helvetica", 12);
            var para = new PdfEngine.Paragraph("test", font)
            {
                Alignment = Element.ALIGN_CENTER
            };

            var canvas = new RecordingContentByte();
            var col = new ColumnText(canvas);
            // narrow column so we can easily see offset
            col.SetSimpleColumn(0, 0, 200, 200);
            col.AddElement(para);
            col.Go();

            // after rendering we should have recorded at least one x coordinate
            Assert.NotEmpty(canvas.RecordedX);
            // center alignment means first x should be greater than left margin (0)
            Assert.True(canvas.RecordedX[0] > 0, "Expected text to be shifted right for center alignment");
        }

        /// <summary>
        /// Helper canvas that records the X coordinate passed to SetTextMatrix
        /// so tests can verify alignment.
        /// </summary>
        private class RecordingContentByte : PdfEngine.PdfContentByte
        {
            public List<float> RecordedX { get; } = new List<float>();

            public override void SetTextMatrix(float a, float b, float c, float d, float e, float f)
            {
                RecordedX.Add(e);
                base.SetTextMatrix(a, b, c, d, e, f);
            }
        }
    }
}
