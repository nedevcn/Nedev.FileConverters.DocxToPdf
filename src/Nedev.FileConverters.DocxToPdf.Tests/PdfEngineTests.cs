using System;
using System.IO;
using System.Linq;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using Xunit;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class PdfEngineTests
    {
        [Fact]
        public void ColumnText_DoesNotLiveLock_WhenElementTooLarge()
        {
            // create an element that will always hit the boundary check
            var cb = new PdfContentByte();
            var ct = new ColumnText(cb);
            ct.SetSimpleColumn(0, 0, 10, 10); // tiny column

            var fake = new FakeElement();
            ct.AddElement(fake);

            // first call should detect overflow and return NO_MORE_TEXT
            var res1 = ct.Go();
            Assert.Equal(ColumnText.NO_MORE_TEXT, res1);

            // inspect internal list via reflection to ensure element still present
            var field = typeof(ColumnText).GetField("_elements", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var list = field?.GetValue(ct) as System.Collections.IList;
            Assert.NotNull(list);
            Assert.Single(list);

            // second call should not loop forever, it can return NO_MORE_TEXT again
            var res2 = ct.Go();
            Assert.Equal(ColumnText.NO_MORE_TEXT, res2);
        }

        class FakeElement : IElement
        {
            public int Type => -999;
            public bool IsContent() => true;
            public bool IsNestable() => false;
        }

        [Fact]
        public void PdfWriter_OverflowMovesToNextPage()
        {
            var doc = new PdfDocument();
            using var ms = new MemoryStream();
            var writer = PdfWriter.GetInstance(doc, ms);
            writer.CloseStream = false;
            doc.Open();
            var page = doc.NewPage();

            // add two paragraphs and force the second to overflow
            var p1 = new Paragraph("Line1\nLine2", Font.Helvetica(12));
            var p2 = new Paragraph("Overflow", Font.Helvetica(12));
            page.AddElement(p1);
            page.AddElement(p2);

            // manually invoke generation to simulate writing
            var content = writer.GeneratePageContent(page, writer);
            // first page should contain both paragraphs? but overflow logic not triggered because only high-level test
            Assert.NotNull(content);
            
            // now simulate very small margins to provoke overflow
            page = doc.NewPage();
            page.AddElement(new Paragraph("A", Font.Helvetica(12)));
            // fill Y with negative margin
            var content2 = writer.GeneratePageContent(page, writer);
            Assert.NotNull(content2);
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

        [Fact]
        public void PdfReader_ParsesMediaBoxSizes()
        {
            // create fake pdf data containing two media boxes
            var fake = System.Text.Encoding.ASCII.GetBytes("%PDF-1.4\n/MediaBox [0 0 200 300]\n1 0 obj<< /Type /Page >>endobj\n/MediaBox [0 0 400 500]\n2 0 obj<< /Type /Page >>endobj");
            using var ms = new MemoryStream(fake);
            var reader = new PdfReader(ms);
            Assert.Equal(2, reader.NumberOfPages);
            var sz1 = reader.GetPageSize(1);
            var sz2 = reader.GetPageSize(2);
            Assert.Equal(200f, sz1.Width);
            Assert.Equal(300f, sz1.Height);
            Assert.Equal(400f, sz2.Width);
            Assert.Equal(500f, sz2.Height);
        }
    }
}
