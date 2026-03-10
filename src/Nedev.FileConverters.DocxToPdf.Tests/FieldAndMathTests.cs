using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.Models;
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
    }
}
