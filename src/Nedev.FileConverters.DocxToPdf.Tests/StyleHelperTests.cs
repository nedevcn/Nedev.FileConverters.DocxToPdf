using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.Models;
using SkiaSharp;
using Xunit;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class StyleHelperTests
    {
        [Fact]
        public void ResolveWordColor_HslColor_ConvertsToRgb()
        {
            // construct a dummy color node
            var color = new DocumentFormat.OpenXml.Wordprocessing.Color();
            var hsl = new DocumentFormat.OpenXml.Drawing.HslColorModelHex
            {
                Hue = "8000", // mid value
                Saturation = "FFFF",
                Luminance = "8000"
            };
            color.Append(hsl);
            var result = StyleHelper.ResolveWordColor(null, color);
            Assert.NotNull(result);
            // saturation full, luminance 50% -> vivid color, not grey
            Assert.False(result.R == result.G && result.G == result.B);
        }

        [Fact]
        public void ResolveWordColor_SystemColor_UsesLastColor()
        {
            var color = new DocumentFormat.OpenXml.Wordprocessing.Color();
            var sys = new DocumentFormat.OpenXml.Drawing.SystemColor { LastColor = "FF00FF" };
            color.Append(sys);
            var result = StyleHelper.ResolveWordColor(null, color);
            Assert.NotNull(result);
            Assert.Equal(0xFF, result.R);
            Assert.Equal(0x00, result.G);
            Assert.Equal(0xFF, result.B);
        }
    }
}
