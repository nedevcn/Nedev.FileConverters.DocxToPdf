using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Rendering;
using System.IO;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class OMMLRendererTests
    {
        [Fact]
        public void RenderToPng_SimpleFraction_ReturnsPngBytes()
        {
            // 创建分数: 1/2
            var fraction = new M.Fraction();
            var numerator = new M.Numerator();
            var numRun = new M.Run(new M.Text("1"));
            numerator.Append(numRun);
            
            var denominator = new M.Denominator();
            var denRun = new M.Run(new M.Text("2"));
            denominator.Append(denRun);
            
            fraction.Append(numerator, denominator);
            
            var omath = new M.OfficeMath();
            omath.Append(fraction);

            // 渲染
            var renderer = new OMMLRenderer(20f);
            var pngBytes = renderer.RenderToPng(omath, 200);

            Assert.NotNull(pngBytes);
            Assert.True(pngBytes.Length > 0);
            // 验证 PNG 文件头
            Assert.Equal(0x89, pngBytes[0]);
            Assert.Equal(0x50, pngBytes[1]);
            Assert.Equal(0x4E, pngBytes[2]);
            Assert.Equal(0x47, pngBytes[3]);
        }

        [Fact]
        public void RenderToPng_Superscript_ReturnsPngBytes()
        {
            // 创建上标: x²
            var sup = new M.Superscript();
            var baseElem = new M.Base();
            baseElem.Append(new M.Run(new M.Text("x")));
            
            var superArg = new M.SuperArgument();
            superArg.Append(new M.Run(new M.Text("2")));
            
            sup.Append(baseElem, superArg);
            
            var omath = new M.OfficeMath();
            omath.Append(sup);

            var renderer = new OMMLRenderer(20f);
            var pngBytes = renderer.RenderToPng(omath, 200);

            Assert.NotNull(pngBytes);
            Assert.True(pngBytes.Length > 0);
        }

        [Fact]
        public void RenderToPng_Subscript_ReturnsPngBytes()
        {
            // 创建下标: a₁
            var sub = new M.Subscript();
            var baseElem = new M.Base();
            baseElem.Append(new M.Run(new M.Text("a")));
            
            var subArg = new M.SubArgument();
            subArg.Append(new M.Run(new M.Text("1")));
            
            sub.Append(baseElem, subArg);
            
            var omath = new M.OfficeMath();
            omath.Append(sub);

            var renderer = new OMMLRenderer(20f);
            var pngBytes = renderer.RenderToPng(omath, 200);

            Assert.NotNull(pngBytes);
            Assert.True(pngBytes.Length > 0);
        }

        [Fact]
        public void RenderToPng_Radical_ReturnsPngBytes()
        {
            // 创建根号: √x
            var rad = new M.Radical();
            
            // 被开方数
            var radicand = new M.OfficeMath();
            radicand.Append(new M.Run(new M.Text("x")));
            rad.Append(radicand);
            
            var omath = new M.OfficeMath();
            omath.Append(rad);

            var renderer = new OMMLRenderer(20f);
            var pngBytes = renderer.RenderToPng(omath, 200);

            Assert.NotNull(pngBytes);
            Assert.True(pngBytes.Length > 0);
        }

        [Fact]
        public void RenderToPng_RadicalWithDegree_ReturnsPngBytes()
        {
            // 创建带指数的根号: ³√x
            var rad = new M.Radical();
            
            // 根指数
            var degree = new M.OfficeMath();
            degree.Append(new M.Run(new M.Text("3")));
            rad.Append(degree);
            
            // 被开方数
            var radicand = new M.OfficeMath();
            radicand.Append(new M.Run(new M.Text("x")));
            rad.Append(radicand);
            
            var omath = new M.OfficeMath();
            omath.Append(rad);

            var renderer = new OMMLRenderer(20f);
            var pngBytes = renderer.RenderToPng(omath, 200);

            Assert.NotNull(pngBytes);
            Assert.True(pngBytes.Length > 0);
        }

        [Fact]
        public void RenderToPng_EmptyFormula_ReturnsNull()
        {
            var omath = new M.OfficeMath();
            
            var renderer = new OMMLRenderer(20f);
            var pngBytes = renderer.RenderToPng(omath, 200);

            // 空公式应该返回有效但很小的图片
            Assert.NotNull(pngBytes);
        }

        [Fact]
        public void RenderToPng_TextOnly_ReturnsPngBytes()
        {
            // 纯文本公式: E = mc²
            var omath = new M.OfficeMath();
            omath.Append(new M.Run(new M.Text("E = mc")));
            
            var sup = new M.Superscript();
            var supBase = new M.Base();
            supBase.Append(new M.Run(new M.Text("2")));
            sup.Append(supBase);
            omath.Append(sup);

            var renderer = new OMMLRenderer(20f);
            var pngBytes = renderer.RenderToPng(omath, 300);

            Assert.NotNull(pngBytes);
            Assert.True(pngBytes.Length > 0);
        }

        [Fact]
        public void RenderToPng_ComplexFormula_ReturnsPngBytes()
        {
            // 创建复杂公式: (a + b)²
            var sup = new M.Superscript();
            
            // 基: (a + b)
            var supBase = new M.Base();
            supBase.Append(new M.Run(new M.Text("(a + b)")));
            
            // 上标: 2
            var superArg = new M.SuperArgument();
            superArg.Append(new M.Run(new M.Text("2")));
            
            sup.Append(supBase, superArg);
            
            var omath = new M.OfficeMath();
            omath.Append(sup);

            var renderer = new OMMLRenderer(20f);
            var pngBytes = renderer.RenderToPng(omath, 300);

            Assert.NotNull(pngBytes);
            Assert.True(pngBytes.Length > 0);
        }
    }
}
