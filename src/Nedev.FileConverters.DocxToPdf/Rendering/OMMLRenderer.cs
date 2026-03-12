using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Math;
using SkiaSharp;
using M = DocumentFormat.OpenXml.Math;

namespace Nedev.FileConverters.DocxToPdf.Rendering;

/// <summary>
/// OMML (Office Math Markup Language) 渲染器
/// 将 Word 数学公式渲染为图片
/// </summary>
public class OMMLRenderer
{
    private readonly float _baseFontSize;
    private readonly SKTypeface _mathTypeface;

    public OMMLRenderer(float baseFontSize = 16f)
    {
        _baseFontSize = baseFontSize;
        // 尝试加载 Cambria Math 字体，如果不存在则使用默认字体
        _mathTypeface = SKTypeface.FromFamilyName("Cambria Math") ?? SKTypeface.Default;
    }

    /// <summary>
    /// 渲染 OfficeMath 元素为 PNG 图片
    /// </summary>
    public byte[]? RenderToPng(M.OfficeMath officeMath, int maxWidth = 800)
    {
        try
        {
            // 首先计算公式尺寸
            var measureCanvas = new SKCanvas(new SKBitmap(1, 1));
            var size = MeasureElement(measureCanvas, officeMath, _baseFontSize);
            measureCanvas.Dispose();

            // 添加边距
            var padding = 10f;
            var width = (int)(size.Width + padding * 2);
            var height = (int)(size.Height + padding * 2);

            // 限制最大宽度
            if (width > maxWidth)
            {
                var scale = maxWidth / (float)width;
                width = maxWidth;
                height = (int)(height * scale);
            }

            // 创建画布
            var info = new SKImageInfo(width, height, SKColorType.Bgra8888, SKAlphaType.Premul);
            using var surface = SKSurface.Create(info);
            if (surface == null) return null;

            var canvas = surface.Canvas;
            canvas.Clear(SKColors.Transparent);

            // 渲染公式
            RenderElement(canvas, officeMath, padding, padding + size.Baseline, _baseFontSize);

            using var image = surface.Snapshot();
            using var data = image.Encode(SKEncodedImageFormat.Png, 90);
            return data.ToArray();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[OMMLRenderer] Failed to render: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 渲染单个元素
    /// </summary>
    private void RenderElement(SKCanvas canvas, OpenXmlElement element, float x, float y, float fontSize, bool isScript = false)
    {
        var actualFontSize = isScript ? fontSize * 0.7f : fontSize;

        switch (element)
        {
            case M.Run run:
                RenderRun(canvas, run, x, y, actualFontSize);
                break;

            case M.Fraction frac:
                RenderFraction(canvas, frac, x, y, actualFontSize);
                break;

            case M.Superscript sup:
                RenderSuperscript(canvas, sup, x, y, actualFontSize);
                break;

            case M.Subscript sub:
                RenderSubscript(canvas, sub, x, y, actualFontSize);
                break;

            case M.SubSuperscript subSup:
                RenderSubSuperscript(canvas, subSup, x, y, actualFontSize);
                break;

            case M.Radical rad:
                RenderRadical(canvas, rad, x, y, actualFontSize);
                break;

            case M.OfficeMath omath:
                // 递归渲染 OfficeMath 子元素
                var currentX = x;
                foreach (var child in omath.ChildElements)
                {
                    RenderElement(canvas, child, currentX, y, actualFontSize);
                    currentX += MeasureElement(canvas, child, actualFontSize).Width;
                }
                break;

            default:
                // 对于未知类型，尝试渲染其子元素
                if (element.HasChildren)
                {
                    var currX = x;
                    foreach (var child in element.ChildElements)
                    {
                        RenderElement(canvas, child, currX, y, actualFontSize);
                        currX += MeasureElement(canvas, child, actualFontSize).Width;
                    }
                }
                break;
        }
    }

    /// <summary>
    /// 渲染文本 Run
    /// </summary>
    private void RenderRun(SKCanvas canvas, M.Run run, float x, float y, float fontSize)
    {
        var text = run.GetFirstChild<M.Text>()?.Text ?? "";
        if (string.IsNullOrEmpty(text)) return;

        using var paint = CreateTextPaint(fontSize);
        canvas.DrawText(text, x, y, paint);
    }

    /// <summary>
    /// 渲染分数
    /// </summary>
    private void RenderFraction(SKCanvas canvas, M.Fraction frac, float x, float y, float fontSize)
    {
        var numerator = frac.GetFirstChild<M.Numerator>();
        var denominator = frac.GetFirstChild<M.Denominator>();

        if (numerator == null || denominator == null) return;

        var numMath = numerator.GetFirstChild<M.OfficeMath>();
        var denMath = denominator.GetFirstChild<M.OfficeMath>();

        if (numMath == null || denMath == null) return;

        var numSize = MeasureElement(canvas, numMath, fontSize * 0.8f);
        var denSize = MeasureElement(canvas, denMath, fontSize * 0.8f);

        var width = Math.Max(numSize.Width, denSize.Width) + 10;
        var lineY = y - fontSize * 0.3f;

        // 渲染分子（居中）
        var numX = x + (width - numSize.Width) / 2;
        RenderElement(canvas, numMath, numX, lineY - 5, fontSize * 0.8f);

        // 渲染分母（居中）
        var denX = x + (width - denSize.Width) / 2;
        RenderElement(canvas, denMath, denX, lineY + denSize.Height + 5, fontSize * 0.8f);

        // 绘制分数线
        using var linePaint = new SKPaint
        {
            Color = SKColors.Black,
            StrokeWidth = 1.5f,
            IsAntialias = true
        };
        canvas.DrawLine(x, lineY, x + width, lineY, linePaint);
    }

    /// <summary>
    /// 渲染上标
    /// </summary>
    private void RenderSuperscript(SKCanvas canvas, M.Superscript sup, float x, float y, float fontSize)
    {
        var baseElem = sup.GetFirstChild<M.Base>()?.GetFirstChild<M.OfficeMath>();
        var superElem = sup.GetFirstChild<M.SuperArgument>()?.GetFirstChild<M.OfficeMath>();

        if (baseElem == null) return;

        // 渲染基
        RenderElement(canvas, baseElem, x, y, fontSize);
        var baseWidth = MeasureElement(canvas, baseElem, fontSize).Width;

        // 渲染上标
        if (superElem != null)
        {
            RenderElement(canvas, superElem, x + baseWidth + 2, y - fontSize * 0.4f, fontSize * 0.7f, true);
        }
    }

    /// <summary>
    /// 渲染下标
    /// </summary>
    private void RenderSubscript(SKCanvas canvas, M.Subscript sub, float x, float y, float fontSize)
    {
        var baseElem = sub.GetFirstChild<M.Base>()?.GetFirstChild<M.OfficeMath>();
        var subElem = sub.GetFirstChild<M.SubArgument>()?.GetFirstChild<M.OfficeMath>();

        if (baseElem == null) return;

        // 渲染基
        RenderElement(canvas, baseElem, x, y, fontSize);
        var baseWidth = MeasureElement(canvas, baseElem, fontSize).Width;

        // 渲染下标
        if (subElem != null)
        {
            RenderElement(canvas, subElem, x + baseWidth + 2, y + fontSize * 0.2f, fontSize * 0.7f, true);
        }
    }

    /// <summary>
    /// 渲染上下标
    /// </summary>
    private void RenderSubSuperscript(SKCanvas canvas, M.SubSuperscript subSup, float x, float y, float fontSize)
    {
        var baseElem = subSup.GetFirstChild<M.Base>()?.GetFirstChild<M.OfficeMath>();
        var subElem = subSup.GetFirstChild<M.SubArgument>()?.GetFirstChild<M.OfficeMath>();
        var supElem = subSup.GetFirstChild<M.SuperArgument>()?.GetFirstChild<M.OfficeMath>();

        if (baseElem == null) return;

        // 渲染基
        RenderElement(canvas, baseElem, x, y, fontSize);
        var baseWidth = MeasureElement(canvas, baseElem, fontSize).Width;

        // 渲染下标和上标
        if (subElem != null)
        {
            RenderElement(canvas, subElem, x + baseWidth + 2, y + fontSize * 0.2f, fontSize * 0.7f, true);
        }
        if (supElem != null)
        {
            RenderElement(canvas, supElem, x + baseWidth + 2, y - fontSize * 0.4f, fontSize * 0.7f, true);
        }
    }

    /// <summary>
    /// 渲染根号
    /// </summary>
    private void RenderRadical(SKCanvas canvas, M.Radical rad, float x, float y, float fontSize)
    {
        // 获取被开方数
        var radicand = rad.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "radicand");
        var degree = rad.Elements().FirstOrDefault(e => e.LocalName == "deg" || e.LocalName == "degree");

        if (radicand == null) return;

        var radicandMath = radicand.GetFirstChild<M.OfficeMath>();
        if (radicandMath == null) return;

        var radicandSize = MeasureElement(canvas, radicandMath, fontSize);
        var radicalWidth = fontSize * 0.8f;

        // 渲染根指数（如果有）
        if (degree != null)
        {
            var degreeMath = degree.GetFirstChild<M.OfficeMath>();
            if (degreeMath != null)
            {
                RenderElement(canvas, degreeMath, x, y - radicandSize.Height * 0.5f, fontSize * 0.6f, true);
            }
        }

        // 绘制根号符号
        using var paint = CreateTextPaint(fontSize * 1.2f);
        canvas.DrawText("√", x + (degree != null ? fontSize * 0.4f : 0), y, paint);

        // 渲染被开方数
        var radicandX = x + radicalWidth + (degree != null ? fontSize * 0.3f : 0);
        RenderElement(canvas, radicandMath, radicandX, y, fontSize);

        // 绘制根号横线
        using var linePaint = new SKPaint
        {
            Color = SKColors.Black,
            StrokeWidth = 1.5f,
            IsAntialias = true
        };
        var lineY = y - radicandSize.Height + fontSize * 0.2f;
        canvas.DrawLine(radicandX - 2, lineY, radicandX + radicandSize.Width + 2, lineY, linePaint);
    }

    /// <summary>
    /// 测量元素尺寸
    /// </summary>
    private ElementSize MeasureElement(SKCanvas canvas, OpenXmlElement element, float fontSize)
    {
        var size = new ElementSize();

        switch (element)
        {
            case M.Run run:
                var text = run.GetFirstChild<M.Text>()?.Text ?? "";
                using (var paint = CreateTextPaint(fontSize))
                {
                    size.Width = paint.MeasureText(text);
                    size.Height = fontSize;
                    size.Baseline = fontSize * 0.8f;
                }
                break;

            case M.Fraction frac:
                var num = frac.GetFirstChild<M.Numerator>()?.GetFirstChild<M.OfficeMath>();
                var den = frac.GetFirstChild<M.Denominator>()?.GetFirstChild<M.OfficeMath>();
                if (num != null && den != null)
                {
                    var numSize = MeasureElement(canvas, num, fontSize * 0.8f);
                    var denSize = MeasureElement(canvas, den, fontSize * 0.8f);
                    size.Width = Math.Max(numSize.Width, denSize.Width) + 10;
                    size.Height = numSize.Height + denSize.Height + 15;
                    size.Baseline = numSize.Height + 7;
                }
                break;

            case M.Superscript sup:
                var supBase = sup.GetFirstChild<M.Base>()?.GetFirstChild<M.OfficeMath>();
                var super = sup.GetFirstChild<M.SuperArgument>()?.GetFirstChild<M.OfficeMath>();
                if (supBase != null)
                {
                    var baseSize = MeasureElement(canvas, supBase, fontSize);
                    var superSize = super != null ? MeasureElement(canvas, super, fontSize * 0.7f) : new ElementSize();
                    size.Width = baseSize.Width + superSize.Width + 2;
                    size.Height = Math.Max(baseSize.Height, superSize.Height + fontSize * 0.4f);
                    size.Baseline = baseSize.Baseline;
                }
                break;

            case M.Subscript sub:
                var subBase = sub.GetFirstChild<M.Base>()?.GetFirstChild<M.OfficeMath>();
                var subArg = sub.GetFirstChild<M.SubArgument>()?.GetFirstChild<M.OfficeMath>();
                if (subBase != null)
                {
                    var baseSize = MeasureElement(canvas, subBase, fontSize);
                    var subSize = subArg != null ? MeasureElement(canvas, subArg, fontSize * 0.7f) : new ElementSize();
                    size.Width = baseSize.Width + subSize.Width + 2;
                    size.Height = Math.Max(baseSize.Height, subSize.Height + fontSize * 0.2f);
                    size.Baseline = baseSize.Baseline;
                }
                break;

            case M.Radical rad:
                var radicand = rad.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "radicand");
                if (radicand != null)
                {
                    var radMath = radicand.GetFirstChild<M.OfficeMath>();
                    if (radMath != null)
                    {
                        var radSize = MeasureElement(canvas, radMath, fontSize);
                        size.Width = radSize.Width + fontSize * 0.8f + 5;
                        size.Height = radSize.Height + 5;
                        size.Baseline = radSize.Baseline;
                    }
                }
                break;

            case M.OfficeMath omath:
                float totalWidth = 0;
                float maxHeight = 0;
                float baseline = 0;
                foreach (var child in omath.ChildElements)
                {
                    var childSize = MeasureElement(canvas, child, fontSize);
                    totalWidth += childSize.Width;
                    maxHeight = Math.Max(maxHeight, childSize.Height);
                    baseline = Math.Max(baseline, childSize.Baseline);
                }
                size.Width = totalWidth;
                size.Height = maxHeight;
                size.Baseline = baseline;
                break;

            default:
                if (element.HasChildren)
                {
                    float totalW = 0;
                    float maxH = 0;
                    float baseL = 0;
                    foreach (var child in element.ChildElements)
                    {
                        var childSize = MeasureElement(canvas, child, fontSize);
                        totalW += childSize.Width;
                        maxH = Math.Max(maxH, childSize.Height);
                        baseL = Math.Max(baseL, childSize.Baseline);
                    }
                    size.Width = totalW;
                    size.Height = maxH;
                    size.Baseline = baseL;
                }
                break;
        }

        return size;
    }

    /// <summary>
    /// 创建文本画笔
    /// </summary>
    private SKPaint CreateTextPaint(float fontSize)
    {
        return new SKPaint
        {
            Color = SKColors.Black,
            TextSize = fontSize,
            Typeface = _mathTypeface,
            IsAntialias = true,
            TextAlign = SKTextAlign.Left
        };
    }

    /// <summary>
    /// 元素尺寸结构
    /// </summary>
    private struct ElementSize
    {
        public float Width { get; set; }
        public float Height { get; set; }
        public float Baseline { get; set; }
    }
}
