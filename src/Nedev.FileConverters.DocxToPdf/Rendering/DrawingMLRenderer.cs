using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.Models;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Nedev.FileConverters.DocxToPdf.Rendering;

/// <summary>
/// DrawingML 渲染器 - 使用 SkiaSharp 将 DrawingML 形状渲染为图片
/// </summary>
public class DrawingMLRenderer
{
    private readonly WordprocessingDocument _document;
    private readonly ConvertOptions _options;

    public DrawingMLRenderer(WordprocessingDocument document, ConvertOptions options)
    {
        _document = document;
        _options = options;
    }

    /// <summary>
    /// 渲染 DrawingML 元素为 PNG 图片
    /// </summary>
    public byte[]? RenderToPng(DocumentFormat.OpenXml.Wordprocessing.Drawing drawing, int pixelWidth, int pixelHeight)
    {
        try
        {
            var graphicData = drawing.Descendants<A.GraphicData>().FirstOrDefault();
            if (graphicData == null) return null;

            // 创建 SkiaSharp 画布
            var info = new SKImageInfo(pixelWidth, pixelHeight, SKColorType.Bgra8888, SKAlphaType.Premul);
            using var surface = SKSurface.Create(info);
            if (surface == null) return null;

            var canvas = surface.Canvas;
            canvas.Clear(SKColors.Transparent);

            // 渲染形状
            var shape = graphicData.Descendants<A.Shape>().FirstOrDefault();
            if (shape != null)
            {
                RenderShape(canvas, shape, pixelWidth, pixelHeight);
            }

            // 渲染图片
            var pic = graphicData.Descendants<A.Picture>().FirstOrDefault();
            if (pic != null)
            {
                RenderPicture(canvas, pic, pixelWidth, pixelHeight);
            }

            using var image = surface.Snapshot();
            using var data = image.Encode(SKEncodedImageFormat.Png, 90);
            return data.ToArray();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[DrawingMLRenderer] Failed to render: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 渲染单个形状
    /// </summary>
    private void RenderShape(SKCanvas canvas, A.Shape shape, int width, int height)
    {
        var spPr = shape.GetFirstChild<A.ShapeProperties>();
        if (spPr == null) return;

        // 获取变换
        var xfrm = spPr.GetFirstChild<A.Transform2D>();
        var offsetX = xfrm?.Offset?.X?.Value ?? 0;
        var offsetY = xfrm?.Offset?.Y?.Value ?? 0;
        var extCx = xfrm?.Extents?.Cx?.Value ?? 0;
        var extCy = xfrm?.Extents?.Cy?.Value ?? 0;

        // 转换为像素坐标
        var x = (int)EMU.ToPixels(offsetX);
        var y = (int)EMU.ToPixels(offsetY);
        var w = (int)EMU.ToPixels(extCx);
        var h = (int)EMU.ToPixels(extCy);

        if (w <= 0 || h <= 0) return;

        // 创建绘制区域
        var rect = new SKRect(x, y, x + w, y + h);

        // 渲染几何形状
        var prstGeom = spPr.GetFirstChild<A.PresetGeometry>();
        var custGeom = spPr.GetFirstChild<A.CustomGeometry>();

        if (prstGeom != null)
        {
            RenderPresetGeometry(canvas, rect, prstGeom, spPr);
        }
        else if (custGeom != null)
        {
            RenderCustomGeometry(canvas, rect, custGeom, spPr);
        }
        else
        {
            // 默认矩形
            RenderRectangle(canvas, rect, spPr);
        }

        // 渲染文本
        var txBody = shape.GetFirstChild<A.TextBody>();
        if (txBody != null)
        {
            RenderTextBody(canvas, rect, txBody);
        }
    }

    /// <summary>
    /// 渲染预设几何形状
    /// </summary>
    private void RenderPresetGeometry(SKCanvas canvas, SKRect rect, A.PresetGeometry prstGeom, A.ShapeProperties spPr)
    {
        var shapeType = prstGeom.Preset?.Value;
        var shapeTypeStr = shapeType.ToString()?.ToLower() ?? "rect";

        using var fillPaint = CreateFillPaint(spPr, rect);
        using var strokePaint = CreateStrokePaint(spPr);

        switch (shapeTypeStr)
        {
            case "ellipse":
            case "oval":
                canvas.DrawOval(rect, fillPaint);
                if (strokePaint != null) canvas.DrawOval(rect, strokePaint);
                break;

            case "roundrect":
            case "roundedrectangle":
                var radius = Math.Min(rect.Width, rect.Height) * 0.1f;
                canvas.DrawRoundRect(rect, radius, radius, fillPaint);
                if (strokePaint != null) canvas.DrawRoundRect(rect, radius, radius, strokePaint);
                break;

            case "line":
                canvas.DrawLine(rect.Left, rect.Top, rect.Right, rect.Bottom, strokePaint ?? fillPaint);
                break;

            case "triangle":
                DrawTriangle(canvas, rect, fillPaint, strokePaint);
                break;

            case "pentagon":
                DrawPolygon(canvas, rect, 5, fillPaint, strokePaint);
                break;

            case "hexagon":
                DrawPolygon(canvas, rect, 6, fillPaint, strokePaint);
                break;

            case "star":
                DrawStar(canvas, rect, 5, fillPaint, strokePaint);
                break;

            case "arrow":
                DrawArrow(canvas, rect, fillPaint, strokePaint);
                break;

            case "rect":
            default:
                RenderRectangle(canvas, rect, spPr);
                break;
        }
    }

    /// <summary>
    /// 渲染自定义几何形状
    /// </summary>
    private void RenderCustomGeometry(SKCanvas canvas, SKRect rect, A.CustomGeometry custGeom, A.ShapeProperties spPr)
    {
        // 简化处理：渲染为矩形
        RenderRectangle(canvas, rect, spPr);
    }

    /// <summary>
    /// 渲染矩形
    /// </summary>
    private void RenderRectangle(SKCanvas canvas, SKRect rect, A.ShapeProperties spPr)
    {
        using var fillPaint = CreateFillPaint(spPr, rect);
        using var strokePaint = CreateStrokePaint(spPr);

        canvas.DrawRect(rect, fillPaint);
        if (strokePaint != null)
        {
            canvas.DrawRect(rect, strokePaint);
        }
    }

    /// <summary>
    /// 渲染图片
    /// </summary>
    private void RenderPicture(SKCanvas canvas, A.Picture pic, int width, int height)
    {
        var blip = pic.GetFirstChild<A.BlipFill>()?.Blip;
        if (blip == null) return;

        var embedId = blip.Embed?.Value;
        if (string.IsNullOrEmpty(embedId)) return;

        try
        {
            var imagePart = _document.MainDocumentPart?.GetPartById(embedId) as ImagePart;
            if (imagePart == null) return;

            using var stream = imagePart.GetStream();
            using var skStream = new SKManagedStream(stream);
            using var bitmap = SKBitmap.Decode(skStream);
            if (bitmap == null) return;

            // 获取位置和尺寸
            var spPr = pic.GetFirstChild<A.ShapeProperties>();
            var xfrm = spPr?.GetFirstChild<A.Transform2D>();
            if (xfrm == null) return;

            var x = EMU.ToPixels(xfrm.Offset?.X?.Value ?? 0);
            var y = EMU.ToPixels(xfrm.Offset?.Y?.Value ?? 0);
            var w = EMU.ToPixels(xfrm.Extents?.Cx?.Value ?? 0);
            var h = EMU.ToPixels(xfrm.Extents?.Cy?.Value ?? 0);

            var destRect = new SKRect(x, y, x + w, y + h);
            canvas.DrawBitmap(bitmap, destRect);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[DrawingMLRenderer] Failed to render picture: {ex.Message}");
        }
    }

    /// <summary>
    /// 渲染文本体
    /// </summary>
    private void RenderTextBody(SKCanvas canvas, SKRect rect, A.TextBody txBody)
    {
        var bodyPr = txBody.GetFirstChild<A.BodyProperties>();
        var paragraphs = txBody.Elements<A.Paragraph>().ToList();
        if (paragraphs.Count == 0) return;

        // 计算文本区域
        var textRect = rect;
        var anchor = bodyPr?.Anchor?.Value;

        // 简化为顶部居中
        var y = rect.Top + 5;

        foreach (var para in paragraphs)
        {
            var paraText = string.Join("", para.Descendants<A.Text>().Select(t => t.Text));
            if (string.IsNullOrWhiteSpace(paraText)) continue;

            // 获取段落对齐方式
            var align = para.GetFirstChild<A.ParagraphProperties>()?.Alignment?.Value;

            // 获取字体属性
            var run = para.Descendants<A.Run>().FirstOrDefault();
            var runPr = run?.GetFirstChild<A.RunProperties>();
            var fontSize = (runPr?.FontSize?.Value ?? 1100) / 100f; // 转换为磅

            using var paint = new SKPaint
            {
                Color = SKColors.Black,
                IsAntialias = true
            };

            // 计算文本位置
            var textWidth = paint.MeasureText(paraText);
            float x;
            if (align != null && align == A.TextAlignmentTypeValues.Center)
            {
                x = rect.Left + (rect.Width - textWidth) / 2;
            }
            else if (align != null && align == A.TextAlignmentTypeValues.Right)
            {
                x = rect.Right - textWidth - 5;
            }
            else
            {
                x = rect.Left + 5;
            }

            canvas.DrawText(paraText, x, y + fontSize, paint);
            y += fontSize * 1.2f;

            if (y > rect.Bottom - 5) break;
        }
    }

    /// <summary>
    /// 创建填充画笔
    /// </summary>
    private SKPaint CreateFillPaint(A.ShapeProperties spPr, SKRect rect)
    {
        var paint = new SKPaint
        {
            IsAntialias = true,
            Style = SKPaintStyle.Fill
        };

        // 获取填充
        var solidFill = spPr.GetFirstChild<A.SolidFill>();
        var gradFill = spPr.GetFirstChild<A.GradientFill>();
        var noFill = spPr.GetFirstChild<A.NoFill>();

        if (noFill != null)
        {
            paint.Color = SKColors.Transparent;
        }
        else if (solidFill != null)
        {
            paint.Color = ExtractColor(solidFill);
        }
        else if (gradFill != null)
        {
            paint.Shader = CreateGradientShader(gradFill, rect);
            paint.Color = SKColors.White;
        }
        else
        {
            paint.Color = SKColors.White; // 默认白色填充
        }

        return paint;
    }

    /// <summary>
    /// 创建描边画笔
    /// </summary>
    private SKPaint? CreateStrokePaint(A.ShapeProperties spPr)
    {
        var ln = spPr.GetFirstChild<A.Outline>();
        if (ln == null) return null;

        // 检查是否无描边
        var noFill = ln.GetFirstChild<A.NoFill>();
        if (noFill != null) return null;

        var paint = new SKPaint
        {
            IsAntialias = true,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = (ln.Width?.Value ?? 12700) / 12700f // 转换为磅
        };

        // 获取描边颜色
        var solidFill = ln.GetFirstChild<A.SolidFill>();
        if (solidFill != null)
        {
            paint.Color = ExtractColor(solidFill);
        }
        else
        {
            paint.Color = SKColors.Black;
        }

        return paint;
    }

    /// <summary>
    /// 提取颜色
    /// </summary>
    private SKColor ExtractColor(A.SolidFill solidFill)
    {
        var rgb = solidFill.GetFirstChild<A.RgbColorModelHex>();
        if (rgb?.Val != null)
        {
            var hex = rgb.Val.Value;
            if (hex.Length == 6)
            {
                var r = Convert.ToByte(hex.Substring(0, 2), 16);
                var g = Convert.ToByte(hex.Substring(2, 2), 16);
                var b = Convert.ToByte(hex.Substring(4, 2), 16);
                return new SKColor(r, g, b);
            }
        }

        // 检查主题颜色
        var schemeClr = solidFill.GetFirstChild<A.SchemeColor>();
        if (schemeClr != null)
        {
            var schemeVal = schemeClr.Val?.Value;
            if (schemeVal == A.SchemeColorValues.Accent1) return new SKColor(0x4F, 0x81, 0xBD);
            if (schemeVal == A.SchemeColorValues.Accent2) return new SKColor(0xC0, 0x50, 0x4D);
            if (schemeVal == A.SchemeColorValues.Accent3) return new SKColor(0x9C, 0xBB, 0x58);
            if (schemeVal == A.SchemeColorValues.Accent4) return new SKColor(0x80, 0x64, 0xA0);
            if (schemeVal == A.SchemeColorValues.Accent5) return new SKColor(0x4B, 0xAC, 0xC6);
            if (schemeVal == A.SchemeColorValues.Accent6) return new SKColor(0xF7, 0x96, 0x46);
        }

        return SKColors.Black;
    }

    /// <summary>
    /// 创建渐变着色器
    /// </summary>
    private SKShader? CreateGradientShader(A.GradientFill gradFill, SKRect rect)
    {
        var gsList = gradFill.GetFirstChild<A.GradientStopList>();
        if (gsList == null) return null;

        var stops = gsList.Elements<A.GradientStop>().ToList();
        if (stops.Count < 2) return null;

        var colors = stops.Select(s =>
        {
            var rgb = s.GetFirstChild<A.RgbColorModelHex>();
            if (rgb?.Val != null)
            {
                var hex = rgb.Val.Value;
                if (hex.Length == 6)
                {
                    var r = Convert.ToByte(hex.Substring(0, 2), 16);
                    var g = Convert.ToByte(hex.Substring(2, 2), 16);
                    var b = Convert.ToByte(hex.Substring(4, 2), 16);
                    return new SKColor(r, g, b);
                }
            }
            return SKColors.White;
        }).ToArray();

        var positions = stops.Select(s => (s.Position?.Value ?? 0) / 100000f).ToArray();

        // 获取渐变类型
        var gradType = gradFill.GetFirstChild<A.LinearGradientFill>();
        if (gradType != null)
        {
            var angle = (gradType.Angle?.Value ?? 0) / 60000f; // 转换为度
            var radians = angle * MathF.PI / 180f;

            var x0 = rect.Left + rect.Width * (0.5f - 0.5f * MathF.Cos(radians));
            var y0 = rect.Top + rect.Height * (0.5f - 0.5f * MathF.Sin(radians));
            var x1 = rect.Left + rect.Width * (0.5f + 0.5f * MathF.Cos(radians));
            var y1 = rect.Top + rect.Height * (0.5f + 0.5f * MathF.Sin(radians));

            return SKShader.CreateLinearGradient(
                new SKPoint(x0, y0),
                new SKPoint(x1, y1),
                colors,
                positions,
                SKShaderTileMode.Clamp);
        }

        // 径向渐变 - 使用 PathGradientFill 替代
        var pathGrad = gradFill.GetFirstChild<A.PathGradientFill>();
        if (pathGrad != null)
        {
            var centerX = rect.MidX;
            var centerY = rect.MidY;
            var radius = Math.Min(rect.Width, rect.Height) / 2;

            return SKShader.CreateRadialGradient(
                new SKPoint(centerX, centerY),
                radius,
                colors,
                positions,
                SKShaderTileMode.Clamp);
        }

        return null;
    }

    // 辅助绘图方法
    private void DrawTriangle(SKCanvas canvas, SKRect rect, SKPaint fillPaint, SKPaint? strokePaint)
    {
        var path = new SKPath();
        path.MoveTo(rect.Left + rect.Width / 2, rect.Top);
        path.LineTo(rect.Right, rect.Bottom);
        path.LineTo(rect.Left, rect.Bottom);
        path.Close();

        canvas.DrawPath(path, fillPaint);
        if (strokePaint != null) canvas.DrawPath(path, strokePaint);
    }

    private void DrawPolygon(SKCanvas canvas, SKRect rect, int sides, SKPaint fillPaint, SKPaint? strokePaint)
    {
        var path = new SKPath();
        var centerX = rect.MidX;
        var centerY = rect.MidY;
        var radius = Math.Min(rect.Width, rect.Height) / 2;

        for (int i = 0; i < sides; i++)
        {
            var angle = (float)(i * 2 * Math.PI / sides - Math.PI / 2);
            var x = centerX + radius * MathF.Cos(angle);
            var y = centerY + radius * MathF.Sin(angle);

            if (i == 0)
                path.MoveTo(x, y);
            else
                path.LineTo(x, y);
        }
        path.Close();

        canvas.DrawPath(path, fillPaint);
        if (strokePaint != null) canvas.DrawPath(path, strokePaint);
    }

    private void DrawStar(SKCanvas canvas, SKRect rect, int points, SKPaint fillPaint, SKPaint? strokePaint)
    {
        var path = new SKPath();
        var centerX = rect.MidX;
        var centerY = rect.MidY;
        var outerRadius = Math.Min(rect.Width, rect.Height) / 2;
        var innerRadius = outerRadius * 0.4f;

        for (int i = 0; i < points * 2; i++)
        {
            var angle = (float)(i * Math.PI / points - Math.PI / 2);
            var radius = (i % 2 == 0) ? outerRadius : innerRadius;
            var x = centerX + radius * MathF.Cos(angle);
            var y = centerY + radius * MathF.Sin(angle);

            if (i == 0)
                path.MoveTo(x, y);
            else
                path.LineTo(x, y);
        }
        path.Close();

        canvas.DrawPath(path, fillPaint);
        if (strokePaint != null) canvas.DrawPath(path, strokePaint);
    }

    private void DrawArrow(SKCanvas canvas, SKRect rect, SKPaint fillPaint, SKPaint? strokePaint)
    {
        var path = new SKPath();
        var headWidth = rect.Width * 0.3f;
        var headHeight = rect.Height * 0.5f;
        var shaftHeight = rect.Height * 0.3f;

        path.MoveTo(rect.Left, rect.Top + (rect.Height - shaftHeight) / 2);
        path.LineTo(rect.Right - headWidth, rect.Top + (rect.Height - shaftHeight) / 2);
        path.LineTo(rect.Right - headWidth, rect.Top);
        path.LineTo(rect.Right, rect.MidY);
        path.LineTo(rect.Right - headWidth, rect.Bottom);
        path.LineTo(rect.Right - headWidth, rect.Top + (rect.Height + shaftHeight) / 2);
        path.LineTo(rect.Left, rect.Top + (rect.Height + shaftHeight) / 2);
        path.Close();

        canvas.DrawPath(path, fillPaint);
        if (strokePaint != null) canvas.DrawPath(path, strokePaint);
    }
}

/// <summary>
/// EMU (English Metric Units) 转换辅助类
/// </summary>
public static class EMU
{
    public const float EMU_PER_INCH = 914400f;
    public const float PIXELS_PER_INCH = 96f;

    public static float ToPixels(long emu)
    {
        return emu * PIXELS_PER_INCH / EMU_PER_INCH;
    }

    public static long FromPixels(float pixels)
    {
        return (long)(pixels * EMU_PER_INCH / PIXELS_PER_INCH);
    }
}
