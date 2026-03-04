using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Nedev.DocxToPdf.Models;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Nedev.DocxToPdf.Rasterization;

/// <summary>
/// 将复杂 DrawingML 对象（图表、SmartArt、形状等）栅格化为 PNG 图片。
/// 支持从 ChartPart 提取标题、从 Diagram 提取节点文本、从形状提取 txBody 文本。
/// </summary>
public sealed class DrawingRasterizer
{
    private readonly WordprocessingDocument _document;
    private readonly ConvertOptions _options;

    public DrawingRasterizer(WordprocessingDocument document, ConvertOptions options)
    {
        _document = document;
        _options = options;
    }

    /// <summary>
    /// 判断给定元素是否是可栅格化的复杂 DrawingML 对象。
    /// </summary>
    public bool CanRasterize(OpenXmlElement element)
    {
        var graphicData = FindGraphicData(element);
        if (graphicData == null) return false;

        var uri = graphicData.Uri?.Value ?? string.Empty;
        if (string.IsNullOrEmpty(uri)) return false;

        if (uri.Contains("/chart", StringComparison.OrdinalIgnoreCase))
            return _options.RasterizeCharts;
        if (uri.Contains("/diagram", StringComparison.OrdinalIgnoreCase))
            return _options.RasterizeSmartArt;
        if (uri.Contains("/main", StringComparison.OrdinalIgnoreCase))
            return _options.RasterizeShapes;

        return false;
    }

    /// <summary>
    /// 将复杂 DrawingML 对象栅格化为 PNG 字节数组。
    /// </summary>
    public byte[]? RasterizeToPng(OpenXmlElement element, int pixelWidth, int pixelHeight)
    {
        var graphicData = FindGraphicData(element);
        if (graphicData == null) return null;

        var uri = graphicData.Uri?.Value ?? string.Empty;
        string typeLabel;
        string? summary;

        if (uri.Contains("/chart", StringComparison.OrdinalIgnoreCase))
        {
            typeLabel = "Chart";
            summary = ExtractChartSummary(graphicData);
        }
        else if (uri.Contains("/diagram", StringComparison.OrdinalIgnoreCase))
        {
            typeLabel = "SmartArt";
            summary = ExtractDiagramSummary(graphicData);
        }
        else
        {
            typeLabel = "Shape";
            summary = ExtractShapeSummary(graphicData);
        }

        if (string.IsNullOrWhiteSpace(summary))
            summary = ExtractDrawingTextFallback(graphicData);

        return RasterizePlaceholder(typeLabel, summary, pixelWidth, pixelHeight);
    }

    private static A.GraphicData? FindGraphicData(OpenXmlElement element)
    {
        return element.Descendants<A.GraphicData>().FirstOrDefault();
    }

    /// <summary>
    /// 从 ChartPart 提取图表标题及简要信息（通用遍历，兼容 OpenXML 2.x/3.x）。
    /// </summary>
    private string? ExtractChartSummary(A.GraphicData graphicData)
    {
        var chartEl = graphicData.Elements().FirstOrDefault(e => e.LocalName == "chart");
        if (chartEl == null) return null;

        var relAttr = chartEl.GetAttributes().FirstOrDefault(a => a.LocalName == "id");
        var relId = relAttr.Value;
        if (string.IsNullOrEmpty(relId)) return null;

        try
        {
            var part = _document.MainDocumentPart?.GetPartById(relId);
            if (part is ChartPart chartPart)
            {
                var root = chartPart.RootElement;
                if (root == null) return null;

                var parts = new List<string>();

                // c:title / c:chart / c:title / c:tx / c:rich / a:t
                var titleEl = root.Descendants().FirstOrDefault(e => e.LocalName == "title");
                if (titleEl != null)
                {
                    var titleTexts = titleEl.Descendants<A.Text>().Select(t => t.Text).Where(t => !string.IsNullOrWhiteSpace(t));
                    var titleText = string.Join(" ", titleTexts);
                    if (!string.IsNullOrWhiteSpace(titleText))
                        parts.Add(titleText);
                }

                // c:ser 系列名称 (c:tx / c:strRef / c:strCache / c:pt / c:v)
                foreach (var ser in root.Descendants().Where(e => e.LocalName == "ser").Take(5))
                {
                    var tx = ser.Descendants().FirstOrDefault(e => e.LocalName == "tx");
                    var strCache = tx?.Descendants().FirstOrDefault(e => e.LocalName == "strCache");
                    var pt = strCache?.Descendants().FirstOrDefault(e => e.LocalName == "pt");
                    var v = pt?.Descendants().FirstOrDefault(e => e.LocalName == "v");
                    var serName = v?.InnerText?.Trim();
                    if (!string.IsNullOrWhiteSpace(serName))
                        parts.Add(serName);
                }

                return parts.Count > 0 ? string.Join(" · ", parts) : null;
            }
        }
        catch
        {
            // 忽略解析错误
        }

        return null;
    }

    /// <summary>
    /// 从 Diagram/SmartArt 提取节点文本（通用遍历）。
    /// </summary>
    private string? ExtractDiagramSummary(A.GraphicData graphicData)
    {
        var diagramEl = graphicData.Elements().FirstOrDefault(e => e.LocalName == "diagram");
        if (diagramEl == null) return null;

        var relAttr = diagramEl.GetAttributes().FirstOrDefault(a => a.LocalName == "id");
        var relId = relAttr.Value;
        if (string.IsNullOrEmpty(relId)) return null;

        try
        {
            var part = _document.MainDocumentPart?.GetPartById(relId);
            var root = (part as OpenXmlPart)?.RootElement;
            if (root != null)
            {
                var texts = root.Descendants<A.Text>()
                    .Select(t => t.Text)
                    .Where(t => !string.IsNullOrWhiteSpace(t))
                    .Take(20);
                return string.Join(" ", texts);
            }
        }
        catch
        {
            // 忽略
        }

        return ExtractDrawingTextFallback(graphicData);
    }

    /// <summary>
    /// 从形状/文本框提取 txBody 文本。
    /// </summary>
    private static string? ExtractShapeSummary(A.GraphicData graphicData)
    {
        // a:spTree / a:sp / a:txBody / a:p / a:r / a:t
        var txBodies = graphicData.Descendants().Where(e => e.LocalName == "txBody");
        var allTexts = new List<string>();
        foreach (var tx in txBodies)
        {
            var texts = tx.Descendants<A.Text>().Select(t => t.Text).Where(t => !string.IsNullOrWhiteSpace(t));
            allTexts.AddRange(texts);
        }
        return allTexts.Count > 0 ? string.Join(" ", allTexts) : null;
    }

    /// <summary>
    /// 通用回退：从 DrawingML 和 WordprocessingML 提取文本。
    /// </summary>
    private static string ExtractDrawingTextFallback(OpenXmlElement root)
    {
        var aTexts = root.Descendants<A.Text>().Select(t => t.Text).Where(t => !string.IsNullOrWhiteSpace(t)).ToList();
        if (aTexts.Count > 0)
            return string.Join(" ", aTexts);

        var wTexts = root.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
            .Select(t => t.Text)
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .ToList();
        if (wTexts.Count > 0)
            return string.Join(" ", wTexts);

        return string.Empty;
    }

    /// <summary>
    /// 用 SkiaSharp 生成占位 PNG。
    /// </summary>
    private static byte[]? RasterizePlaceholder(string typeLabel, string? summary, int pixelWidth, int pixelHeight)
    {
        pixelWidth = Math.Max(160, pixelWidth);
        pixelHeight = Math.Max(120, pixelHeight);

        try
        {
            var info = new SKImageInfo(pixelWidth, pixelHeight, SKColorType.Bgra8888, SKAlphaType.Premul);
            using var surface = SKSurface.Create(info);
            if (surface == null) return null;

            var canvas = surface.Canvas;
            canvas.Clear(SKColors.White);

            using var borderPaint = new SKPaint
            {
                Color = new SKColor(0x66, 0x66, 0x66),
                IsAntialias = true,
                StrokeWidth = 2,
                Style = SKPaintStyle.Stroke
            };

            using var headerPaint = new SKPaint
            {
                Color = new SKColor(0x33, 0x66, 0x99),
                IsAntialias = true,
                Style = SKPaintStyle.Fill
            };

            using var headerFont = new SKFont(SKTypeface.Default, 18);
            using var bodyFont = new SKFont(SKTypeface.Default, 13);

            using var headerTextPaint = new SKPaint
            {
                Color = SKColors.White,
                IsAntialias = true,
                IsStroke = false
            };

            using var bodyTextPaint = new SKPaint
            {
                Color = new SKColor(0x33, 0x33, 0x33),
                IsAntialias = true,
                IsStroke = false
            };

            var rect = new SKRect(1, 1, pixelWidth - 2, pixelHeight - 2);
            canvas.DrawRect(rect, borderPaint);

            var headerHeight = Math.Min(32, pixelHeight / 4f);
            var headerRect = new SKRect(rect.Left, rect.Top, rect.Right, rect.Top + headerHeight);
            canvas.DrawRect(headerRect, headerPaint);

            var typeText = typeLabel;
            var headerTextWidth = headerFont.MeasureText(typeText);
            var headerX = rect.MidX - headerTextWidth / 2f;
            var headerY = headerRect.MidY + headerFont.Size / 3f;
            canvas.DrawText(typeText, headerX, headerY, SKTextAlign.Left, headerFont, headerTextPaint);

            if (!string.IsNullOrWhiteSpace(summary))
            {
                var maxWidth = rect.Width - 12;
                var availableHeight = rect.Height - headerHeight - 8;
                var lineHeight = bodyFont.Size * 1.3f;
                var maxLines = Math.Max(1, (int)(availableHeight / lineHeight));

                var lines = WrapText(summary, bodyFont, maxWidth, maxLines);

                var startY = headerRect.Bottom + 8 + bodyFont.Size;
                var y = startY;
                foreach (var line in lines)
                {
                    canvas.DrawText(line, rect.Left + 6, y, SKTextAlign.Left, bodyFont, bodyTextPaint);
                    y += lineHeight;
                    if (y > rect.Bottom - 4) break;
                }
            }

            using var image = surface.Snapshot();
            using var data = image.Encode(SKEncodedImageFormat.Png, 90);
            return data.ToArray();
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 按最大宽度和行数换行。
    /// </summary>
    private static List<string> WrapText(string text, SKFont font, float maxWidth, int maxLines)
    {
        var words = text.Split([' ', '\r', '\n', '\t'], StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0) return [];

        var lines = new List<string>();
        var currentLine = new List<string>();

        foreach (var w in words)
        {
            var test = currentLine.Count == 0 ? w : string.Join(" ", currentLine) + " " + w;
            var testWidth = font.MeasureText(test);

            if (testWidth > maxWidth && currentLine.Count > 0)
            {
                lines.Add(string.Join(" ", currentLine));
                currentLine.Clear();
                currentLine.Add(w);
                if (lines.Count >= maxLines) break;
            }
            else
            {
                currentLine.Add(w);
            }
        }

        if (currentLine.Count > 0 && lines.Count < maxLines)
        {
            var lastLine = string.Join(" ", currentLine);
            if (lines.Count == maxLines - 1 && lastLine.Length > 50)
                lastLine = lastLine[..Math.Min(47, lastLine.Length)] + "...";
            lines.Add(lastLine);
        }

        return lines;
    }
}
