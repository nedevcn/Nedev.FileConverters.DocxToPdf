using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using Nedev.FileConverters.DocxToPdf.Models;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;

namespace Nedev.FileConverters.DocxToPdf.Rendering;

/// <summary>
/// SmartArt 渲染器 - 将 SmartArt 图形渲染为图片
/// </summary>
public class SmartArtRenderer
{
    private readonly WordprocessingDocument _document;
    private readonly ConvertOptions _options;

    public SmartArtRenderer(WordprocessingDocument document, ConvertOptions options)
    {
        _document = document;
        _options = options;
    }

    /// <summary>
    /// 渲染 SmartArt 为 PNG 图片
    /// </summary>
    public byte[]? RenderToPng(OpenXmlElement diagramElement, int pixelWidth, int pixelHeight)
    {
        try
        {
            // 获取关系 ID
            var relId = GetDiagramRelationshipId(diagramElement);
            if (string.IsNullOrEmpty(relId)) return null;

            var part = _document.MainDocumentPart?.GetPartById(relId);
            if (part == null) return null;

            // 提取 SmartArt 数据
            var smartArtData = ExtractSmartArtData(part);
            if (smartArtData == null || smartArtData.Nodes.Count == 0)
                return RenderPlaceholder("SmartArt", "No data", pixelWidth, pixelHeight);

            // 创建画布
            var info = new SKImageInfo(pixelWidth, pixelHeight, SKColorType.Bgra8888, SKAlphaType.Premul);
            using var surface = SKSurface.Create(info);
            if (surface == null) return null;

            var canvas = surface.Canvas;
            canvas.Clear(SKColors.White);

            // 根据布局类型渲染
            switch (smartArtData.LayoutType)
            {
                case SmartArtLayoutType.Hierarchy:
                    RenderHierarchyLayout(canvas, smartArtData, pixelWidth, pixelHeight);
                    break;
                case SmartArtLayoutType.Process:
                    RenderProcessLayout(canvas, smartArtData, pixelWidth, pixelHeight);
                    break;
                case SmartArtLayoutType.Cycle:
                    RenderCycleLayout(canvas, smartArtData, pixelWidth, pixelHeight);
                    break;
                case SmartArtLayoutType.Matrix:
                    RenderMatrixLayout(canvas, smartArtData, pixelWidth, pixelHeight);
                    break;
                case SmartArtLayoutType.Pyramid:
                    RenderPyramidLayout(canvas, smartArtData, pixelWidth, pixelHeight);
                    break;
                default:
                    RenderGenericLayout(canvas, smartArtData, pixelWidth, pixelHeight);
                    break;
            }

            using var image = surface.Snapshot();
            using var data = image.Encode(SKEncodedImageFormat.Png, 90);
            return data.ToArray();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[SmartArtRenderer] Failed to render SmartArt: {ex.Message}");
            return RenderPlaceholder("SmartArt", "Error", pixelWidth, pixelHeight);
        }
    }

    /// <summary>
    /// 从 Diagram 元素获取关系 ID
    /// </summary>
    private string? GetDiagramRelationshipId(OpenXmlElement diagramElement)
    {
        var diagramEl = diagramElement.Elements().FirstOrDefault(e => e.LocalName == "diagram");
        if (diagramEl == null) return null;

        var relAttr = diagramEl.GetAttributes().FirstOrDefault(a => a.LocalName == "id");
        return relAttr.Value;
    }

    /// <summary>
    /// 提取 SmartArt 数据
    /// </summary>
    private SmartArtData? ExtractSmartArtData(OpenXmlPart part)
    {
        var root = part.RootElement;
        if (root == null) return null;

        var data = new SmartArtData();

        // 提取标题
        var titleElement = root.GetFirstChild<DocumentFormat.OpenXml.Drawing.Diagrams.Title>();
        var titleText = titleElement?.InnerText;
        data.Title = titleText ?? "";

        // 确定布局类型（从根元素名称或属性）
        var layoutName = root.LocalName;
        data.LayoutType = DetermineLayoutType(layoutName);

        // 提取节点文本
        var texts = root.Descendants<A.Text>()
            .Select(t => t.Text)
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .ToList();

        for (int i = 0; i < texts.Count; i++)
        {
            data.Nodes.Add(new SmartArtNode
            {
                Id = i.ToString(),
                Text = texts[i],
                Level = 0
            });
        }

        return data;
    }

    /// <summary>
    /// 确定布局类型
    /// </summary>
    private SmartArtLayoutType DetermineLayoutType(string layoutName)
    {
        var name = layoutName.ToLower();
        if (name.Contains("hierarchy") || name.Contains("org"))
            return SmartArtLayoutType.Hierarchy;
        if (name.Contains("process") || name.Contains("flow"))
            return SmartArtLayoutType.Process;
        if (name.Contains("cycle"))
            return SmartArtLayoutType.Cycle;
        if (name.Contains("matrix"))
            return SmartArtLayoutType.Matrix;
        if (name.Contains("pyramid"))
            return SmartArtLayoutType.Pyramid;
        return SmartArtLayoutType.Other;
    }

    /// <summary>
    /// 渲染层次结构布局（组织结构图）
    /// </summary>
    private void RenderHierarchyLayout(SKCanvas canvas, SmartArtData data, int width, int height)
    {
        var margin = 40f;
        var availableWidth = width - margin * 2;
        var availableHeight = height - margin * 2;

        // 绘制标题
        if (!string.IsNullOrEmpty(data.Title))
        {
            using var titlePaint = new SKPaint
            {
                Color = SKColors.Black,
                TextSize = 16,
                IsAntialias = true
            };
            canvas.DrawText(data.Title, width / 2f, margin - 10, titlePaint);
        }

        if (data.Nodes.Count == 0) return;

        var nodeWidth = 120f;
        var nodeHeight = 40f;
        var levelHeight = availableHeight / 3;

        // 绘制根节点
        var rootX = width / 2f - nodeWidth / 2;
        var rootY = margin + 20;
        DrawNode(canvas, data.Nodes[0], rootX, rootY, nodeWidth, nodeHeight);

        // 绘制子节点
        if (data.Nodes.Count > 1)
        {
            var childY = rootY + levelHeight;
            var children = data.Nodes.Skip(1).ToList();
            var spacing = availableWidth / (children.Count + 1);

            for (int i = 0; i < children.Count; i++)
            {
                var childX = margin + spacing * (i + 1) - nodeWidth / 2;
                DrawNode(canvas, children[i], childX, childY, nodeWidth, nodeHeight);

                // 绘制连接线
                using var linePaint = new SKPaint
                {
                    Color = SKColors.Gray,
                    StrokeWidth = 2,
                    IsAntialias = true
                };
                canvas.DrawLine(rootX + nodeWidth / 2, rootY + nodeHeight, childX + nodeWidth / 2, childY, linePaint);
            }
        }
    }

    /// <summary>
    /// 渲染流程布局
    /// </summary>
    private void RenderProcessLayout(SKCanvas canvas, SmartArtData data, int width, int height)
    {
        var margin = 40f;
        var availableWidth = width - margin * 2;
        var nodeWidth = 100f;
        var nodeHeight = 60f;
        var spacing = 20f;

        var totalWidth = data.Nodes.Count * nodeWidth + (data.Nodes.Count - 1) * spacing;
        var startX = (width - totalWidth) / 2;
        var startY = (height - nodeHeight) / 2;

        for (int i = 0; i < data.Nodes.Count; i++)
        {
            var x = startX + i * (nodeWidth + spacing);
            DrawNode(canvas, data.Nodes[i], x, startY, nodeWidth, nodeHeight);

            // 绘制箭头连接
            if (i < data.Nodes.Count - 1)
            {
                using var arrowPaint = new SKPaint
                {
                    Color = SKColors.Gray,
                    StrokeWidth = 2,
                    IsAntialias = true
                };
                canvas.DrawLine(x + nodeWidth, startY + nodeHeight / 2, x + nodeWidth + spacing, startY + nodeHeight / 2, arrowPaint);
            }
        }
    }

    /// <summary>
    /// 渲染循环布局
    /// </summary>
    private void RenderCycleLayout(SKCanvas canvas, SmartArtData data, int width, int height)
    {
        var centerX = width / 2f;
        var centerY = height / 2f;
        var radius = Math.Min(width, height) / 3f;
        var nodeSize = 80f;

        for (int i = 0; i < data.Nodes.Count; i++)
        {
            var angle = (float)(i * 2 * Math.PI / data.Nodes.Count - Math.PI / 2);
            var x = centerX + radius * MathF.Cos(angle) - nodeSize / 2;
            var y = centerY + radius * MathF.Sin(angle) - nodeSize / 2;

            DrawNode(canvas, data.Nodes[i], x, y, nodeSize, nodeSize);
        }
    }

    /// <summary>
    /// 渲染矩阵布局
    /// </summary>
    private void RenderMatrixLayout(SKCanvas canvas, SmartArtData data, int width, int height)
    {
        var margin = 40f;
        var cols = (int)Math.Ceiling(Math.Sqrt(data.Nodes.Count));
        var rows = (int)Math.Ceiling((double)data.Nodes.Count / cols);

        var cellWidth = (width - margin * 2) / cols;
        var cellHeight = (height - margin * 2) / rows;
        var nodeSize = Math.Min(cellWidth, cellHeight) * 0.8f;

        for (int i = 0; i < data.Nodes.Count; i++)
        {
            var col = i % cols;
            var row = i / cols;

            var x = margin + col * cellWidth + (cellWidth - nodeSize) / 2;
            var y = margin + row * cellHeight + (cellHeight - nodeSize) / 2;

            DrawNode(canvas, data.Nodes[i], x, y, nodeSize, nodeSize);
        }
    }

    /// <summary>
    /// 渲染金字塔布局
    /// </summary>
    private void RenderPyramidLayout(SKCanvas canvas, SmartArtData data, int width, int height)
    {
        var margin = 40f;
        var availableWidth = width - margin * 2;
        var availableHeight = height - margin * 2;

        // 绘制标题
        if (!string.IsNullOrEmpty(data.Title))
        {
            using var titlePaint = new SKPaint
            {
                Color = SKColors.Black,
                TextSize = 16,
                IsAntialias = true
            };
            canvas.DrawText(data.Title, width / 2f, margin - 10, titlePaint);
        }

        if (data.Nodes.Count == 0) return;

        // 金字塔层级计算
        var levels = Math.Min(data.Nodes.Count, 5); // 最多5层
        var levelHeight = availableHeight / levels;

        for (int i = 0; i < levels && i < data.Nodes.Count; i++)
        {
            // 从顶层到底层，宽度递增
            var levelWidthRatio = (i + 1) / (float)levels;
            var levelWidth = availableWidth * levelWidthRatio * 0.8f; // 最大宽度为可用宽度的80%
            var levelX = margin + (availableWidth - levelWidth) / 2;
            var levelY = margin + i * levelHeight;

            // 绘制梯形（金字塔层级）
            using var pyramidPaint = new SKPaint
            {
                Color = new SKColor((byte)(0x44 + i * 20), 0x72, (byte)(0xC4 - i * 20)),
                IsAntialias = true,
                Style = SKPaintStyle.Fill
            };

            using var borderPaint = new SKPaint
            {
                Color = SKColors.White,
                StrokeWidth = 2,
                IsAntialias = true,
                Style = SKPaintStyle.Stroke
            };

            using var path = new SKPath();
            // 计算上层宽度（如果是顶层则为0）
            var upperWidthRatio = i / (float)levels;
            var upperWidth = availableWidth * upperWidthRatio * 0.8f;
            var upperX = margin + (availableWidth - upperWidth) / 2;

            // 绘制梯形
            path.MoveTo(upperX, levelY);
            path.LineTo(upperX + upperWidth, levelY);
            path.LineTo(levelX + levelWidth, levelY + levelHeight - 10);
            path.LineTo(levelX, levelY + levelHeight - 10);
            path.Close();

            canvas.DrawPath(path, pyramidPaint);
            canvas.DrawPath(path, borderPaint);

            // 绘制文本
            if (i < data.Nodes.Count)
            {
                using var textPaint = new SKPaint
                {
                    Color = SKColors.White,
                    TextSize = 12,
                    IsAntialias = true
                };

                var text = data.Nodes[i].Text;
                var textWidth = textPaint.MeasureText(text);
                var textX = levelX + (levelWidth - textWidth) / 2;
                var textY = levelY + levelHeight / 2 + 4;

                canvas.DrawText(text, textX, textY, textPaint);
            }
        }
    }

    /// <summary>
    /// 渲染通用布局
    /// </summary>
    private void RenderGenericLayout(SKCanvas canvas, SmartArtData data, int width, int height)
    {
        // 默认使用流程布局
        RenderProcessLayout(canvas, data, width, height);
    }

    /// <summary>
    /// 绘制单个节点
    /// </summary>
    private void DrawNode(SKCanvas canvas, SmartArtNode node, float x, float y, float width, float height)
    {
        // 绘制节点背景
        using var bgPaint = new SKPaint
        {
            Color = new SKColor(0x44, 0x72, 0xC4),
            IsAntialias = true,
            Style = SKPaintStyle.Fill
        };
        canvas.DrawRoundRect(x, y, width, height, 5, 5, bgPaint);

        // 绘制边框
        using var borderPaint = new SKPaint
        {
            Color = SKColors.DarkBlue,
            StrokeWidth = 2,
            IsAntialias = true,
            Style = SKPaintStyle.Stroke
        };
        canvas.DrawRoundRect(x, y, width, height, 5, 5, borderPaint);

        // 绘制文本
        if (!string.IsNullOrEmpty(node.Text))
        {
            using var textPaint = new SKPaint
            {
                Color = SKColors.White,
                TextSize = 12,
                IsAntialias = true
            };

            // 文本截断
            var text = node.Text.Length > 20 ? node.Text[..17] + "..." : node.Text;
            var textWidth = textPaint.MeasureText(text);
            var textX = x + (width - textWidth) / 2;
            var textY = y + height / 2 + 4;

            canvas.DrawText(text, textX, textY, textPaint);
        }
    }

    /// <summary>
    /// 渲染占位符
    /// </summary>
    private static byte[]? RenderPlaceholder(string typeLabel, string? summary, int pixelWidth, int pixelHeight)
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
                var textY = startY;
                foreach (var line in lines)
                {
                    canvas.DrawText(line, rect.Left + 6, textY, SKTextAlign.Left, bodyFont, bodyTextPaint);
                    textY += lineHeight;
                    if (textY > rect.Bottom - 4) break;
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
    /// 文本换行
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

/// <summary>
/// SmartArt 数据
/// </summary>
public class SmartArtData
{
    public string Title { get; set; } = "";
    public SmartArtLayoutType LayoutType { get; set; } = SmartArtLayoutType.Other;
    public List<SmartArtNode> Nodes { get; set; } = [];
}

/// <summary>
/// SmartArt 节点
/// </summary>
public class SmartArtNode
{
    public string Id { get; set; } = "";
    public string Text { get; set; } = "";
    public int Level { get; set; }
    public string ParentId { get; set; } = "";
}

/// <summary>
/// SmartArt 布局类型
/// </summary>
public enum SmartArtLayoutType
{
    Hierarchy,
    Process,
    Cycle,
    Matrix,
    Pyramid,
    Other
}
