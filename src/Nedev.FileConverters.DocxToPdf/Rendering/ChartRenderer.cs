using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Math;
using Nedev.FileConverters.DocxToPdf.Models;
using SkiaSharp;
using M = DocumentFormat.OpenXml.Math;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Nedev.FileConverters.DocxToPdf.Rendering;

/// <summary>
/// Chart 图表渲染器 - 将 Chart 渲染为图片
/// </summary>
public class ChartRenderer
{
    private readonly WordprocessingDocument _document;
    private readonly ConvertOptions _options;

    public ChartRenderer(WordprocessingDocument document, ConvertOptions options)
    {
        _document = document;
        _options = options;
    }

    /// <summary>
    /// 渲染 Chart 为 PNG 图片
    /// </summary>
    public byte[]? RenderToPng(OpenXmlElement chartElement, int pixelWidth, int pixelHeight)
    {
        try
        {
            // 获取 ChartPart
            var relId = GetChartRelationshipId(chartElement);
            if (string.IsNullOrEmpty(relId)) return null;

            var part = _document.MainDocumentPart?.GetPartById(relId);
            if (part is not ChartPart chartPart) return null;

            var chartSpace = chartPart.ChartSpace;
            if (chartSpace == null) return null;

            // 获取图表数据
            var chartData = ExtractChartData(chartSpace);
            if (chartData == null || chartData.Series.Count == 0)
                return RenderPlaceholder("Chart", "No data", pixelWidth, pixelHeight);

            // create higher‑resolution canvas for smoother output (2× scaling)
            const float renderScale = 2f;
            var realWidth = (int)(pixelWidth * renderScale);
            var realHeight = (int)(pixelHeight * renderScale);
            var info = new SKImageInfo(realWidth, realHeight, SKColorType.Bgra8888, SKAlphaType.Premul);
            using var surface = SKSurface.Create(info);
            if (surface == null) return null;

            var canvas = surface.Canvas;
            canvas.Scale(renderScale, renderScale);
            canvas.Clear(SKColors.White);

            // 根据图表类型渲染
            switch (chartData.ChartType)
            {
                case ChartType.Bar:
                case ChartType.Column:
                    RenderBarChart(canvas, chartData, pixelWidth, pixelHeight);
                    break;
                case ChartType.Line:
                    RenderLineChart(canvas, chartData, pixelWidth, pixelHeight);
                    break;
                case ChartType.Pie:
                    RenderPieChart(canvas, chartData, pixelWidth, pixelHeight);
                    break;
                case ChartType.Area:
                    RenderAreaChart(canvas, chartData, pixelWidth, pixelHeight);
                    break;
                case ChartType.Scatter:
                    RenderScatterChart(canvas, chartData, pixelWidth, pixelHeight);
                    break;
                case ChartType.Radar:
                    RenderRadarChart(canvas, chartData, pixelWidth, pixelHeight);
                    break;
                default:
                    RenderGenericChart(canvas, chartData, pixelWidth, pixelHeight);
                    break;
            }

            using var image = surface.Snapshot();
            using var data = image.Encode(SKEncodedImageFormat.Png, 90);
            return data.ToArray();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[ChartRenderer] Failed to render chart: {ex.Message}");
            return RenderPlaceholder("Chart", "Error", pixelWidth, pixelHeight);
        }
    }

    /// <summary>
    /// 从 Chart 元素获取关系 ID
    /// </summary>
    private string? GetChartRelationshipId(OpenXmlElement chartElement)
    {
        var chartEl = chartElement.Elements().FirstOrDefault(e => e.LocalName == "chart");
        if (chartEl == null) return null;

        var relAttr = chartEl.GetAttributes().FirstOrDefault(a => a.LocalName == "id");
        return relAttr.Value;
    }

    /// <summary>
    /// 提取图表数据
    /// </summary>
    private ChartData? ExtractChartData(C.ChartSpace chartSpace)
    {
        var chart = chartSpace.GetFirstChild<C.Chart>();
        if (chart == null) return null;

        var data = new ChartData();

        // 获取标题
        var title = chart.GetFirstChild<C.Title>();
        if (title != null)
        {
            var rich = title.GetFirstChild<C.RichText>();
            if (rich != null)
            {
                data.Title = string.Join("", rich.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text));
            }
        }

        // 确定图表类型
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return null;

        foreach (var child in plotArea.ChildElements)
        {
            switch (child.LocalName)
            {
                case "barChart":
                    data.ChartType = ChartType.Bar;
                    ExtractBarSeries(plotArea, data);
                    break;
                case "lineChart":
                    data.ChartType = ChartType.Line;
                    ExtractLineSeries(plotArea, data);
                    break;
                case "pieChart":
                    data.ChartType = ChartType.Pie;
                    ExtractPieSeries(plotArea, data);
                    break;
                case "areaChart":
                    data.ChartType = ChartType.Area;
                    ExtractAreaSeries(plotArea, data);
                    break;
                case "scatterChart":
                    data.ChartType = ChartType.Scatter;
                    ExtractScatterSeries(plotArea, data);
                    break;
                case "radarChart":
                    data.ChartType = ChartType.Radar;
                    ExtractRadarSeries(plotArea, data);
                    break;
            }
        }

        // 获取分类轴标签
        var catAx = plotArea.Elements().FirstOrDefault(e => e.LocalName == "catAx");
        if (catAx != null)
        {
            var titleEl = catAx.Elements().FirstOrDefault(e => e.LocalName == "title");
            if (titleEl != null)
            {
                data.CategoryAxisTitle = GetTextFromElement(titleEl);
            }
        }

        // 获取数值轴标签
        var valAx = plotArea.Elements().FirstOrDefault(e => e.LocalName == "valAx");
        if (valAx != null)
        {
            var titleEl = valAx.Elements().FirstOrDefault(e => e.LocalName == "title");
            if (titleEl != null)
            {
                data.ValueAxisTitle = GetTextFromElement(titleEl);
            }
        }

        return data;
    }

    /// <summary>
    /// 从元素中提取文本
    /// </summary>
    private string GetTextFromElement(OpenXmlElement element)
    {
        return string.Join("", element.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text));
    }

    /// <summary>
    /// 提取柱状图系列数据
    /// </summary>
    private void ExtractBarSeries(C.PlotArea plotArea, ChartData data)
    {
        var barChart = plotArea.Elements().FirstOrDefault(e => e.LocalName == "barChart");
        if (barChart == null) return;

        foreach (var ser in barChart.Elements().Where(e => e.LocalName == "ser"))
        {
            var series = ExtractSeriesData(ser);
            if (series != null) data.Series.Add(series);
        }
    }

    /// <summary>
    /// 提取折线图系列数据
    /// </summary>
    private void ExtractLineSeries(C.PlotArea plotArea, ChartData data)
    {
        var lineChart = plotArea.Elements().FirstOrDefault(e => e.LocalName == "lineChart");
        if (lineChart == null) return;

        foreach (var ser in lineChart.Elements().Where(e => e.LocalName == "ser"))
        {
            var series = ExtractSeriesData(ser);
            if (series != null) data.Series.Add(series);
        }
    }

    /// <summary>
    /// 提取饼图系列数据
    /// </summary>
    private void ExtractPieSeries(C.PlotArea plotArea, ChartData data)
    {
        var pieChart = plotArea.Elements().FirstOrDefault(e => e.LocalName == "pieChart");
        if (pieChart == null) return;

        foreach (var ser in pieChart.Elements().Where(e => e.LocalName == "ser"))
        {
            var series = ExtractSeriesData(ser);
            if (series != null) data.Series.Add(series);
        }
    }

    /// <summary>
    /// 提取面积图系列数据
    /// </summary>
    private void ExtractAreaSeries(C.PlotArea plotArea, ChartData data)
    {
        var areaChart = plotArea.Elements().FirstOrDefault(e => e.LocalName == "areaChart");
        if (areaChart == null) return;

        foreach (var ser in areaChart.Elements().Where(e => e.LocalName == "ser"))
        {
            var series = ExtractSeriesData(ser);
            if (series != null) data.Series.Add(series);
        }
    }

    /// <summary>
    /// 提取散点图系列数据
    /// </summary>
    private void ExtractScatterSeries(C.PlotArea plotArea, ChartData data)
    {
        var scatterChart = plotArea.Elements().FirstOrDefault(e => e.LocalName == "scatterChart");
        if (scatterChart == null) return;

        foreach (var ser in scatterChart.Elements().Where(e => e.LocalName == "ser"))
        {
            var series = ExtractSeriesData(ser);
            if (series != null) data.Series.Add(series);
        }
    }

    /// <summary>
    /// 提取雷达图系列数据
    /// </summary>
    private void ExtractRadarSeries(C.PlotArea plotArea, ChartData data)
    {
        var radarChart = plotArea.Elements().FirstOrDefault(e => e.LocalName == "radarChart");
        if (radarChart == null) return;

        foreach (var ser in radarChart.Elements().Where(e => e.LocalName == "ser"))
        {
            var series = ExtractSeriesData(ser);
            if (series != null) data.Series.Add(series);
        }
    }

    /// <summary>
    /// 提取系列数据
    /// </summary>
    private ChartSeries? ExtractSeriesData(OpenXmlElement serElement)
    {
        var series = new ChartSeries();

        // 系列名称
        var tx = serElement.Elements().FirstOrDefault(e => e.LocalName == "tx");
        if (tx != null)
        {
            series.Name = GetTextFromElement(tx);
        }

        // 分类数据
        var cat = serElement.Elements().FirstOrDefault(e => e.LocalName == "cat");
        if (cat != null)
        {
            var strLit = cat.Elements().FirstOrDefault(e => e.LocalName == "strLit");
            if (strLit != null)
            {
                foreach (var pt in strLit.Elements().Where(e => e.LocalName == "pt"))
                {
                    var v = pt.Elements().FirstOrDefault(e => e.LocalName == "v");
                    if (v != null)
                    {
                        series.Categories.Add(v.InnerText);
                    }
                }
            }
        }

        // 数值数据
        var val = serElement.Elements().FirstOrDefault(e => e.LocalName == "val");
        if (val != null)
        {
            var numLit = val.Elements().FirstOrDefault(e => e.LocalName == "numLit");
            if (numLit != null)
            {
                foreach (var pt in numLit.Elements().Where(e => e.LocalName == "pt"))
                {
                    var v = pt.Elements().FirstOrDefault(e => e.LocalName == "v");
                    if (v != null && double.TryParse(v.InnerText, out var value))
                    {
                        series.Values.Add(value);
                    }
                }
            }
        }

        return series.Values.Count > 0 ? series : null;
    }

    /// <summary>
    /// 渲染柱状图
    /// </summary>
    private void RenderBarChart(SKCanvas canvas, ChartData data, int width, int height)
    {
        var margin = 60f;
        var chartWidth = width - margin * 2;
        var chartHeight = height - margin * 2 - 30; // 留出标题空间

        // 绘制标题
        if (!string.IsNullOrEmpty(data.Title))
        {
            using var titlePaint = new SKPaint
            {
                Color = SKColors.Black,
                TextSize = 16,
                IsAntialias = true,
                Typeface = SKTypeface.Default
            };
            canvas.DrawText(data.Title, width / 2f, margin - 20, titlePaint);
        }

        // 计算最大值
        var maxValue = data.Series.SelectMany(s => s.Values).Max();
        if (maxValue == 0) maxValue = 1;

        // 绘制坐标轴
        using var axisPaint = new SKPaint
        {
            Color = SKColors.Gray,
            StrokeWidth = 1,
            IsAntialias = true
        };

        // Y轴
        canvas.DrawLine(margin, margin, margin, height - margin, axisPaint);
        // X轴
        canvas.DrawLine(margin, height - margin, width - margin, height - margin, axisPaint);

        // 绘制Y轴刻度和标签
        var ySteps = 5;
        using var labelPaint = new SKPaint
        {
            Color = SKColors.Gray,
            TextSize = 10,
            IsAntialias = true
        };

        for (int i = 0; i <= ySteps; i++)
        {
            var y = height - margin - (chartHeight * i / ySteps);
            var value = maxValue * i / ySteps;
            canvas.DrawText(value.ToString("F1"), margin - 5, y + 3, labelPaint);
            canvas.DrawLine(margin - 3, y, margin, y, axisPaint);
        }

        // 绘制柱状图
        var colors = new[] { SKColors.Blue, SKColors.Red, SKColors.Green, SKColors.Orange, SKColors.Purple };
        var barWidth = chartWidth / (data.Series[0].Categories.Count * data.Series.Count + data.Series[0].Categories.Count + 1);

        for (int catIdx = 0; catIdx < data.Series[0].Categories.Count; catIdx++)
        {
            var categoryX = margin + (catIdx * (data.Series.Count + 1) + 1) * barWidth;

            // 绘制分类标签
            if (catIdx < data.Series[0].Categories.Count)
            {
                canvas.DrawText(data.Series[0].Categories[catIdx], categoryX + barWidth * data.Series.Count / 2, height - margin + 15, labelPaint);
            }

            for (int serIdx = 0; serIdx < data.Series.Count; serIdx++)
            {
                var series = data.Series[serIdx];
                if (catIdx >= series.Values.Count) continue;

                var value = series.Values[catIdx];
                var barHeight = (float)(value / maxValue * chartHeight);
                var barX = categoryX + serIdx * barWidth;
                var barY = height - margin - barHeight;

                using var barPaint = new SKPaint
                {
                    Color = colors[serIdx % colors.Length],
                    IsAntialias = true,
                    Style = SKPaintStyle.Fill
                };

                canvas.DrawRect(barX, barY, barWidth - 2, barHeight, barPaint);

                // 绘制数值标签
                if (barHeight > 15)
                {
                    using var valuePaint = new SKPaint
                    {
                        Color = SKColors.White,
                        TextSize = 8,
                        IsAntialias = true
                    };
                    canvas.DrawText(value.ToString("F1"), barX + barWidth / 2 - 10, barY + barHeight / 2, valuePaint);
                }
            }
        }

        // 绘制图例
        RenderLegend(canvas, data, width - margin - 100, margin);
    }

    /// <summary>
    /// 渲染折线图
    /// </summary>
    private void RenderLineChart(SKCanvas canvas, ChartData data, int width, int height)
    {
        var margin = 60f;
        var chartWidth = width - margin * 2;
        var chartHeight = height - margin * 2 - 30;

        // 绘制标题
        if (!string.IsNullOrEmpty(data.Title))
        {
            using var titlePaint = new SKPaint
            {
                Color = SKColors.Black,
                TextSize = 16,
                IsAntialias = true
            };
            canvas.DrawText(data.Title, width / 2f, margin - 20, titlePaint);
        }

        // 计算最大值
        var maxValue = data.Series.SelectMany(s => s.Values).Max();
        if (maxValue == 0) maxValue = 1;

        // 绘制坐标轴
        using var axisPaint = new SKPaint
        {
            Color = SKColors.Gray,
            StrokeWidth = 1,
            IsAntialias = true
        };

        canvas.DrawLine(margin, margin, margin, height - margin, axisPaint);
        canvas.DrawLine(margin, height - margin, width - margin, height - margin, axisPaint);

        // 绘制Y轴刻度
        var ySteps = 5;
        using var labelPaint = new SKPaint
        {
            Color = SKColors.Gray,
            TextSize = 10,
            IsAntialias = true
        };

        for (int i = 0; i <= ySteps; i++)
        {
            var y = height - margin - (chartHeight * i / ySteps);
            var value = maxValue * i / ySteps;
            canvas.DrawText(value.ToString("F1"), margin - 5, y + 3, labelPaint);
        }

        // 绘制折线
        var colors = new[] { SKColors.Blue, SKColors.Red, SKColors.Green, SKColors.Orange, SKColors.Purple };
        var xStep = chartWidth / (data.Series[0].Categories.Count - 1);

        for (int serIdx = 0; serIdx < data.Series.Count; serIdx++)
        {
            var series = data.Series[serIdx];
            using var linePaint = new SKPaint
            {
                Color = colors[serIdx % colors.Length],
                StrokeWidth = 2,
                IsAntialias = true,
                Style = SKPaintStyle.Stroke
            };

            using var pointPaint = new SKPaint
            {
                Color = colors[serIdx % colors.Length],
                IsAntialias = true,
                Style = SKPaintStyle.Fill
            };

            var path = new SKPath();
            for (int i = 0; i < series.Values.Count; i++)
            {
                var x = margin + i * xStep;
                var y = height - margin - (float)(series.Values[i] / maxValue * chartHeight);

                if (i == 0)
                    path.MoveTo(x, y);
                else
                    path.LineTo(x, y);

                // 绘制数据点
                canvas.DrawCircle(x, y, 4, pointPaint);
            }

            canvas.DrawPath(path, linePaint);
        }

        // 绘制X轴标签
        for (int i = 0; i < data.Series[0].Categories.Count; i++)
        {
            var x = margin + i * xStep;
            canvas.DrawText(data.Series[0].Categories[i], x - 10, height - margin + 15, labelPaint);
        }

        // 绘制图例
        RenderLegend(canvas, data, width - margin - 100, margin);
    }

    /// <summary>
    /// 渲染饼图
    /// </summary>
    private void RenderPieChart(SKCanvas canvas, ChartData data, int width, int height)
    {
        var margin = 60f;
        var chartSize = Math.Min(width - margin * 2, height - margin * 2 - 30);
        var centerX = width / 2f;
        var centerY = height / 2f + 10;
        var radius = chartSize / 2f;

        // 绘制标题
        if (!string.IsNullOrEmpty(data.Title))
        {
            using var titlePaint = new SKPaint
            {
                Color = SKColors.Black,
                TextSize = 16,
                IsAntialias = true
            };
            canvas.DrawText(data.Title, width / 2f, margin - 20, titlePaint);
        }

        var colors = new[] { SKColors.Blue, SKColors.Red, SKColors.Green, SKColors.Orange, SKColors.Purple, SKColors.Cyan, SKColors.Yellow, SKColors.Magenta };

        // 只使用第一个系列
        if (data.Series.Count == 0) return;
        var series = data.Series[0];

        var total = series.Values.Sum();
        if (total == 0) return;

        var startAngle = 0f;
        using var labelPaint = new SKPaint
        {
            Color = SKColors.Black,
            TextSize = 10,
            IsAntialias = true
        };

        for (int i = 0; i < series.Values.Count; i++)
        {
            var value = series.Values[i];
            var sweepAngle = (float)(value / total * 360);

            using var piePaint = new SKPaint
            {
                Color = colors[i % colors.Length],
                IsAntialias = true,
                Style = SKPaintStyle.Fill
            };

            using var piePath = new SKPath();
            var rect = new SKRect(centerX - radius, centerY - radius, centerX + radius, centerY + radius);
            piePath.AddArc(rect, startAngle, sweepAngle);
            piePath.LineTo(centerX, centerY);
            piePath.Close();

            canvas.DrawPath(piePath, piePaint);

            // 绘制标签
            var midAngle = startAngle + sweepAngle / 2;
            var labelRadius = radius * 0.7f;
            var labelX = centerX + labelRadius * MathF.Cos(midAngle * MathF.PI / 180);
            var labelY = centerY + labelRadius * MathF.Sin(midAngle * MathF.PI / 180);

            var percentage = value / total * 100;
            canvas.DrawText($"{percentage:F1}%", labelX - 15, labelY, labelPaint);

            startAngle += sweepAngle;
        }

        // 绘制图例
        RenderLegend(canvas, data, width - margin - 80, margin);
    }

    /// <summary>
    /// 渲染面积图
    /// </summary>
    private void RenderAreaChart(SKCanvas canvas, ChartData data, int width, int height)
    {
        var margin = 60f;
        var chartWidth = width - margin * 2;
        var chartHeight = height - margin * 2 - 30;

        // 绘制标题
        if (!string.IsNullOrEmpty(data.Title))
        {
            using var titlePaint = new SKPaint
            {
                Color = SKColors.Black,
                TextSize = 16,
                IsAntialias = true
            };
            canvas.DrawText(data.Title, width / 2f, margin - 20, titlePaint);
        }

        // 计算最大值
        var maxValue = data.Series.SelectMany(s => s.Values).Max();
        if (maxValue == 0) maxValue = 1;

        // 绘制坐标轴
        using var axisPaint = new SKPaint
        {
            Color = SKColors.Gray,
            StrokeWidth = 1,
            IsAntialias = true
        };

        canvas.DrawLine(margin, margin, margin, height - margin, axisPaint);
        canvas.DrawLine(margin, height - margin, width - margin, height - margin, axisPaint);

        // 绘制Y轴刻度和标签
        var ySteps = 5;
        using var labelPaint = new SKPaint
        {
            Color = SKColors.Gray,
            TextSize = 10,
            IsAntialias = true
        };

        for (int i = 0; i <= ySteps; i++)
        {
            var y = height - margin - (chartHeight * i / ySteps);
            var value = maxValue * i / ySteps;
            canvas.DrawText(value.ToString("F1"), margin - 5, y + 3, labelPaint);
            canvas.DrawLine(margin - 3, y, margin, y, axisPaint);
        }

        // 绘制面积图
        var colors = new[] { SKColors.Blue, SKColors.Red, SKColors.Green, SKColors.Orange, SKColors.Purple };
        var xStep = chartWidth / (data.Series[0].Categories.Count - 1);

        for (int serIdx = 0; serIdx < data.Series.Count; serIdx++)
        {
            var series = data.Series[serIdx];
            var color = colors[serIdx % colors.Length];

            using var areaPaint = new SKPaint
            {
                Color = color.WithAlpha(128), // 半透明填充
                IsAntialias = true,
                Style = SKPaintStyle.Fill
            };

            using var linePaint = new SKPaint
            {
                Color = color,
                StrokeWidth = 2,
                IsAntialias = true,
                Style = SKPaintStyle.Stroke
            };

            using var path = new SKPath();
            path.MoveTo(margin, height - margin);

            for (int i = 0; i < series.Values.Count; i++)
            {
                var x = margin + i * xStep;
                var y = height - margin - (float)(series.Values[i] / maxValue * chartHeight);
                
                if (i == 0)
                    path.LineTo(x, y);
                else
                    path.LineTo(x, y);
            }

            path.LineTo(margin + (series.Values.Count - 1) * xStep, height - margin);
            path.Close();

            canvas.DrawPath(path, areaPaint);

            // 绘制边界线
            using var linePath = new SKPath();
            for (int i = 0; i < series.Values.Count; i++)
            {
                var x = margin + i * xStep;
                var y = height - margin - (float)(series.Values[i] / maxValue * chartHeight);
                
                if (i == 0)
                    linePath.MoveTo(x, y);
                else
                    linePath.LineTo(x, y);
            }
            canvas.DrawPath(linePath, linePaint);
        }

        // 绘制X轴标签
        for (int i = 0; i < data.Series[0].Categories.Count; i++)
        {
            var x = margin + i * xStep;
            canvas.DrawText(data.Series[0].Categories[i], x - 10, height - margin + 15, labelPaint);
        }

        // 绘制图例
        RenderLegend(canvas, data, width - margin - 100, margin);
    }

    /// <summary>
    /// 渲染散点图
    /// </summary>
    private void RenderScatterChart(SKCanvas canvas, ChartData data, int width, int height)
    {
        var margin = 60f;
        var chartWidth = width - margin * 2;
        var chartHeight = height - margin * 2 - 30;

        // 绘制标题
        if (!string.IsNullOrEmpty(data.Title))
        {
            using var titlePaint = new SKPaint
            {
                Color = SKColors.Black,
                TextSize = 16,
                IsAntialias = true
            };
            canvas.DrawText(data.Title, width / 2f, margin - 20, titlePaint);
        }

        // 计算最大值
        var maxValue = data.Series.SelectMany(s => s.Values).Max();
        if (maxValue == 0) maxValue = 1;

        // 绘制坐标轴
        using var axisPaint = new SKPaint
        {
            Color = SKColors.Gray,
            StrokeWidth = 1,
            IsAntialias = true
        };

        canvas.DrawLine(margin, margin, margin, height - margin, axisPaint);
        canvas.DrawLine(margin, height - margin, width - margin, height - margin, axisPaint);

        // 绘制Y轴刻度和标签
        var ySteps = 5;
        using var labelPaint = new SKPaint
        {
            Color = SKColors.Gray,
            TextSize = 10,
            IsAntialias = true
        };

        for (int i = 0; i <= ySteps; i++)
        {
            var y = height - margin - (chartHeight * i / ySteps);
            var value = maxValue * i / ySteps;
            canvas.DrawText(value.ToString("F1"), margin - 5, y + 3, labelPaint);
            canvas.DrawLine(margin - 3, y, margin, y, axisPaint);
        }

        // 绘制散点
        var colors = new[] { SKColors.Blue, SKColors.Red, SKColors.Green, SKColors.Orange, SKColors.Purple };
        var xStep = chartWidth / (data.Series[0].Categories.Count - 1);

        for (int serIdx = 0; serIdx < data.Series.Count; serIdx++)
        {
            var series = data.Series[serIdx];
            var color = colors[serIdx % colors.Length];

            using var pointPaint = new SKPaint
            {
                Color = color,
                IsAntialias = true,
                Style = SKPaintStyle.Fill
            };

            for (int i = 0; i < series.Values.Count; i++)
            {
                var x = margin + i * xStep;
                var y = height - margin - (float)(series.Values[i] / maxValue * chartHeight);
                
                // 绘制圆点
                canvas.DrawCircle(x, y, 6, pointPaint);
            }
        }

        // 绘制X轴标签
        for (int i = 0; i < data.Series[0].Categories.Count; i++)
        {
            var x = margin + i * xStep;
            canvas.DrawText(data.Series[0].Categories[i], x - 10, height - margin + 15, labelPaint);
        }

        // 绘制图例
        RenderLegend(canvas, data, width - margin - 100, margin);
    }

    /// <summary>
    /// 渲染雷达图
    /// </summary>
    private void RenderRadarChart(SKCanvas canvas, ChartData data, int width, int height)
    {
        var margin = 60f;
        var chartSize = Math.Min(width - margin * 2, height - margin * 2 - 30);
        var centerX = width / 2f;
        var centerY = height / 2f + 10;
        var radius = chartSize / 2f;

        // 绘制标题
        if (!string.IsNullOrEmpty(data.Title))
        {
            using var titlePaint = new SKPaint
            {
                Color = SKColors.Black,
                TextSize = 16,
                IsAntialias = true
            };
            canvas.DrawText(data.Title, width / 2f, margin - 20, titlePaint);
        }

        var categories = data.Series[0].Categories;
        var numCategories = categories.Count;
        var angleStep = 2 * MathF.PI / numCategories;

        // 绘制网格
        using var gridPaint = new SKPaint
        {
            Color = SKColors.LightGray,
            StrokeWidth = 1,
            IsAntialias = true
        };

        // 绘制同心圆网格
        for (int i = 1; i <= 5; i++)
        {
            var r = radius * i / 5;
            canvas.DrawCircle(centerX, centerY, r, gridPaint);
        }

        // 绘制轴线
        using var axisPaint = new SKPaint
        {
            Color = SKColors.Gray,
            StrokeWidth = 1,
            IsAntialias = true
        };

        for (int i = 0; i < numCategories; i++)
        {
            var angle = i * angleStep - MathF.PI / 2;
            var x = centerX + radius * MathF.Cos(angle);
            var y = centerY + radius * MathF.Sin(angle);
            canvas.DrawLine(centerX, centerY, x, y, axisPaint);

            // 绘制标签
            using var labelPaint = new SKPaint
            {
                Color = SKColors.Gray,
                TextSize = 10,
                IsAntialias = true
            };
            var labelX = centerX + (radius + 20) * MathF.Cos(angle);
            var labelY = centerY + (radius + 20) * MathF.Sin(angle);
            canvas.DrawText(categories[i], labelX - 20, labelY, labelPaint);
        }

        // 绘制数据系列
        var colors = new[] { SKColors.Blue, SKColors.Red, SKColors.Green, SKColors.Orange, SKColors.Purple };
        var maxValue = data.Series.SelectMany(s => s.Values).Max();
        if (maxValue == 0) maxValue = 1;

        for (int serIdx = 0; serIdx < data.Series.Count; serIdx++)
        {
            var series = data.Series[serIdx];
            var color = colors[serIdx % colors.Length];

            using var fillPaint = new SKPaint
            {
                Color = color.WithAlpha(64),
                IsAntialias = true,
                Style = SKPaintStyle.Fill
            };

            using var linePaint = new SKPaint
            {
                Color = color,
                StrokeWidth = 2,
                IsAntialias = true,
                Style = SKPaintStyle.Stroke
            };

            using var path = new SKPath();
            for (int i = 0; i < series.Values.Count; i++)
            {
                var angle = i * angleStep - MathF.PI / 2;
                var r = radius * (float)(series.Values[i] / maxValue);
                var x = centerX + r * MathF.Cos(angle);
                var y = centerY + r * MathF.Sin(angle);

                if (i == 0)
                    path.MoveTo(x, y);
                else
                    path.LineTo(x, y);
            }
            path.Close();

            canvas.DrawPath(path, fillPaint);
            canvas.DrawPath(path, linePaint);

            // 绘制数据点
            using var pointPaint = new SKPaint
            {
                Color = color,
                IsAntialias = true,
                Style = SKPaintStyle.Fill
            };

            for (int i = 0; i < series.Values.Count; i++)
            {
                var angle = i * angleStep - MathF.PI / 2;
                var r = radius * (float)(series.Values[i] / maxValue);
                var x = centerX + r * MathF.Cos(angle);
                var y = centerY + r * MathF.Sin(angle);
                canvas.DrawCircle(x, y, 4, pointPaint);
            }
        }

        // 绘制图例
        RenderLegend(canvas, data, width - margin - 100, margin);
    }

    /// <summary>
    /// 渲染通用图表
    /// </summary>
    private void RenderGenericChart(SKCanvas canvas, ChartData data, int width, int height)
    {
        // 默认使用柱状图渲染
        RenderBarChart(canvas, data, width, height);
    }

    /// <summary>
    /// 渲染图例
    /// </summary>
    private void RenderLegend(SKCanvas canvas, ChartData data, float x, float y)
    {
        var colors = new[] { SKColors.Blue, SKColors.Red, SKColors.Green, SKColors.Orange, SKColors.Purple };
        var legendY = y;

        using var textPaint = new SKPaint
        {
            Color = SKColors.Black,
            TextSize = 10,
            IsAntialias = true
        };

        for (int i = 0; i < data.Series.Count; i++)
        {
            using var colorPaint = new SKPaint
            {
                Color = colors[i % colors.Length],
                IsAntialias = true,
                Style = SKPaintStyle.Fill
            };

            canvas.DrawRect(x, legendY, 12, 12, colorPaint);
            canvas.DrawText(data.Series[i].Name ?? $"Series {i + 1}", x + 16, legendY + 10, textPaint);

            legendY += 18;
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
/// 图表数据
/// </summary>
public class ChartData
{
    public string Title { get; set; } = "";
    public ChartType ChartType { get; set; } = ChartType.Bar;
    public List<ChartSeries> Series { get; set; } = [];
    public string CategoryAxisTitle { get; set; } = "";
    public string ValueAxisTitle { get; set; } = "";
}

/// <summary>
/// 图表系列
/// </summary>
public class ChartSeries
{
    public string Name { get; set; } = "";
    public List<string> Categories { get; set; } = [];
    public List<double> Values { get; set; } = [];
}

/// <summary>
/// 图表类型
/// </summary>
public enum ChartType
{
    Bar,
    Column,
    Line,
    Pie,
    Area,
    Scatter,
    Radar
}
