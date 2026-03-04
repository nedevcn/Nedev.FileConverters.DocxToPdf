using System.Linq;
using Nedev.DocxToPdf.Models;
using Nedev.DocxToPdf.PdfEngine;

namespace Nedev.DocxToPdf.Helpers;

/// <summary>
/// PDF 水印渲染器
/// </summary>
public static class WatermarkRenderer
{
    /// <summary>
    /// 为 PDF 添加水印
    /// </summary>
    public static void ApplyWatermark(Stream inputStream, Stream outputStream, WatermarkOptions options)
    {
        using var ms = new MemoryStream();
        inputStream.CopyTo(ms);
        var pdfBytes = ms.ToArray();
        
        using var reader = new PdfReader(pdfBytes);
        using var stamper = new PdfStamper(reader, outputStream);
        
        var totalPages = reader.NumberOfPages;
        
        for (var i = 1; i <= totalPages; i++)
        {
            var pageSize = reader.GetPageSize(i);
            var contentByte = stamper.GetOverContent(i);
            
            if (!string.IsNullOrEmpty(options.ImagePath))
            {
                ApplyImageWatermark(contentByte, pageSize, options);
            }
            else if (!string.IsNullOrEmpty(options.Text))
            {
                if (options.Tiled)
                {
                    ApplyTiledTextWatermark(contentByte, pageSize, options);
                }
                else
                {
                    ApplyTextWatermark(contentByte, pageSize, options);
                }
            }
        }
    }

    /// <summary>
    /// 应用文本水印（单个）
    /// </summary>
    private static void ApplyTextWatermark(PdfContentByte contentByte, Rectangle pageSize, WatermarkOptions options)
    {
        contentByte.SaveState();
        contentByte.SetColorFill(options.Color ?? BaseColor.LightGray);
        
        var (x, y) = GetPosition(pageSize, options);
        
        contentByte.BeginText();
        
        var font = FontFactory.GetFont("STSong-Light", options.FontSize);
        contentByte.SetFontAndSize(font.Family, options.FontSize);
        
        var textWidth = font.GetWidthPoint(options.Text!);
        var textHeight = options.FontSize;
        
        var adjustedX = x - textWidth / 2f;
        var adjustedY = y - textHeight / 2f;
        
        contentByte.ShowTextAligned(
            Element.ALIGN_LEFT,
            options.Text!,
            adjustedX,
            adjustedY,
            options.Rotation
        );
        
        contentByte.EndText();
        contentByte.RestoreState();
    }

    /// <summary>
    /// 应用平铺文本水印
    /// </summary>
    private static void ApplyTiledTextWatermark(PdfContentByte contentByte, Rectangle pageSize, WatermarkOptions options)
    {
        contentByte.SaveState();
        contentByte.SetColorFill(options.Color ?? BaseColor.LightGray);
        
        var font = FontFactory.GetFont("STSong-Light", options.FontSize);
        var textWidth = font.GetWidthPoint(options.Text!);
        var textHeight = options.FontSize;
        
        var diagonal = (float)Math.Sqrt(Math.Pow(pageSize.Width, 2) + Math.Pow(pageSize.Height, 2));
        var spacing = Math.Max(textWidth, textHeight) + options.HorizontalSpacing;
        
        for (var y = -diagonal; y < diagonal * 2; y += spacing + textHeight)
        {
            for (var x = -diagonal; x < diagonal * 2; x += spacing + textWidth)
            {
                contentByte.BeginText();
                contentByte.SetFontAndSize(font.Family, options.FontSize);
                contentByte.ShowTextAligned(
                    Element.ALIGN_LEFT,
                    options.Text!,
                    x,
                    y,
                    options.Rotation
                );
                contentByte.EndText();
            }
        }
        
        contentByte.RestoreState();
    }

    /// <summary>
    /// 应用图片水印
    /// </summary>
    private static void ApplyImageWatermark(PdfContentByte contentByte, Rectangle pageSize, WatermarkOptions options)
    {
        if (!File.Exists(options.ImagePath)) return;
        
        contentByte.SaveState();
        
        using var imageStream = File.OpenRead(options.ImagePath);
        using var ms = new MemoryStream();
        imageStream.CopyTo(ms);
        var imageBytes = ms.ToArray();
        
        var image = Image.GetInstance(imageBytes);
        if (image == null) return;
        
        var imgWidth = image.OriginalWidth * 0.5f;
        var imgHeight = image.OriginalHeight * 0.5f;
        image.ScaleAbsolute(imgWidth, imgHeight);
        
        var (x, y) = GetPosition(pageSize, options, imgWidth, imgHeight);
        
        image.SetAbsolutePosition(x - imgWidth / 2f, y - imgHeight / 2f);
        
        contentByte.DrawImage(image, x - imgWidth / 2f, y - imgHeight / 2f);
        
        contentByte.RestoreState();
    }

    /// <summary>
    /// 获取水印位置坐标
    /// </summary>
    private static (float X, float Y) GetPosition(
        Rectangle pageSize, 
        WatermarkOptions options,
        float elementWidth = 0,
        float elementHeight = 0)
    {
        var margin = 50f;
        
        return options.Position switch
        {
            WatermarkPosition.TopLeft => (
                margin + elementWidth / 2f,
                pageSize.Height - margin - elementHeight / 2f
            ),
            WatermarkPosition.TopRight => (
                pageSize.Width - margin - elementWidth / 2f,
                pageSize.Height - margin - elementHeight / 2f
            ),
            WatermarkPosition.BottomLeft => (
                margin + elementWidth / 2f,
                margin + elementHeight / 2f
            ),
            WatermarkPosition.BottomRight => (
                pageSize.Width - margin - elementWidth / 2f,
                margin + elementHeight / 2f
            ),
            WatermarkPosition.Center => (
                pageSize.Width / 2f,
                pageSize.Height / 2f
            ),
            _ => (pageSize.Width / 2f, pageSize.Height / 2f)
        };
    }
}
