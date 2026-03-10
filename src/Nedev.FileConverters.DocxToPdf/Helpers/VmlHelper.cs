using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextChunk = Nedev.FileConverters.DocxToPdf.PdfEngine.Chunk;
using iTextFont = Nedev.FileConverters.DocxToPdf.PdfEngine.Font;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace Nedev.FileConverters.DocxToPdf.Helpers;

public static class VmlHelper
{
    public static List<IElement> ExtractVmlElements(OpenXmlElement parent, MainDocumentPart mainPart)
    {
        var elements = new List<IElement>();
        
        var vmlShapes = parent.Descendants().Where(e => 
            e.LocalName == "shape" || 
            e.LocalName == "rect" || 
            e.LocalName == "oval" ||
            e.LocalName == "line" ||
            e.LocalName == "polyline" ||
            e.LocalName == "curve"
        ).ToList();

        foreach (var shape in vmlShapes)
        {
            var wordArtText = ExtractWordArtText(shape);
            if (!string.IsNullOrWhiteSpace(wordArtText))
            {
                var styleAttr = shape.GetAttribute("style", "");
                var style = styleAttr.Value ?? "";
                var fontInfo = ParseVmlStyle(style);
                
                var font = FontFactory.GetFont(
                    fontInfo.FontName ?? "Arial", 
                    fontInfo.FontSize > 0 ? fontInfo.FontSize : 14f,
                    fontInfo.IsBold ? iTextFont.BOLD : iTextFont.NORMAL
                );
                
                var chunk = new iTextChunk(wordArtText, font);
                
                if (fontInfo.HasFill && fontInfo.FillColor != null)
                {
                    chunk.SetBackground(fontInfo.FillColor);
                }
                
                elements.Add(chunk);
            }
            else
            {
                // Basic Shape Support
                var shapeElement = CreateShapeElement(shape);
                if (shapeElement != null)
                {
                    elements.Add(shapeElement);
                }
            }
        }

        return elements;
    }

    private static string ExtractWordArtText(OpenXmlElement shape)
    {
        var textPath = shape.Descendants().FirstOrDefault(e => e.LocalName == "textpath");
        if (textPath != null)
        {
            var textAttr = textPath.GetAttribute("string", "");
            if (!string.IsNullOrEmpty(textAttr.Value))
            {
                return textAttr.Value;
            }
        }

        var txxt = shape.Descendants().FirstOrDefault(e => e.LocalName == "txxt");
        if (txxt != null)
        {
            var textContent = txxt.GetAttribute("text", "");
            if (!string.IsNullOrEmpty(textContent.Value))
            {
                return textContent.Value;
            }
        }

        return "";
    }

    private static VmlStyleInfo ParseVmlStyle(string style)
    {
        var info = new VmlStyleInfo();
        
        if (string.IsNullOrEmpty(style))
            return info;

        var parts = style.Split(';', StringSplitOptions.RemoveEmptyEntries);
        foreach (var part in parts)
        {
            var kv = part.Split(':');
            if (kv.Length != 2) continue;
            
            var key = kv[0].Trim().ToLowerInvariant();
            var value = kv[1].Trim();

            switch (key)
            {
                case "font-size":
                    if (value.EndsWith("pt"))
                    {
                        if (float.TryParse(value.Replace("pt", ""), out var size))
                            info.FontSize = size;
                    }
                    else if (float.TryParse(value, out var size2))
                    {
                        info.FontSize = size2 / 100f;
                    }
                    break;
                    
                case "font-weight":
                    info.IsBold = value == "bold" || value == "700";
                    break;
                    
                case "font-family":
                    info.FontName = value.Trim('"').Trim('\'');
                    break;
                    
                case "color":
                    info.TextColor = ParseColor(value);
                    break;
                    
                case "fill":
                    if (value != "none")
                    {
                        info.HasFill = true;
                        info.FillColor = ParseColor(value);
                    }
                    break;
            }
        }

        return info;
    }

    private static BaseColor? ParseColor(string colorValue)
    {
        if (string.IsNullOrEmpty(colorValue) || colorValue == "none")
            return null;

        try
        {
            if (colorValue.StartsWith("#"))
            {
                var hex = colorValue.Substring(1);
                if (hex.Length == 6)
                {
                    var r = Convert.ToByte(hex.Substring(0, 2), 16);
                    var g = Convert.ToByte(hex.Substring(2, 2), 16);
                    var b = Convert.ToByte(hex.Substring(4, 2), 16);
                    return new BaseColor(r, g, b);
                }
            }
            else
            {
                return colorValue.ToLowerInvariant() switch
                {
                    "red" => BaseColor.Red,
                    "blue" => BaseColor.Blue,
                    "green" => BaseColor.Green,
                    "black" => BaseColor.Black,
                    "white" => BaseColor.White,
                    "yellow" => BaseColor.Yellow,
                    "gray" => BaseColor.Gray,
                    "cyan" => BaseColor.Cyan,
                    "magenta" => BaseColor.Magenta,
                    "orange" => new BaseColor(255, 165, 0),
                    "pink" => new BaseColor(255, 192, 203),
                    _ => null
                };
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    private class VmlStyleInfo
    {
        public string? FontName { get; set; }
        public float FontSize { get; set; }
        public bool IsBold { get; set; }
        public BaseColor? TextColor { get; set; }
        public bool HasFill { get; set; }
        public BaseColor? FillColor { get; set; }
    }

    private static IElement? CreateShapeElement(OpenXmlElement shape)
    {
        var styleAttr = shape.GetAttribute("style", "");
        var style = styleAttr.Value ?? "";
        var info = ParseVmlStyle(style);

        // Determine size from style (e.g., width:100pt;height:50pt)
        float width = 50f, height = 50f;
        var parts = style.Split(';');
        foreach(var p in parts)
        {
            var kv = p.Split(':');
            if (kv.Length == 2)
            {
                var k = kv[0].Trim().ToLowerInvariant();
                var v = kv[1].Trim().ToLowerInvariant();
                if (k == "width" && v.EndsWith("pt")) float.TryParse(v.Replace("pt", ""), out width);
                else if (k == "height" && v.EndsWith("pt")) float.TryParse(v.Replace("pt", ""), out height);
            }
        }

        if (shape.LocalName == "rect" || shape.LocalName == "shape" || shape.LocalName == "oval")
        {
            var fill = shape.Descendants().FirstOrDefault(e => e.LocalName == "fill");
            if (fill != null)
            {
                var colorStr = fill.GetAttribute("color", "").Value;
                if (!string.IsNullOrEmpty(colorStr))
                {
                    var c = ParseColor(colorStr);
                    if (c != null)
                    {
                        info.HasFill = true;
                        info.FillColor = c;
                    }
                }
            }

            var stroke = shape.Descendants().FirstOrDefault(e => e.LocalName == "stroke");
            float borderWidth = 1f;
            BaseColor borderColor = info.TextColor ?? BaseColor.Black;

            if (stroke != null)
            {
                var weightStr = stroke.GetAttribute("weight", "").Value;
                if (!string.IsNullOrEmpty(weightStr))
                {
                    if (weightStr.EndsWith("pt")) float.TryParse(weightStr.Replace("pt", ""), out borderWidth);
                    else if (float.TryParse(weightStr, out var w)) borderWidth = w / 12700f; // EMU to pt
                }

                var colorStr = stroke.GetAttribute("color", "").Value;
                if (!string.IsNullOrEmpty(colorStr))
                {
                    var c = ParseColor(colorStr);
                    if (c != null) borderColor = c;
                }
            }

            var table = new PdfPTable(1) { TotalWidth = width, LockedWidth = true };
            var cell = new PdfPCell(new Phrase(" ")) 
            { 
                FixedHeight = height,
                BorderWidth = borderWidth,
                BorderColor = borderColor
            };
            if (info.HasFill && info.FillColor != null) cell.BackgroundColor = info.FillColor;
            table.AddCell(cell);
            return table;
        }

        return null;
    }
}
