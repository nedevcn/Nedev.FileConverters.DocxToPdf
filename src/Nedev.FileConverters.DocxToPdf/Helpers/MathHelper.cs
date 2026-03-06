using DocumentFormat.OpenXml;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextChunk = Nedev.FileConverters.DocxToPdf.PdfEngine.Chunk;
using iTextFont = Nedev.FileConverters.DocxToPdf.PdfEngine.Font;

namespace Nedev.FileConverters.DocxToPdf.Helpers;

public static class MathHelper
{
    public static List<Chunk> ExtractMathChunks(OpenXmlElement parent)
    {
        var chunks = new List<Chunk>();
        var omathElements = parent.Descendants().Where(e => e.LocalName == "oMath" && e.NamespaceUri.Contains("math")).ToList();

        foreach (var omath in omathElements)
        {
            var mathText = ExtractMathText(omath);
            if (!string.IsNullOrWhiteSpace(mathText))
            {
                var mathFont = FontFactory.GetFont("Cambria Math", 12);
                var chunk = new iTextChunk($"[公式: {mathText}]", mathFont);
                chunks.Add(chunk);
            }
        }

        return chunks;
    }

    private static string ExtractMathText(OpenXmlElement element)
    {
        if (element == null) return "";
        
        var parts = new List<string>();
        
        foreach (var child in element.ChildElements)
        {
            var localName = child.LocalName;
            
            if (localName == "t" || localName == "text")
            {
                var text = child.InnerText;
                if (!string.IsNullOrEmpty(text))
                    parts.Add(text);
            }
            else if (localName == "r" || localName == "run")
            {
                var runText = ExtractMathText(child);
                if (!string.IsNullOrEmpty(runText))
                    parts.Add(runText);
            }
            else if (localName == "f" || localName == "frac" || localName == "fraction")
            {
                var num = child.Descendants().FirstOrDefault(e => e.LocalName == "num" || e.LocalName == "numerator");
                var den = child.Descendants().FirstOrDefault(e => e.LocalName == "den" || e.LocalName == "denominator");
                var numText = num != null ? ExtractMathText(num) : "";
                var denText = den != null ? ExtractMathText(den) : "";
                if (!string.IsNullOrEmpty(numText) && !string.IsNullOrEmpty(denText))
                    parts.Add($"({numText})/({denText})");
            }
            else if (localName == "sSup" || localName == "sup" || localName == "superscript")
            {
                var baseElem = child.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "base");
                var supElem = child.Elements().FirstOrDefault(e => e.LocalName == "sup" || e.LocalName == "superScript");
                var baseText = baseElem != null ? ExtractMathText(baseElem) : "";
                var supText = supElem != null ? ExtractMathText(supElem) : "";
                if (!string.IsNullOrEmpty(baseText) && !string.IsNullOrEmpty(supText))
                    parts.Add($"{baseText}^{supText}");
            }
            else if (localName == "sSub" || localName == "sub" || localName == "subscript")
            {
                var baseElem = child.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "base");
                var subElem = child.Elements().FirstOrDefault(e => e.LocalName == "sub" || e.LocalName == "subScript");
                var baseText = baseElem != null ? ExtractMathText(baseElem) : "";
                var subText = subElem != null ? ExtractMathText(subElem) : "";
                if (!string.IsNullOrEmpty(baseText) && !string.IsNullOrEmpty(subText))
                    parts.Add($"{baseText}_{subText}");
            }
            else if (localName == "rad" || localName == "radical")
            {
                var radicand = child.Descendants().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "radicand");
                var degree = child.Descendants().FirstOrDefault(e => e.LocalName == "deg" || e.LocalName == "degree");
                var radicandText = radicand != null ? ExtractMathText(radicand) : "";
                var degreeText = degree != null ? ExtractMathText(degree) : "";
                if (!string.IsNullOrEmpty(radicandText))
                {
                    if (!string.IsNullOrEmpty(degreeText))
                        parts.Add($"{degreeText}√({radicandText})");
                    else
                        parts.Add($"√({radicandText})");
                }
            }
            else if (localName == "lim" || localName == "limit")
            {
                var baseElem = child.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "base");
                var limitElem = child.Elements().FirstOrDefault(e => e.LocalName == "lim" || e.LocalName == "limit");
                var baseText = baseElem != null ? ExtractMathText(baseElem) : "";
                var limitText = limitElem != null ? ExtractMathText(limitElem) : "";
                if (!string.IsNullOrEmpty(baseText) && !string.IsNullOrEmpty(limitText))
                    parts.Add($"lim_{limitText} {baseText}");
            }
            else if (localName == "sum" || localName == "product" || localName == "int" || localName == "integral")
            {
                var baseElem = child.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "base");
                var subElem = child.Elements().FirstOrDefault(e => e.LocalName == "sub" || e.LocalName == "subScript");
                var supElem = child.Elements().FirstOrDefault(e => e.LocalName == "sup" || e.LocalName == "superScript");
                var baseText = baseElem != null ? ExtractMathText(baseElem) : "";
                var subText = subElem != null ? ExtractMathText(subElem) : "";
                var supText = supElem != null ? ExtractMathText(supElem) : "";
                var op = localName == "sum" ? "∑" : localName == "product" ? "∏" : "∫";
                parts.Add($"{op}_{{{subText}}}^{{{supText}}}{baseText}");
            }
            else if (localName == "func" || localName == "function")
            {
                var fname = child.Descendants().FirstOrDefault(e => e.LocalName == "fName" || e.LocalName == "functionName");
                var fe = child.Descendants().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "element");
                var fnameText = fname != null ? ExtractMathText(fname) : "";
                var feText = fe != null ? ExtractMathText(fe) : "";
                if (!string.IsNullOrEmpty(fnameText) && !string.IsNullOrEmpty(feText))
                    parts.Add($"{fnameText}({feText})");
            }
            else if (localName == "bar" || localName == "overline")
            {
                var barElem = child.Descendants().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "element");
                var barText = barElem != null ? ExtractMathText(barElem) : "";
                if (!string.IsNullOrEmpty(barText))
                    parts.Add($"({barText})¯");
            }
            else if (localName == "vec" || localName == "vector")
            {
                var vecElem = child.Descendants().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "element");
                var vecText = vecElem != null ? ExtractMathText(vecElem) : "";
                if (!string.IsNullOrEmpty(vecText))
                    parts.Add($"({vecText})⃗");
            }
            else
            {
                var childText = ExtractMathText(child);
                if (!string.IsNullOrEmpty(childText))
                    parts.Add(childText);
            }
        }
        
        if (parts.Count == 0 && !string.IsNullOrEmpty(element.InnerText))
            return element.InnerText;
            
        return string.Join("", parts);
    }
}
