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
                // Recursive call for internal elements (like <m:t>)
                parts.Add(ExtractMathText(child));
            }
            else if (localName == "f" || localName == "frac" || localName == "fraction")
            {
                var num = child.Elements().FirstOrDefault(e => e.LocalName == "num" || e.LocalName == "numerator");
                var den = child.Elements().FirstOrDefault(e => e.LocalName == "den" || e.LocalName == "denominator");
                var numText = num != null ? ExtractMathText(num) : "";
                var denText = den != null ? ExtractMathText(den) : "";
                parts.Add($"({numText})/({denText})");
            }
            else if (localName == "sSup" || localName == "sup" || localName == "superscript")
            {
                var baseElem = child.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "base");
                var supElem = child.Elements().FirstOrDefault(e => e.LocalName == "sup" || e.LocalName == "superScript");
                var baseText = baseElem != null ? ExtractMathText(baseElem) : "";
                var supText = supElem != null ? ExtractMathText(supElem) : "";
                parts.Add($"{baseText}^{supText}");
            }
            else if (localName == "sSub" || localName == "sub" || localName == "subscript")
            {
                var baseElem = child.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "base");
                var subElem = child.Elements().FirstOrDefault(e => e.LocalName == "sub" || e.LocalName == "subScript");
                var baseText = baseElem != null ? ExtractMathText(baseElem) : "";
                var subText = subElem != null ? ExtractMathText(subElem) : "";
                parts.Add($"{baseText}_{subText}");
            }
            else if (localName == "sSubSup" || localName == "subsup")
            {
                var baseElem = child.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "base");
                var subElem = child.Elements().FirstOrDefault(e => e.LocalName == "sub" || e.LocalName == "subScript");
                var supElem = child.Elements().FirstOrDefault(e => e.LocalName == "sup" || e.LocalName == "superScript");
                var baseText = baseElem != null ? ExtractMathText(baseElem) : "";
                var subText = subElem != null ? ExtractMathText(subElem) : "";
                var supText = supElem != null ? ExtractMathText(supElem) : "";
                parts.Add($"{baseText}_{{{subText}}}^{{{supText}}}");
            }
            else if (localName == "rad" || localName == "radical")
            {
                var radicand = child.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "radicand");
                var degree = child.Elements().FirstOrDefault(e => e.LocalName == "deg" || e.LocalName == "degree");
                var radicandText = radicand != null ? ExtractMathText(radicand) : "";
                var degreeText = degree != null ? ExtractMathText(degree) : "";
                if (!string.IsNullOrEmpty(degreeText))
                    parts.Add($"{degreeText}√({radicandText})");
                else
                    parts.Add($"√({radicandText})");
            }
            else if (localName == "lim" || localName == "limit" || localName == "limLow" || localName == "limUp")
            {
                var baseElem = child.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "base");
                var limitElem = child.Elements().FirstOrDefault(e => e.LocalName == "lim" || e.LocalName == "limit" || e.LocalName == "sub" || e.LocalName == "sup");
                var baseText = baseElem != null ? ExtractMathText(baseElem) : "";
                var limitText = limitElem != null ? ExtractMathText(limitElem) : "";
                var op = localName.Contains("Low") ? "lim_" : localName.Contains("Up") ? "lim^" : "lim_";
                parts.Add($"{op}{{{limitText}}} {baseText}");
            }
            else if (localName == "sum" || localName == "product" || localName == "int" || localName == "integral" || localName == "nary")
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
                var fname = child.Elements().FirstOrDefault(e => e.LocalName == "fName" || e.LocalName == "functionName");
                var fe = child.Elements().FirstOrDefault(e => e.LocalName == "e" || e.LocalName == "element");
                var fnameText = fname != null ? ExtractMathText(fname) : "";
                var feText = fe != null ? ExtractMathText(fe) : "";
                parts.Add($"{fnameText}({feText})");
            }
            else if (localName == "d" || localName == "delimiter")
            {
                var open = child.Elements().FirstOrDefault(e => e.LocalName == "begChr")?.InnerText ?? "(";
                var close = child.Elements().FirstOrDefault(e => e.LocalName == "endChr")?.InnerText ?? ")";
                var content = child.Elements().FirstOrDefault(e => e.LocalName == "e");
                parts.Add($"{open}{(content != null ? ExtractMathText(content) : "")}{close}");
            }
            else if (localName == "m" || localName == "matrix")
            {
                parts.Add("[Matrix]");
            }
            else if (localName == "groupChr")
            {
                var eElem = child.Elements().FirstOrDefault(e => e.LocalName == "e");
                var chrElem = child.Elements().FirstOrDefault(e => e.LocalName == "chr")?.InnerText ?? "¯";
                parts.Add($"({(eElem != null ? ExtractMathText(eElem) : "")}){chrElem}");
            }
            else
            {
                // Fallback: just proceed with children, but avoid innerText here to prevent duplication
                parts.Add(ExtractMathText(child));
            }
        }
        
        if (parts.Count == 0 && !element.HasChildren && !string.IsNullOrEmpty(element.InnerText))
            return element.InnerText;
            
        return string.Join("", parts);
    }
}
