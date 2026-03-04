using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.DocxToPdf.Helpers;
using Nedev.DocxToPdf.PdfEngine;
using iTextFont = Nedev.DocxToPdf.PdfEngine.Font;
using iTextList = Nedev.DocxToPdf.PdfEngine.List;
using iTextListItem = Nedev.DocxToPdf.PdfEngine.ListItem;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace Nedev.DocxToPdf.Converters;

/// <summary>
/// DOCX 列表转 PDF 列表
/// </summary>
public class ListConverter
{
    private readonly FontHelper _fontHelper;
    private readonly Styles? _styles;
    private readonly DocumentFormat.OpenXml.OpenXmlElement? _colorScheme;

    public ListConverter(FontHelper fontHelper, Styles? styles = null, DocumentFormat.OpenXml.OpenXmlElement? colorScheme = null)
    {
        _fontHelper = fontHelper;
        _styles = styles;
        _colorScheme = colorScheme;
    }

    /// <summary>
    /// 判断段落是否为列表项
    /// </summary>
    public static bool IsListItem(WParagraph paragraph)
    {
        var paraProps = paragraph.ParagraphProperties;
        return paraProps?.NumberingProperties?.NumberingId?.Val?.Value != null;
    }

    /// <summary>
    /// 获取列表项的编号 ID
    /// </summary>
    public static int? GetNumberingId(WParagraph paragraph)
    {
        return paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
    }

    /// <summary>
    /// 获取列表项的级别
    /// </summary>
    public static int GetListLevel(WParagraph paragraph)
    {
        return paragraph.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
    }

    /// <summary>
    /// 将连续的列表段落转为 iTextSharp List
    /// </summary>
    public iTextList ConvertListItems(
        IEnumerable<WParagraph> listParagraphs,
        Numbering? numbering,
        int numberingId)
    {
        var abstractNum = ResolveAbstractNum(numbering, numberingId, out var levelOverrides);

        var root = new iTextList(iTextList.UNORDERED)
        {
            IndentationLeft = 20f,
            SymbolIndent = 15f,
            Autoindent = true
        };

        var counters = new int[9];
        var started = new bool[9];

        var stack = new System.Collections.Generic.List<(int Level, iTextList List, iTextListItem? LastItem)>
        {
            (0, root, null)
        };

        foreach (var para in listParagraphs)
        {
            var level = Math.Clamp(GetListLevel(para), 0, 8);

            while (stack.Count - 1 > level)
            {
                stack.RemoveAt(stack.Count - 1);
            }

            while (stack.Count - 1 < level)
            {
                var parent = stack[^1];
                var parentItem = parent.LastItem;
                if (parentItem == null)
                {
                    parentItem = ConvertListItem(para);
                    parent.List.Add(parentItem);
                    stack[^1] = (parent.Level, parent.List, parentItem);
                }

                var nested = new iTextList(iTextList.UNORDERED)
                {
                    IndentationLeft = parent.List.IndentationLeft + 20f,
                    SymbolIndent = parent.List.SymbolIndent,
                    Autoindent = true
                };
                // nested 是 List，不能直接添加到 ListItem
                // 这里需要特殊处理嵌套列表
                stack.Add((stack.Count, nested, null));
            }

            var (lvlPr, startOverride) = ResolveLevelDefinition(abstractNum, level, levelOverrides);
            var (leftIndent, symbolIndent) = ResolveIndentation(lvlPr, level);

            var current = stack[^1];
            current.List.IndentationLeft = leftIndent;
            current.List.SymbolIndent = symbolIndent;

            var label = BuildListLabel(lvlPr, startOverride, level, counters, started);
            var item = ConvertListItem(para);
            ApplyListSymbol(item, label, para);

            current.List.Add(item);
            stack[^1] = (current.Level, current.List, item);
        }

        return root;
    }

    /// <summary>
    /// 将单个列表段落转为 ListItem
    /// </summary>
    public iTextListItem ConvertListItem(WParagraph paragraph)
    {
        var listItem = new iTextListItem();

        var paraProps = paragraph.ParagraphProperties;
        var styleId = paraProps?.ParagraphStyleId?.Val?.Value;
        var effectiveSpacing = paraProps?.SpacingBetweenLines ?? GetStyleSpacing(styleId);

        float actualFontSize = 12f;
        var firstRun = paragraph.Descendants<Run>().FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.InnerText));
        var runProps = firstRun?.RunProperties;
        var paraRunProps = paraProps?.GetFirstChild<ParagraphMarkRunProperties>();
        var styleFontSizeStr = GetStyleFontSize(styleId);
        var fontSizeStr = runProps?.FontSize?.Val?.Value
                          ?? paraRunProps?.GetFirstChild<FontSize>()?.Val?.Value
                          ?? styleFontSizeStr;
        if (fontSizeStr != null && float.TryParse(fontSizeStr, out var halfPt))
            actualFontSize = halfPt / 2f;

        var sampleFont = _fontHelper.GetFont(runProps, paraRunProps, actualFontSize);
        var baseLineHeight = GetBaseLineHeight(sampleFont, actualFontSize);
        ApplySpacing(listItem, effectiveSpacing, actualFontSize, baseLineHeight);

        var hasAnyChunk = false;
        foreach (var element in paragraph.ChildElements)
        {
            switch (element)
            {
                case Run run:
                    hasAnyChunk |= AddRunChunks(listItem, run, paraRunProps, actualFontSize);
                    break;
                case Hyperlink hyperlink:
                    var linkColor = StyleHelper.ResolveSchemeColor(_colorScheme, "hlink") ?? BaseColor.Blue;
                    foreach (var hlRun in hyperlink.Elements<Run>())
                        hasAnyChunk |= AddRunChunks(listItem, hlRun, paraRunProps, actualFontSize, linkColor, true);
                    break;
            }
        }

        if (!hasAnyChunk)
        {
            var emptyFont = _fontHelper.GetFont(null, paraRunProps, actualFontSize);
            listItem.Add(new Chunk(" ", emptyFont));
        }

        // 列表项缩进基于级别
        var level = GetListLevel(paragraph);
        if (level > 0)
        {
            listItem.IndentationLeft = level * 20f;
        }

        return listItem;
    }

    private void ApplyListSymbol(iTextListItem item, string label, WParagraph paragraph)
    {
        var paraProps = paragraph.ParagraphProperties;
        var styleId = paraProps?.ParagraphStyleId?.Val?.Value;

        float actualFontSize = 12f;
        var firstRun = paragraph.Descendants<Run>().FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.InnerText));
        var runProps = firstRun?.RunProperties;
        var paraRunProps = paraProps?.GetFirstChild<ParagraphMarkRunProperties>();
        var styleFontSizeStr = GetStyleFontSize(styleId);
        var fontSizeStr = runProps?.FontSize?.Val?.Value
                          ?? paraRunProps?.GetFirstChild<FontSize>()?.Val?.Value
                          ?? styleFontSizeStr;
        if (fontSizeStr != null && float.TryParse(fontSizeStr, out var halfPt))
            actualFontSize = halfPt / 2f;

        var symbolFont = _fontHelper.GetFont(runProps, paraRunProps, actualFontSize);
        item.ListSymbol = new Chunk(label, symbolFont);
    }

    private static (float LeftIndent, float SymbolIndent) ResolveIndentation(Level? level, int levelIndex)
    {
        var defaultLeft = 20f + levelIndex * 20f;
        var defaultSymbol = 15f;
        if (level?.PreviousParagraphProperties?.Indentation == null)
            return (defaultLeft, defaultSymbol);

        var ind = level.PreviousParagraphProperties.Indentation;
        var left = ind.Left?.Value != null ? StyleHelper.TwipsToPoints(ind.Left.Value) : defaultLeft;
        var hanging = ind.Hanging?.Value != null ? StyleHelper.TwipsToPoints(ind.Hanging.Value) : defaultSymbol;
        if (hanging <= 0 && ind.FirstLine?.Value != null)
        {
            var first = StyleHelper.TwipsToPoints(ind.FirstLine.Value);
            if (first < 0) hanging = -first;
        }

        if (left <= 0) left = defaultLeft;
        if (hanging <= 0) hanging = defaultSymbol;
        return (left, hanging);
    }

    private static AbstractNum? ResolveAbstractNum(Numbering? numbering, int numberingId, out Dictionary<int, int> levelOverrides)
    {
        levelOverrides = new Dictionary<int, int>();
        if (numbering == null) return null;

        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numberingId);

        if (numInstance == null) return null;

        foreach (var ov in numInstance.Elements<LevelOverride>())
        {
            var startNode = ov.GetFirstChild<StartOverrideNumberingValue>();
            if (ov.LevelIndex?.Value is int li && startNode?.Val?.Value is int start)
            {
                levelOverrides[li] = start;
            }
        }

        if (numInstance.AbstractNumId?.Val?.Value is not int abstractNumId) return null;

        return numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
    }

    private static (Level? Level, int? StartOverride) ResolveLevelDefinition(AbstractNum? abstractNum, int levelIndex, Dictionary<int, int> levelOverrides)
    {
        Level? lvl = null;
        if (abstractNum != null)
        {
            lvl = abstractNum.Elements<Level>()
                .FirstOrDefault(l => l.LevelIndex?.Value == levelIndex);
        }

        levelOverrides.TryGetValue(levelIndex, out var start);
        return (lvl, levelOverrides.ContainsKey(levelIndex) ? start : null);
    }

    private static string BuildListLabel(Level? level, int? startOverride, int levelIndex, int[] counters, bool[] started)
    {
        var format = level?.NumberingFormat?.Val?.Value;
        var lvlText = level?.LevelText?.Val?.Value;

        if (!started[levelIndex])
        {
            var start = startOverride ?? level?.StartNumberingValue?.Val?.Value ?? 1;
            counters[levelIndex] = start - 1;
            started[levelIndex] = true;
        }

        counters[levelIndex]++;
        for (var i = levelIndex + 1; i < counters.Length; i++)
        {
            started[i] = false;
            counters[i] = 0;
        }

        string FormatNumber(int idx, NumberFormatValues? fmt)
        {
            var n = counters[idx];
            if (fmt == null) return n.ToString();
            if (fmt.Equals(NumberFormatValues.UpperRoman)) return ToRoman(n).ToUpperInvariant();
            if (fmt.Equals(NumberFormatValues.LowerRoman)) return ToRoman(n).ToLowerInvariant();
            if (fmt.Equals(NumberFormatValues.UpperLetter)) return ToLetters(n).ToUpperInvariant();
            if (fmt.Equals(NumberFormatValues.LowerLetter)) return ToLetters(n).ToLowerInvariant();
            return n.ToString();
        }

        if (format != null && format.Equals(NumberFormatValues.Bullet))
        {
            var bullet = NormalizeBulletSymbol(string.IsNullOrWhiteSpace(lvlText) ? "•" : lvlText);
            return bullet.EndsWith(' ') ? bullet : bullet + " ";
        }

        if (!string.IsNullOrWhiteSpace(lvlText))
        {
            var s = lvlText;
            for (var i = 0; i <= levelIndex && i < 9; i++)
            {
                var token = "%" + (i + 1);
                var tokenFmt = levelIndex == i ? format : null;
                s = s.Replace(token, FormatNumber(i, tokenFmt), StringComparison.Ordinal);
            }
            return s.EndsWith(' ') ? s : s + " ";
        }

        return FormatNumber(levelIndex, format) + ". ";
    }

    private static string NormalizeBulletSymbol(string raw)
    {
        return raw switch
        {
            "" => "•",
            " " => "• ",
            "" => "➢",
            " " => "➢ ",
            "" => "■",
            " " => "■ ",
            "o" => "◦",
            "o " => "◦ ",
            _ => raw
        };
    }

    private static string ToRoman(int number)
    {
        if (number <= 0) return number.ToString();
        var map = new (int Value, string Sym)[]
        {
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I")
        };
        var n = number;
        var sb = new System.Text.StringBuilder();
        foreach (var (v, s) in map)
        {
            while (n >= v)
            {
                sb.Append(s);
                n -= v;
            }
        }
        return sb.ToString();
    }

    private static string ToLetters(int number)
    {
        if (number <= 0) return number.ToString();
        var n = number;
        var sb = new System.Text.StringBuilder();
        while (n > 0)
        {
            n--;
            sb.Insert(0, (char)('A' + (n % 26)));
            n /= 26;
        }
        return sb.ToString();
    }

    private static void ApplySpacing(iTextListItem listItem, SpacingBetweenLines? spacing, float fontSize, float baseLineHeight)
    {
        float minSpacing = Math.Max(8f, fontSize * 0.3f);
        float defaultMultiplier = fontSize > 16f ? 1.2f : 1.15f;

        if (spacing == null)
        {
            listItem.SetLeading(baseLineHeight * defaultMultiplier, 0);
            listItem.SpacingBefore = 0f;
            listItem.SpacingAfter = minSpacing;
            return;
        }

        if (spacing.Before?.Value is string beforeStr)
            listItem.SpacingBefore = StyleHelper.TwipsToPoints(beforeStr);
        else if (spacing.BeforeLines?.Value is int beforeLines && beforeLines > 0)
            listItem.SpacingBefore = fontSize * (beforeLines / 100f);

        if (spacing.After?.Value is string afterStr)
            listItem.SpacingAfter = StyleHelper.TwipsToPoints(afterStr);
        else if (spacing.AfterLines?.Value is int afterLines && afterLines > 0)
            listItem.SpacingAfter = fontSize * (afterLines / 100f);
        else
            listItem.SpacingAfter = minSpacing;

        if (spacing.Line?.Value is string lineStr && float.TryParse(lineStr, out var lineSpacing))
        {
            var lineRule = spacing.LineRule?.Value;
            float leading = lineSpacing / 20f;

            bool isExact = lineRule != null && lineRule.Equals(LineSpacingRuleValues.Exact);
            bool isAtLeast = lineRule != null && lineRule.Equals(LineSpacingRuleValues.AtLeast);

            if (isExact || isAtLeast)
            {
                var minLeading = baseLineHeight * defaultMultiplier;
                listItem.SetLeading(Math.Max(leading, minLeading), 0);
            }
            else
            {
                var lines = lineSpacing / 240f;
                var multiplier = lines > 0 ? lines : defaultMultiplier;
                listItem.SetLeading(baseLineHeight * multiplier, 0);
            }
        }
        else
        {
            listItem.SetLeading(baseLineHeight * defaultMultiplier, 0);
        }
    }

    private static float GetBaseLineHeight(iTextFont? font, float fontSize)
    {
        if (font == null || fontSize <= 0) return Math.Max(fontSize, 12f);
        // 简化实现：使用字体大小作为行高
        return Math.Max(fontSize * 1.2f, 12f);
    }

    private Style? GetStyleById(string? styleId)
    {
        if (string.IsNullOrWhiteSpace(styleId)) return null;
        return _styles?
            .Elements<Style>()
            .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
    }

    private T? GetFromStyleChain<T>(string? styleId, Func<Style, T?> selector) where T : class
    {
        var id = styleId;
        for (var i = 0; i < 20 && !string.IsNullOrWhiteSpace(id); i++)
        {
            var style = GetStyleById(id);
            if (style == null) return null;
            var v = selector(style);
            if (v != null) return v;
            id = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private SpacingBetweenLines? GetStyleSpacing(string? styleId)
    {
        return GetFromStyleChain(styleId, s => s.StyleParagraphProperties?.GetFirstChild<SpacingBetweenLines>());
    }

    private string? GetStyleFontSize(string? styleId)
    {
        var sz = GetFromStyleChain(styleId, s => s.StyleRunProperties?.GetFirstChild<FontSize>());
        return sz?.Val?.Value;
    }

    private bool AddRunChunks(iTextListItem listItem, Run run, ParagraphMarkRunProperties? paraRunProps, float actualFontSize, BaseColor? overrideColor = null, bool forceUnderline = false)
    {
        var runProps = run.RunProperties;
        var font = _fontHelper.GetFont(runProps, paraRunProps, actualFontSize);
        if (overrideColor != null)
            font = new iTextFont(font.Family, font.Size, font.Style, overrideColor);

        var hasAnyChunk = false;
        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case Text text:
                    var t = new Chunk(text.Text, font);
                    if (forceUnderline) t.SetUnderline(0.1f, -1f);
                    listItem.Add(t);
                    hasAnyChunk = true;
                    break;
                case TabChar:
                    var tab = new Chunk("    ", font);
                    if (forceUnderline) tab.SetUnderline(0.1f, -1f);
                    listItem.Add(tab);
                    hasAnyChunk = true;
                    break;
                case Break br:
                    listItem.Add(new Chunk("\n", font));
                    hasAnyChunk = true;
                    break;
            }
        }

        return hasAnyChunk;
    }

    /// <summary>
    /// 判断是否为有序列表
    /// </summary>
    private static bool IsOrderedList(Numbering? numbering, int numberingId)
    {
        if (numbering == null) return false;

        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numberingId);

        if (numInstance?.AbstractNumId?.Val?.Value is int abstractNumId)
        {
            var abstractNum = numbering.Elements<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);

            if (abstractNum != null)
            {
                var level0 = abstractNum.Elements<Level>()
                    .FirstOrDefault(l => l.LevelIndex?.Value == 0);

                if (level0?.NumberingFormat?.Val?.Value is NumberFormatValues format)
                {
                    if (format.Equals(NumberFormatValues.Bullet)) return false;

                    if (format.Equals(NumberFormatValues.Decimal) ||
                        format.Equals(NumberFormatValues.UpperLetter) ||
                        format.Equals(NumberFormatValues.LowerLetter) ||
                        format.Equals(NumberFormatValues.UpperRoman) ||
                        format.Equals(NumberFormatValues.LowerRoman))
                    {
                        return true;
                    }

                    return false;
                }
            }
        }

        return false;
    }
}
