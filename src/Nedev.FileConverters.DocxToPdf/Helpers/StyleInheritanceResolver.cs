using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.PdfEngine;

namespace Nedev.FileConverters.DocxToPdf.Helpers;

/// <summary>
/// 样式继承解析器 - 提供完整的 Word 样式继承链解析
/// </summary>
public class StyleInheritanceResolver
{
    private readonly Styles? _styles;
    private readonly Dictionary<string, ResolvedStyle> _styleCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Style> _styleMap = new(StringComparer.OrdinalIgnoreCase);

    public StyleInheritanceResolver(Styles? styles)
    {
        _styles = styles;
        BuildStyleMap();
    }

    private void BuildStyleMap()
    {
        if (_styles == null) return;

        foreach (var style in _styles.Elements<Style>())
        {
            var styleId = style.StyleId?.Value;
            if (!string.IsNullOrWhiteSpace(styleId))
            {
                _styleMap[styleId] = style;
            }
        }
    }

    /// <summary>
    /// 获取解析后的样式（包含继承链的所有属性）
    /// </summary>
    public ResolvedStyle ResolveStyle(string? styleId)
    {
        if (string.IsNullOrWhiteSpace(styleId))
            return new ResolvedStyle();

        if (_styleCache.TryGetValue(styleId, out var cached))
            return cached;

        var resolved = BuildResolvedStyle(styleId);
        _styleCache[styleId] = resolved;
        return resolved;
    }

    /// <summary>
    /// 构建解析后的样式
    /// </summary>
    private ResolvedStyle BuildResolvedStyle(string? styleId)
    {
        var result = new ResolvedStyle();
        if (string.IsNullOrWhiteSpace(styleId)) return result;

        // 收集样式链（从基础样式到具体样式）
        var styleChain = new List<Style>();
        var currentId = styleId;
        var depth = 0;
        const int maxDepth = 20; // 防止循环继承

        while (!string.IsNullOrWhiteSpace(currentId) && depth < maxDepth)
        {
            if (_styleMap.TryGetValue(currentId, out var style))
            {
                styleChain.Add(style);
                currentId = style.BasedOn?.Val?.Value;
            }
            else
            {
                break;
            }
            depth++;
        }

        // 从基础样式到具体样式依次应用属性（后面的覆盖前面的）
        for (int i = styleChain.Count - 1; i >= 0; i--)
        {
            ApplyStyleProperties(result, styleChain[i]);
        }

        result.StyleId = styleId;
        return result;
    }

    /// <summary>
    /// 应用单个样式的属性到解析结果
    /// </summary>
    private void ApplyStyleProperties(ResolvedStyle target, Style source)
    {
        var paraProps = source.StyleParagraphProperties;
        var runProps = source.StyleRunProperties;

        // 段落属性
        if (paraProps != null)
        {
            // 对齐方式
            if (paraProps.Justification?.Val?.Value is var justification)
                target.Justification = justification;

            // 间距
            if (paraProps.GetFirstChild<SpacingBetweenLines>() is var spacing)
                target.Spacing = spacing;

            // 缩进
            if (paraProps.GetFirstChild<Indentation>() is var indentation)
                target.Indentation = indentation;

            // 边框
            if (paraProps.GetFirstChild<ParagraphBorders>() is var borders)
                target.ParagraphBorders = borders;

            // 底纹
            if (paraProps.GetFirstChild<Shading>() is var shading)
                target.Shading = shading;

            // 分页控制
            if (paraProps.GetFirstChild<KeepLines>() is var keepLines)
                target.KeepLines = keepLines.Val == null || keepLines.Val.Value;

            if (paraProps.GetFirstChild<KeepNext>() is var keepNext)
                target.KeepNext = keepNext.Val == null || keepNext.Val.Value;

            if (paraProps.GetFirstChild<PageBreakBefore>() is var pageBreakBefore)
                target.PageBreakBefore = pageBreakBefore.Val != null && pageBreakBefore.Val.Value;

            if (paraProps.GetFirstChild<WidowControl>()?.Val?.Value is bool widowVal)
                target.WidowControl = widowVal ? 1 : 0;

            // 大纲级别
            if (paraProps.GetFirstChild<OutlineLevel>()?.Val?.Value is var outlineLevel)
                target.OutlineLevel = outlineLevel;
        }

        // 文本属性
        if (runProps != null)
        {
            // 字体大小
            if (runProps.GetFirstChild<FontSize>()?.Val?.Value is var fontSize)
                target.FontSize = fontSize;

            // 复杂脚本字体大小
            if (runProps.GetFirstChild<FontSizeComplexScript>()?.Val?.Value is var fontSizeCs)
                target.FontSizeComplexScript = fontSizeCs;

            // 字体名称
            if (runProps.GetFirstChild<RunFonts>()?.Ascii?.Value is var asciiFont)
                target.FontAscii = asciiFont;

            if (runProps.GetFirstChild<RunFonts>()?.EastAsia?.Value is var eastAsiaFont)
                target.FontEastAsia = eastAsiaFont;

            if (runProps.GetFirstChild<RunFonts>()?.HighAnsi?.Value is var highAnsiFont)
                target.FontHighAnsi = highAnsiFont;

            // 加粗
            if (runProps.GetFirstChild<Bold>()?.Val?.Value is var bold)
                target.Bold = bold;
            else if (runProps.GetFirstChild<Bold>() != null)
                target.Bold = true;

            // 倾斜
            if (runProps.GetFirstChild<Italic>()?.Val?.Value is var italic)
                target.Italic = italic;
            else if (runProps.GetFirstChild<Italic>() != null)
                target.Italic = true;

            // 下划线
            if (runProps.GetFirstChild<Underline>()?.Val?.Value is var underline)
                target.Underline = underline;

            // 删除线
            if (runProps.GetFirstChild<Strike>()?.Val?.Value is var strike)
                target.Strike = strike;
            else if (runProps.GetFirstChild<Strike>() != null)
                target.Strike = true;

            // 双删除线
            if (runProps.GetFirstChild<DoubleStrike>()?.Val?.Value is var doubleStrike)
                target.DoubleStrike = doubleStrike;
            else if (runProps.GetFirstChild<DoubleStrike>() != null)
                target.DoubleStrike = true;

            // 颜色
            if (runProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>()?.Val?.Value is var color)
                target.Color = color;

            // 高亮
            if (runProps.GetFirstChild<Highlight>()?.Val?.Value is var highlight)
                target.Highlight = highlight;

            // 上标/下标
            if (runProps.GetFirstChild<VerticalTextAlignment>()?.Val?.Value is var vertAlign)
                target.VerticalAlignment = vertAlign;

            // 字符间距
            var spacingVal = runProps.GetFirstChild<Spacing>()?.Val?.Value;
            if (spacingVal != null)
                target.CharacterSpacing = spacingVal.ToString();

            // 字间距调整
            var kernVal = runProps.GetFirstChild<Kern>()?.Val?.Value;
            if (kernVal != null)
                target.Kern = kernVal.ToString();

            // 缩放 (Scale 类在当前版本中不可用)
            // if (runProps.GetFirstChild<Scale>()?.Val?.Value is var scale)
            //     target.Scale = scale;

            // 位置
            if (runProps.GetFirstChild<Position>()?.Val?.Value is var position)
                target.Position = position;

            // 小写大写字母
            if (runProps.GetFirstChild<SmallCaps>()?.Val?.Value is var smallCaps)
                target.SmallCaps = smallCaps;
            else if (runProps.GetFirstChild<SmallCaps>() != null)
                target.SmallCaps = true;

            // 全大写
            if (runProps.GetFirstChild<Caps>()?.Val?.Value is var caps)
                target.Caps = caps;
            else if (runProps.GetFirstChild<Caps>() != null)
                target.Caps = true;
        }

        // 样式类型
        if (source.Type?.Value is var styleType)
            target.StyleType = styleType;

        // 默认样式标记
        if (source.Default?.HasValue == true)
            target.IsDefault = source.Default.Value;
    }

    /// <summary>
    /// 合并段落直接属性与继承样式
    /// </summary>
    public ResolvedStyle MergeWithDirectProperties(ResolvedStyle inherited, ParagraphProperties? directProps)
    {
        var result = inherited.Clone();

        if (directProps == null) return result;

        // 直接属性优先于继承属性
        if (directProps.Justification?.Val?.Value is var justification)
            result.Justification = justification;

        if (directProps.GetFirstChild<SpacingBetweenLines>() is var spacing)
            result.Spacing = spacing;

        if (directProps.GetFirstChild<Indentation>() is var indentation)
            result.Indentation = indentation;

        if (directProps.GetFirstChild<ParagraphBorders>() is var borders)
            result.ParagraphBorders = borders;

        if (directProps.GetFirstChild<Shading>() is var shading)
            result.Shading = shading;

        if (directProps.GetFirstChild<KeepLines>() is var keepLines)
            result.KeepLines = keepLines.Val == null || keepLines.Val.Value;

        if (directProps.GetFirstChild<KeepNext>() is var keepNext)
            result.KeepNext = keepNext.Val == null || keepNext.Val.Value;

        if (directProps.GetFirstChild<PageBreakBefore>() is var pageBreakBefore)
            result.PageBreakBefore = pageBreakBefore.Val != null && pageBreakBefore.Val.Value;

        if (directProps.GetFirstChild<OutlineLevel>()?.Val?.Value is var outlineLevel)
            result.OutlineLevel = outlineLevel;

        return result;
    }

    /// <summary>
    /// 合并 Run 直接属性与继承样式
    /// </summary>
    public ResolvedStyle MergeWithDirectProperties(ResolvedStyle inherited, RunProperties? directProps)
    {
        var result = inherited.Clone();

        if (directProps == null) return result;

        // 直接属性优先于继承属性
        if (directProps.GetFirstChild<FontSize>()?.Val?.Value is var fontSize)
            result.FontSize = fontSize;

        if (directProps.GetFirstChild<FontSizeComplexScript>()?.Val?.Value is var fontSizeCs)
            result.FontSizeComplexScript = fontSizeCs;

        if (directProps.GetFirstChild<RunFonts>()?.Ascii?.Value is var asciiFont)
            result.FontAscii = asciiFont;

        if (directProps.GetFirstChild<RunFonts>()?.EastAsia?.Value is var eastAsiaFont)
            result.FontEastAsia = eastAsiaFont;

        if (directProps.GetFirstChild<RunFonts>()?.HighAnsi?.Value is var highAnsiFont)
            result.FontHighAnsi = highAnsiFont;

        if (directProps.GetFirstChild<Bold>()?.Val?.Value is var bold)
            result.Bold = bold;
        else if (directProps.GetFirstChild<Bold>() != null)
            result.Bold = true;

        if (directProps.GetFirstChild<Italic>()?.Val?.Value is var italic)
            result.Italic = italic;
        else if (directProps.GetFirstChild<Italic>() != null)
            result.Italic = true;

        if (directProps.GetFirstChild<Underline>()?.Val?.Value is var underline)
            result.Underline = underline;

        if (directProps.GetFirstChild<Strike>()?.Val?.Value is var strike)
            result.Strike = strike;
        else if (directProps.GetFirstChild<Strike>() != null)
            result.Strike = true;

        if (directProps.GetFirstChild<DoubleStrike>()?.Val?.Value is var doubleStrike)
            result.DoubleStrike = doubleStrike;
        else if (directProps.GetFirstChild<DoubleStrike>() != null)
            result.DoubleStrike = true;

        if (directProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>()?.Val?.Value is var color)
            result.Color = color;

        if (directProps.GetFirstChild<Highlight>()?.Val?.Value is var highlight)
            result.Highlight = highlight;

        if (directProps.GetFirstChild<VerticalTextAlignment>()?.Val?.Value is var vertAlign)
            result.VerticalAlignment = vertAlign;

        var spacing = directProps.GetFirstChild<Spacing>()?.Val?.Value;
        if (spacing != null)
            result.CharacterSpacing = spacing.ToString();

        if (directProps.GetFirstChild<SmallCaps>()?.Val?.Value is var smallCaps)
            result.SmallCaps = smallCaps;
        else if (directProps.GetFirstChild<SmallCaps>() != null)
            result.SmallCaps = true;

        if (directProps.GetFirstChild<Caps>()?.Val?.Value is var caps)
            result.Caps = caps;
        else if (directProps.GetFirstChild<Caps>() != null)
            result.Caps = true;

        return result;
    }
}

/// <summary>
/// 解析后的样式 - 包含从继承链合并的所有属性
/// </summary>
public class ResolvedStyle
{
    public string? StyleId { get; set; }
    public StyleValues? StyleType { get; set; }
    public bool IsDefault { get; set; }

    // 段落属性
    public JustificationValues? Justification { get; set; }
    public SpacingBetweenLines? Spacing { get; set; }
    public Indentation? Indentation { get; set; }
    public ParagraphBorders? ParagraphBorders { get; set; }
    public Shading? Shading { get; set; }
    public bool KeepLines { get; set; }
    public bool KeepNext { get; set; }
    public bool PageBreakBefore { get; set; }
    public int WidowControl { get; set; }
    public int? OutlineLevel { get; set; }

    // 文本属性
    public string? FontSize { get; set; }
    public string? FontSizeComplexScript { get; set; }
    public string? FontAscii { get; set; }
    public string? FontEastAsia { get; set; }
    public string? FontHighAnsi { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public UnderlineValues? Underline { get; set; }
    public bool? Strike { get; set; }
    public bool? DoubleStrike { get; set; }
    public string? Color { get; set; }
    public HighlightColorValues? Highlight { get; set; }
    public VerticalPositionValues? VerticalAlignment { get; set; }
    public string? CharacterSpacing { get; set; }
    public string? Kern { get; set; }
    public string? Scale { get; set; }
    public string? Position { get; set; }
    public bool? SmallCaps { get; set; }
    public bool? Caps { get; set; }

    /// <summary>
    /// 克隆当前样式
    /// </summary>
    public ResolvedStyle Clone()
    {
        return new ResolvedStyle
        {
            StyleId = StyleId,
            StyleType = StyleType,
            IsDefault = IsDefault,
            Justification = Justification,
            Spacing = Spacing,
            Indentation = Indentation,
            ParagraphBorders = ParagraphBorders,
            Shading = Shading,
            KeepLines = KeepLines,
            KeepNext = KeepNext,
            PageBreakBefore = PageBreakBefore,
            WidowControl = WidowControl,
            OutlineLevel = OutlineLevel,
            FontSize = FontSize,
            FontSizeComplexScript = FontSizeComplexScript,
            FontAscii = FontAscii,
            FontEastAsia = FontEastAsia,
            FontHighAnsi = FontHighAnsi,
            Bold = Bold,
            Italic = Italic,
            Underline = Underline,
            Strike = Strike,
            DoubleStrike = DoubleStrike,
            Color = Color,
            Highlight = Highlight,
            VerticalAlignment = VerticalAlignment,
            CharacterSpacing = CharacterSpacing,
            Kern = Kern,
            Scale = Scale,
            Position = Position,
            SmallCaps = SmallCaps,
            Caps = Caps
        };
    }

    /// <summary>
    /// 获取有效的字体大小（以磅为单位）
    /// </summary>
    public float? GetFontSizeInPoints()
    {
        var sizeStr = FontSize ?? FontSizeComplexScript;
        if (string.IsNullOrWhiteSpace(sizeStr)) return null;

        if (float.TryParse(sizeStr, out var halfPoints))
        {
            return halfPoints / 2f;
        }

        return null;
    }

    /// <summary>
    /// 获取首选字体名称
    /// </summary>
    public string? GetPreferredFontName()
    {
        return FontEastAsia ?? FontAscii ?? FontHighAnsi;
    }

    /// <summary>
    /// 判断是否为大纲级别（标题）
    /// </summary>
    public bool IsHeading()
    {
        return OutlineLevel.HasValue && OutlineLevel.Value >= 1 && OutlineLevel.Value <= 9;
    }

    /// <summary>
    /// 获取大纲级别
    /// </summary>
    public int? GetHeadingLevel()
    {
        if (OutlineLevel.HasValue && OutlineLevel.Value >= 1 && OutlineLevel.Value <= 9)
            return OutlineLevel.Value;

        // 从样式 ID 推断
        if (!string.IsNullOrWhiteSpace(StyleId))
        {
            var lower = StyleId.ToLowerInvariant();
            if (lower.StartsWith("heading"))
            {
                var numPart = lower.Substring("heading".Length);
                if (int.TryParse(numPart, out var level) && level >= 1 && level <= 9)
                    return level;
            }
        }

        return null;
    }
}
