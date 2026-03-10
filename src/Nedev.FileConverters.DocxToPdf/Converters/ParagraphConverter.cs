using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextFont = Nedev.FileConverters.DocxToPdf.PdfEngine.Font;
using iTextParagraph = Nedev.FileConverters.DocxToPdf.PdfEngine.Paragraph;
using iTextChunk = Nedev.FileConverters.DocxToPdf.PdfEngine.Chunk;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace Nedev.FileConverters.DocxToPdf.Converters;

/// <summary>
/// DOCX 段落地转 PDF 段落
/// </summary>
public class ParagraphConverter
{
    // debug output buffer (used by tests to inspect parsing behaviour)
    public static System.Text.StringBuilder DebugLog { get; } = new System.Text.StringBuilder();

    private readonly FontHelper _fontHelper;
    private readonly Styles? _styles;
    private readonly OpenXmlElement? _colorScheme;
    private readonly IReadOnlyDictionary<string, string>? _hyperlinkTargets;
    private readonly IReadOnlyDictionary<int, int>? _footnoteNumberById;
    private readonly IReadOnlyDictionary<int, int>? _endnoteNumberById;
    private readonly MainDocumentPart? _mainPart;

    /// <summary>页眉页脚渲染时提供当前页/总页数，用于 PAGE/NUMPAGES 字段</summary>
    public Func<(int Current, int Total)?>? PageNumberProvider { get; set; }

    /// <summary>遇到脚注/尾注引用时记录 ID，用于文末输出内容</summary>
    public ICollection<int>? FootnoteIdsEncountered { get; set; }
    public ICollection<int>? EndnoteIdsEncountered { get; set; }

    /// <summary>书签跟踪器，用于添加标题书签</summary>
    public object? BookmarkTracker { get; set; }

    /// <summary>标题真正落版时回调，供目录等后处理记录实际页码</summary>
    public Action<string, string, int, int>? HeadingRendered { get; set; }

    /// <summary>字段解析器：输入完整指令字符串（如 \"DATE \\@ yyyy-MM-dd\"），返回要显示的文本</summary>
    public Func<string, string?>? FieldResolver { get; set; }

    /// <summary>完整上下文，支持超链接、脚注尾注</summary>
    public ParagraphConverter(
        FontHelper fontHelper,
        Styles? styles = null,
        OpenXmlElement? colorScheme = null,
        IReadOnlyDictionary<string, string>? hyperlinkTargets = null,
        IReadOnlyDictionary<int, int>? footnoteNumberById = null,
        IReadOnlyDictionary<int, int>? endnoteNumberById = null,
        MainDocumentPart? mainPart = null)
    {
        _fontHelper = fontHelper;
        _styles = styles;
        _colorScheme = colorScheme;
        _hyperlinkTargets = hyperlinkTargets;
        _footnoteNumberById = footnoteNumberById;
        _endnoteNumberById = endnoteNumberById;
        _mainPart = mainPart;
    }

    /// <summary>
    /// 将 Run 转为 Chunk 列表
    /// </summary>
    private System.Collections.Generic.List<iTextChunk> ConvertRun(Run run, bool isHeading, float? headingSize, iTextFont font)
    {
        var chunks = new System.Collections.Generic.List<iTextChunk>();
        var runProps = run.RunProperties;

        // 上标/下标：调整字号与基线
        var vertAlign = runProps?.VerticalTextAlignment?.Val?.Value;
        var (textRise, runFont) = ApplyVertAlign(vertAlign, font);

        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case Text text:
                    var chunk = new iTextChunk(text.Text, runFont);
                    if (textRise != 0) chunk.SetTextRise(textRise);

                    if (runProps?.Shading != null)
                    {
                        var bg = StyleHelper.ResolveShadingFill(_colorScheme, runProps.Shading);
                        if (bg != null) chunk.SetBackground(bg);
                    }

                    // 背景高亮
                    if (runProps?.Highlight?.Val?.Value is HighlightColorValues highlight && highlight != HighlightColorValues.None)
                    {
                        var bgColor = HighlightToBaseColor(highlight);
                        if (bgColor != null)
                            chunk.SetBackground(bgColor);
                    }

                    if (ShouldApplySyntheticBold(runFont))
                    {
                        var size = runFont.Size > 0 ? runFont.Size : 12f;
                        var strokeWidth = Math.Clamp(size * 0.02f, 0.18f, 0.35f);
                        chunk.SetTextRenderMode(PdfContentByte.TEXT_RENDER_MODE_FILL_STROKE, strokeWidth, runFont.Color ?? BaseColor.Black);
                    }

                    chunks.Add(chunk);
                    break;

                case Break br:
                    var brFont = isHeading
                        ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                        : _fontHelper.GetFont(runProps, null, null, false);
                    if (br.Type?.Value == BreakValues.Page)
                    {
                        // 发送分页信号
                        chunks.Add(new iTextChunk("NEXTPAGE_SIGNAL", brFont));
                    }
                    else
                    {
                        chunks.Add(new iTextChunk(Environment.NewLine, brFont));
                    }
                    break;
                    
                case LastRenderedPageBreak:
                    // Word 的软分页标记，保持与前文处理逻辑一致，现在空白段落 bug 修复后不会多出上百页空白了
                    // 实际上这是Word用于缓存分页位置的标记，不应在我们的自定义PDF引擎中强制换页，这会导致双重换页（产生空白页）。
                    // chunks.Add(new iTextChunk("NEXTPAGE_SIGNAL", font));
                    break;

                case TabChar:
                    chunks.Add(new iTextChunk("    ", _fontHelper.GetFont(runProps)));
                    break;
            }
        }

        return chunks;
    }

    private static (float TextRise, iTextFont Font) ApplyVertAlign(VerticalPositionValues? vertAlign, iTextFont font)
    {
        if (vertAlign == null) return (0, font);
        var size = font.Size > 0 ? font.Size : 12f;
        if (vertAlign == VerticalPositionValues.Superscript)
        {
            var smallFont = new iTextFont(font.Family, size * 0.6f, font.Style, font.Color);
            return (size * 0.4f, smallFont);
        }
        if (vertAlign == VerticalPositionValues.Subscript)
        {
            var smallFont = new iTextFont(font.Family, size * 0.6f, font.Style, font.Color);
            return (-size * 0.2f, smallFont);
        }
        return (0, font);
    }

    /// <summary>
    /// 将 DOCX Paragraph 转为 iTextSharp 元素列表
    /// </summary>
    public System.Collections.Generic.List<IElement> Convert(WParagraph docxParagraph, bool forceBold = false)
    {
        var elements = new System.Collections.Generic.List<IElement>();
        var pdfParagraph = new iTextParagraph();

        // 段落属性
        var paraProps = docxParagraph.ParagraphProperties;
        var styleId = paraProps?.ParagraphStyleId?.Val?.Value;

        // 对齐方式
        JustificationValues? justification = paraProps?.Justification?.Val?.Value ?? GetStyleJustification(styleId);
        if (justification != null)
        {
            pdfParagraph.Alignment = StyleHelper.ToiTextAlignment(justification);
        }

        // 判断是否为标题
        var isHeading = StyleHelper.IsHeadingStyle(styleId);
        float? headingSize = isHeading ? StyleHelper.GetHeadingFontSize(styleId) : null;
        int? headingLevel = isHeading ? GetHeadingLevel(styleId) : null;

        var effectiveSpacing = paraProps?.SpacingBetweenLines ?? GetStyleSpacing(styleId);
        var effectiveIndentation = paraProps?.Indentation ?? GetStyleIndentation(styleId);

        // 从段落的第一个 Run 获取实际字号
        float actualFontSize = 12f;
        var firstRun = docxParagraph.Descendants<Run>().FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.InnerText));
        var runProps = firstRun?.RunProperties;
        var paraRunProps = paraProps?.GetFirstChild<ParagraphMarkRunProperties>();
        
        var styleFontSizeStr = GetStyleFontSize(styleId);
        var fontSizeStr = runProps?.FontSize?.Val?.Value
                          ?? paraRunProps?.GetFirstChild<FontSize>()?.Val?.Value
                          ?? styleFontSizeStr;

        if (fontSizeStr != null && float.TryParse(fontSizeStr, out var halfPt))
        {
            actualFontSize = halfPt / 2f;
        }
        else if (headingSize.HasValue && headingSize.Value > 0)
        {
            actualFontSize = headingSize.Value;
        }

        var sampleFont = isHeading
            ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
            : _fontHelper.GetFont(runProps, paraRunProps, actualFontSize, forceBold);

        var baseLineHeight = GetBaseLineHeight(sampleFont, actualFontSize);
        
        // 段落间距
        SetParagraphSpacing(pdfParagraph, effectiveSpacing, docxParagraph.Parent is TableCell, isHeading, actualFontSize, baseLineHeight);

        // 段落缩进
        SetParagraphIndentation(pdfParagraph, effectiveIndentation);

        // 段落控制：KeepWithNext、KeepLinesTogether
        ApplyParagraphKeepOptions(pdfParagraph, paraProps, styleId);

        // 当正在处理由字段（SimpleField 或复杂字段）产生的超链接时，临时保存 URL，
        // 以便 AppendInline 在生成 Chunk 时添加 anchor 属性。
        string? currentFieldHref = null;

        // 复杂字段解析状态
        string? currentComplexInstr = null;
        bool inComplexField = false;
        bool complexFieldHasSeparator = false;

        var hasContent = false;

        void AppendInline(OpenXmlElement element)
        {
            switch (element)
            {
                case DocumentFormat.OpenXml.Wordprocessing.FieldChar fieldChar:
                    // 传统复杂字段结构由 FieldChar 开始/分隔/结束
                    if (fieldChar.FieldCharType != null)
                    {
                        var type = fieldChar.FieldCharType.Value;
                        if (type == FieldCharValues.Begin)
                        {
                            DebugLog.AppendLine("[Converter] FieldChar Begin encountered");
                            inComplexField = true;
                            complexFieldHasSeparator = false;
                            currentComplexInstr = string.Empty;
                        }
                        else if (type == FieldCharValues.Separate)
                        {
                            DebugLog.AppendLine($"[Converter] FieldChar Separate encountered; instr='{currentComplexInstr}'");
                            complexFieldHasSeparator = true;

                            // 如果字段是超链接，请从解析器获取 URL 并激活锚点
                            if (!string.IsNullOrEmpty(currentComplexInstr) &&
                                currentComplexInstr.TrimStart().ToUpperInvariant().StartsWith("HYPERLINK"))
                            {
                                currentFieldHref = FieldResolver?.Invoke(currentComplexInstr);
                                DebugLog.AppendLine($"[Converter] Resolved hyperlink target: {currentFieldHref}");
                            }
                        }
                        else if (type == FieldCharValues.End)
                        {
                            // 当字段结束时，如果文档中没有提供显示文本（即未出现分隔符），
                            // 则使用 ProcessComplexField 生成输出；否则内容已经由后续 Run 处理。
                            if (!complexFieldHasSeparator && !string.IsNullOrEmpty(currentComplexInstr))
                            {
                                var cmd = currentComplexInstr.Trim().Split(' ', '\t')[0].ToUpperInvariant();
                                if (cmd != "HYPERLINK")
                                {
                                    ProcessComplexField(currentComplexInstr, pdfParagraph, ref hasContent, isHeading, headingSize, runProps, paraRunProps, actualFontSize, forceBold);
                                }
                            }

                            inComplexField = false;
                            currentComplexInstr = null;
                            currentFieldHref = null; // 清除链接状态
                            complexFieldHasSeparator = false;
                        }
                    }
                    break;

                case DocumentFormat.OpenXml.Wordprocessing.FieldCode fieldCode:
                    if (inComplexField && !complexFieldHasSeparator && fieldCode.InnerText != null)
                    {
                        currentComplexInstr += fieldCode.InnerText;
                    }
                    break;

                case BookmarkStart bookmarkStart:
                    var bmName = bookmarkStart.Name?.Value;
                    if (!string.IsNullOrEmpty(bmName) && BookmarkTracker is BookmarkTracker tracker)
                    {
                        tracker.AddBookmark(bmName);
                    }
                    break;
                case BookmarkEnd:
                    // 书签结束标记，不需要处理
                    break;
                    
                // 公式支持（Office Math ML）
                case DocumentFormat.OpenXml.Math.Paragraph mathParagraph:
                    var mathChunks = MathHelper.ExtractMathChunks(mathParagraph);
                    foreach (var chunk in mathChunks)
                    {
                        var mathFont = _fontHelper.GetFont(actualFontSize * 0.9f, iTextFont.NORMAL);
                        chunk.Font = mathFont;
                        pdfParagraph.Add(chunk);
                        hasContent = true;
                    }

                    // VML Support (WordArt, Shapes)
                    if (_mainPart != null)
                    {
                        var vmlChunks = VmlHelper.ExtractVmlElements(mathParagraph, _mainPart);
                        foreach (var chunk in vmlChunks)
                        {
                            pdfParagraph.Add(chunk);
                            hasContent = true;
                        }
                    }
                    break;
                    
                case Run run:
                    // 先检查嵌套的字段相关元素，让状态机响应 Begin/Separate/End
                    foreach (var nested in run.ChildElements)
                    {
                        if (nested is DocumentFormat.OpenXml.Wordprocessing.FieldChar ||
                            nested is DocumentFormat.OpenXml.Wordprocessing.FieldCode)
                        {
                            AppendInline(nested);
                        }
                    }

                    // 如果当前正在处理复杂字段的指令部分，则忽略剩余内容
                    if (inComplexField && !complexFieldHasSeparator)
                    {
                        break;
                    }

                    // 检查是否在删除的修订中
                    if (IsDeletedRevision(run))
                    {
                        // 跳过低格内容（或者可以用删除线显示）
                        break;
                    }
                    
                    var rProps = run.RunProperties;
                    // 如果是标题或是段落本身要求加粗，或者 Run 要求加粗，或者强制要求加粗
                    var font = isHeading
                        ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                        : _fontHelper.GetFont(rProps, paraProps?.GetFirstChild<ParagraphMarkRunProperties>(), actualFontSize, forceBold);
                    
                    if (forceBold && (font.Style & iTextFont.BOLD) == 0 && !RunExplicitlyDisablesBold(rProps) && !FontLooksBold(font))
                        font = new iTextFont(font.Family, font.Size, font.Style | iTextFont.BOLD, font.Color);
                    
                    var runRes = ConvertRun(run, isHeading, headingSize, font);
                    foreach (var res in runRes)
                    {
                        if (res is iTextChunk c && c.Content == "NEXTPAGE_SIGNAL")
                        {
                            // 提交当前段落并开启分页
                            if (hasContent) elements.Add(pdfParagraph);
                            elements.Add(new iTextChunk("PAGE_BREAK")); // 临时占位
                            
                            // 开始新的一段继续承接剩余内容
                            pdfParagraph = new iTextParagraph();
                            if (justification != null)
                                pdfParagraph.Alignment = StyleHelper.ToiTextAlignment(justification);
                            else
                                pdfParagraph.Alignment = StyleHelper.ToiTextAlignment(JustificationValues.Left);
                            SetParagraphSpacing(pdfParagraph, effectiveSpacing, docxParagraph.Parent is TableCell, isHeading, actualFontSize, baseLineHeight);
                            SetParagraphIndentation(pdfParagraph, effectiveIndentation);
                            hasContent = false;
                        }
                        else if (res is iTextChunk chunk)
                        {
                            // 检查插入的修订，用不同颜色显示
                            if (IsInsertedRevision(run))
                            {
                                var newFont = new iTextFont(chunk.Font.Family, chunk.Font.Size, chunk.Font.Style, new BaseColor(0, 100, 0));
                                chunk.Font = newFont;
                            }
                            // 检查删除的修订，用删除线显示
                            else if (IsDeletedRevision(run))
                            {
                                var newStyle = chunk.Font.Style | iTextFont.STRIKETHRU;
                                var newFont = new iTextFont(chunk.Font.Family, chunk.Font.Size, newStyle, new BaseColor(150, 0, 0));
                                chunk.Font = newFont;
                            }

                            // 如果当前段落处于字段超链接上下文，附加锚点
                            if (!string.IsNullOrEmpty(currentFieldHref))
                            {
                                // 若原本已有锚点，则保留它
                                if (string.IsNullOrEmpty(chunk.Anchor))
                                {
                                    chunk.SetAnchor(currentFieldHref);
                                }
                            }
                            
                            pdfParagraph.Add(chunk);
                            hasContent = true;
                        }
                    }
                    break;

                case Hyperlink hyperlink:
                    var linkColor = StyleHelper.ResolveSchemeColor(_colorScheme, "hlink") ?? BaseColor.Blue;
                    var linkUri = ResolveHyperlinkUri(hyperlink);
                    foreach (var hlRun in hyperlink.Elements<Run>())
                    {
                        var hlRProps = hlRun.RunProperties;
                        var hlFont = isHeading
                            ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                            : _fontHelper.GetFont(hlRProps, paraProps?.GetFirstChild<ParagraphMarkRunProperties>(), actualFontSize, forceBold);
                            
                        var hlChunks = ConvertRun(hlRun, isHeading, headingSize, hlFont);
                        foreach (var chunk in hlChunks)
                        {
                            chunk.Font = chunk.Font.WithColor(linkColor);
                            chunk.SetUnderline(0.1f, -1f);
                            if (!string.IsNullOrEmpty(linkUri))
                                chunk.SetAnchor(linkUri);
                            pdfParagraph.Add(chunk);
                            hasContent = true;
                        }
                    }
                    break;
                case FootnoteReference footnoteRef:
                    var fnIdVal = footnoteRef.Id?.Value;
                    if (fnIdVal.HasValue)
                    {
                        var fnId = (int)fnIdVal.Value;
                        FootnoteIdsEncountered?.Add(fnId);
                        if (_footnoteNumberById != null && _footnoteNumberById.TryGetValue(fnId, out var fnNum))
                        {
                            var fnFont = _fontHelper.GetFont(actualFontSize * 0.7f, iTextFont.NORMAL);
                            var fnChunk = new iTextChunk(fnNum.ToString(), fnFont);
                            fnChunk.SetTextRise(actualFontSize * 0.35f);
                            pdfParagraph.Add(fnChunk);
                            hasContent = true;
                        }
                    }
                    break;
                case EndnoteReference endnoteRef:
                    var enIdVal = endnoteRef.Id?.Value;
                    if (enIdVal.HasValue)
                    {
                        var enId = (int)enIdVal.Value;
                        EndnoteIdsEncountered?.Add(enId);
                        if (_endnoteNumberById != null && _endnoteNumberById.TryGetValue(enId, out var enNum))
                        {
                            var enFont = _fontHelper.GetFont(actualFontSize * 0.7f, iTextFont.NORMAL);
                            var enChunk = new iTextChunk(enNum.ToString(), enFont);
                            enChunk.SetTextRise(actualFontSize * 0.35f);
                            pdfParagraph.Add(enChunk);
                            hasContent = true;
                        }
                    }
                    break;
                case SimpleField field:
                    var instr = field.Instruction?.Value?.Trim();
                    if (!string.IsNullOrEmpty(instr))
                    {
                        var cmd = instr.Split(' ', '\t')[0].ToUpperInvariant();
                        if (cmd == "PAGE" || cmd == "NUMPAGES")
                        {
                            var pageInfo = PageNumberProvider?.Invoke();
                            var display = pageInfo.HasValue
                                ? (cmd == "PAGE" ? pageInfo.Value.Current.ToString() : pageInfo.Value.Total.ToString())
                                : "?";
                            var fieldFont = isHeading
                                ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                                : _fontHelper.GetFont(runProps, paraRunProps, actualFontSize, forceBold);
                            pdfParagraph.Add(new iTextChunk(display, fieldFont));
                            hasContent = true;
                        }
                        else if (cmd == "HYPERLINK")
                        {
                            // 超链接字段：保持显示文本不变，但将后续生成的所有 Chunk
                            // 设置为可点击链接。FieldResolver 仍返回 URL。
                            var link = FieldResolver?.Invoke(instr);
                            var saved = currentFieldHref;
                            currentFieldHref = link;
                            foreach (var child in field.ChildElements)
                                AppendInline(child);
                            currentFieldHref = saved;
                            hasContent = true;
                        }
                        else
                        {
                            var resolved = FieldResolver?.Invoke(instr);
                            if (!string.IsNullOrEmpty(resolved))
                            {
                                var fieldFont = isHeading
                                    ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                                    : _fontHelper.GetFont(runProps, paraRunProps, actualFontSize, forceBold);
                                pdfParagraph.Add(new iTextChunk(resolved, fieldFont));
                                hasContent = true;
                            }
                            else
                            {
                                foreach (var child in field.ChildElements)
                                    AppendInline(child);
                            }
                        }
                    }
                    else
                    {
                        foreach (var child in field.ChildElements)
                            AppendInline(child);
                    }
                    break;
                case SdtRun sdtRun:
                    var sdtContent = sdtRun.SdtContentRun;
                    if (sdtContent != null)
                    {
                        foreach (var child in sdtContent.ChildElements)
                            AppendInline(child);
                    }
                    break;
                    
                // 复杂字段支持（FieldBegin, FieldSeparator, FieldEnd）
                // 注意：这些类型在 DocumentFormat.OpenXml SDK 中存在于不同的命名空间
                // 目前我们使用 SimpleField + FieldResolver 来处理大多数情况（包括 HYPERLINK 字段），
                // 日后可进一步增强对传统 FieldChar 结构的完整支持。
            }
        }

        foreach (var element in docxParagraph.ChildElements)
        {
            AppendInline(element);
        }

        // 检查该段落中是否曾产生过 PAGE_BREAK，如果没有其它内容，则不必加空格充高度
        var hasPageBreak = elements.Any(e => e is iTextChunk chunk && chunk.Content == "PAGE_BREAK");

        // 空段落也输出，保持文档间距
        if (!hasContent && !hasPageBreak)
        {
            var emptyFont = _fontHelper.GetFont(headingSize ?? 12f);
            pdfParagraph.Add(new iTextChunk(" ", emptyFont));
        }

        string? headingTitle = null;
        if (isHeading && headingLevel.HasValue && hasContent)
        {
            headingTitle = string.Join("", docxParagraph.Descendants<Text>().Select(t => t.Text)).Trim();
            if (!string.IsNullOrWhiteSpace(headingTitle))
            {
                var bookmarkId = docxParagraph.Descendants<BookmarkStart>().FirstOrDefault()?.Name?.Value;
                pdfParagraph.OutlineTitle = headingTitle;
                pdfParagraph.OutlineLevel = headingLevel.Value;
                pdfParagraph.OutlineKey = TableOfContentsGenerator.BuildEntryKey(headingTitle, headingLevel.Value, bookmarkId);
                pdfParagraph.RenderedCallback = (paragraph, pageNumber) =>
                {
                    if (!string.IsNullOrWhiteSpace(paragraph.OutlineKey)
                        && !string.IsNullOrWhiteSpace(paragraph.OutlineTitle)
                        && paragraph.OutlineLevel.HasValue)
                    {
                        HeadingRendered?.Invoke(paragraph.OutlineKey, paragraph.OutlineTitle, paragraph.OutlineLevel.Value, pageNumber);
                    }
                };
            }
        }

        if (hasContent || !hasPageBreak)
        {
            elements.Add(pdfParagraph);
        }

        // 标题书签
        if (isHeading && headingLevel.HasValue && hasContent && BookmarkTracker is BookmarkTracker tracker)
        {
            if (!string.IsNullOrWhiteSpace(headingTitle))
            {
                tracker.AddHeadingBookmark(headingTitle, headingLevel.Value);
            }
        }

        // 段落边框与底纹：用单列表格包裹
        var paraBorders = paraProps?.ParagraphBorders ?? GetStyleParagraphBorders(styleId);
        var paraShading = paraProps?.Shading ?? GetStyleParagraphShading(styleId);
        if (paraBorders != null || paraShading != null)
        {
            elements = WrapWithParagraphBordersAndShading(elements, paraBorders, paraShading);
        }
        
        return elements;
    }

    private ParagraphBorders? GetStyleParagraphBorders(string? styleId)
    {
        return GetFromStyleChain(styleId, s => s.StyleParagraphProperties?.GetFirstChild<ParagraphBorders>());
    }

    private Shading? GetStyleParagraphShading(string? styleId)
    {
        return GetFromStyleChain(styleId, s => s.StyleParagraphProperties?.GetFirstChild<Shading>());
    }

    private System.Collections.Generic.List<IElement> WrapWithParagraphBordersAndShading(
        System.Collections.Generic.List<IElement> elements,
        ParagraphBorders? borders,
        Shading? shading)
    {
        var result = new System.Collections.Generic.List<IElement>();
        foreach (var el in elements)
        {
            if (el is iTextChunk chunk && chunk.Content == "PAGE_BREAK")
            {
                result.Add(el);
                continue;
            }
            if (el is iTextParagraph para)
            {
                var table = new PdfPTable(1) { WidthPercentage = 100, SpacingAfter = para.SpacingAfter };
                var cell = new PdfPCell(para) { Padding = 4f };
                if (shading != null)
                {
                    var bg = StyleHelper.ResolveShadingFill(_colorScheme, shading);
                    if (bg != null) cell.BackgroundColor = bg;
                }
                ApplyParagraphBordersToCell(cell, borders);
                table.AddCell(cell);
                result.Add(table);
            }
            else
            {
                result.Add(el);
            }
        }
        return result;
    }

    private void ApplyParagraphBordersToCell(PdfPCell cell, ParagraphBorders? borders)
    {
        if (borders == null) return;
        cell.UseVariableBorders = true;
        void SetBorder(BorderType? b, Action<float, BaseColor?> setter)
        {
            if (b == null) return;
            var w = StyleHelper.GetBorderWidth(b);
            var c = StyleHelper.ResolveBorderColor(_colorScheme, b);
            if (w > 0) setter(w, c ?? BaseColor.Black);
        }
        SetBorder(borders.TopBorder, (w, c) => { cell.BorderWidthTop = w; if (c != null) cell.BorderColorTop = c; });
        SetBorder(borders.BottomBorder, (w, c) => { cell.BorderWidthBottom = w; if (c != null) cell.BorderColorBottom = c; });
        SetBorder(borders.LeftBorder, (w, c) => { cell.BorderWidthLeft = w; if (c != null) cell.BorderColorLeft = c; });
        SetBorder(borders.RightBorder, (w, c) => { cell.BorderWidthRight = w; if (c != null) cell.BorderColorRight = c; });
    }

    private static bool FontLooksBold(iTextFont? font)
    {
        if (font == null) return false;
        var family = font.Family;
        if (string.IsNullOrWhiteSpace(family)) return false;
        if (family.Contains("Bold", StringComparison.OrdinalIgnoreCase)) return true;
        if (family.Contains("Black", StringComparison.OrdinalIgnoreCase)) return true;
        if (family.Contains("Heavy", StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private static bool ShouldApplySyntheticBold(iTextFont font)
    {
        if ((font.Style & iTextFont.BOLD) == 0) return false;
        var ps = font.BaseFont?.PostscriptFontName;
        if (string.IsNullOrWhiteSpace(ps)) return false;
        if (ps.Contains("Bold", StringComparison.OrdinalIgnoreCase)) return false;
        if (ps.Contains("Black", StringComparison.OrdinalIgnoreCase)) return false;
        if (ps.Contains("Heavy", StringComparison.OrdinalIgnoreCase)) return false;
        if (ps.Contains("SimHei", StringComparison.OrdinalIgnoreCase)) return false;
        if (ps.Contains("Heiti", StringComparison.OrdinalIgnoreCase)) return false;
        return ps.Contains("STSong", StringComparison.OrdinalIgnoreCase);
    }

    private static bool RunExplicitlyDisablesBold(RunProperties? runProps)
    {
        if (runProps == null) return false;
        var b = runProps.GetFirstChild<Bold>();
        if (b?.Val?.Value == false) return true;
        var bcs = runProps.GetFirstChild<BoldComplexScript>();
        if (bcs?.Val?.Value == false) return true;
        return false;
    }

    private static void SetParagraphSpacing(iTextParagraph pdfParagraph, SpacingBetweenLines? spacing, bool inTableCell, bool isHeading, float fontSize, float baseLineHeight)
    {
        float minSpacing = Math.Max(8f, fontSize * 0.3f);
        float defaultMultiplier = fontSize > 16f ? 1.2f : 1.15f;

        if (spacing == null)
        {
            pdfParagraph.SetLeading(baseLineHeight * defaultMultiplier, 0);
            pdfParagraph.SpacingBefore = isHeading ? Math.Max(minSpacing * 0.5f, 4f) : 0f;
            pdfParagraph.SpacingAfter = inTableCell ? 0f : minSpacing;
            return;
        }

        var hasBefore = spacing.Before != null || spacing.BeforeLines != null;
        var hasAfter = spacing.After != null || spacing.AfterLines != null;

        // 段前距
        if (spacing.Before?.Value is string beforeStr)
        {
            var before = StyleHelper.TwipsToPoints(beforeStr);
            pdfParagraph.SpacingBefore = before;
        }
        else if (spacing.BeforeLines?.Value is int beforeLines && beforeLines > 0)
        {
            pdfParagraph.SpacingBefore = fontSize * (beforeLines / 100f);
        }
        else if (isHeading)
        {
            pdfParagraph.SpacingBefore = minSpacing * 0.5f;
        }

        // 段后距
        if (spacing.After?.Value is string afterStr)
        {
            var after = StyleHelper.TwipsToPoints(afterStr);
            pdfParagraph.SpacingAfter = after;
        }
        else if (spacing.AfterLines?.Value is int afterLines && afterLines > 0)
        {
            pdfParagraph.SpacingAfter = fontSize * (afterLines / 100f);
        }
        else
        {
            pdfParagraph.SpacingAfter = inTableCell ? 0f : minSpacing;
        }

        // 行距
        if (spacing.Line?.Value is string lineStr && float.TryParse(lineStr, out var lineSpacing))
        {
            var lineRule = spacing.LineRule?.Value;
            float leading = lineSpacing / 20f; // 以 point 为单位 (主要用于 Exact 和 AtLeast)

            bool isExact = lineRule != null && lineRule.Equals(LineSpacingRuleValues.Exact);
            bool isAtLeast = lineRule != null && lineRule.Equals(LineSpacingRuleValues.AtLeast);

            if (isExact || isAtLeast)
            {
                var minLeading = baseLineHeight * defaultMultiplier;
                pdfParagraph.SetLeading(Math.Max(leading, minLeading), 0);
            }
            else
            {
                var lines = lineSpacing / 240f;
                var multiplier = lines > 0 ? lines : defaultMultiplier;
                pdfParagraph.SetLeading(baseLineHeight * multiplier, 0);
            }
        }
        else
        {
            pdfParagraph.SetLeading(baseLineHeight * defaultMultiplier, 0);
        }

        if (inTableCell)
        {
            if (!hasBefore) pdfParagraph.SpacingBefore = 0f;
            if (!hasAfter) pdfParagraph.SpacingAfter = 0f;
        }
    }

    private static float GetBaseLineHeight(iTextFont? font, float fontSize)
    {
        if (font == null || fontSize <= 0) return Math.Max(fontSize, 12f);
        // 简化实现：使用字体大小作为行高
        return Math.Max(fontSize * 1.2f, 12f);
    }

    /// <summary>
    /// 设置段落缩进
    /// </summary>
    private static void SetParagraphIndentation(iTextParagraph pdfParagraph, Indentation? indent)
    {
        if (indent == null) return;

        if (indent.Left?.Value is string leftStr)
        {
            pdfParagraph.IndentationLeft = StyleHelper.TwipsToPoints(leftStr);
        }

        if (indent.Right?.Value is string rightStr)
        {
            pdfParagraph.IndentationRight = StyleHelper.TwipsToPoints(rightStr);
        }

        if (indent.FirstLine?.Value is string firstLineStr)
        {
            pdfParagraph.FirstLineIndent = StyleHelper.TwipsToPoints(firstLineStr);
        }

        if (indent.Hanging?.Value is string hangingStr)
        {
            var hanging = StyleHelper.TwipsToPoints(hangingStr);
            pdfParagraph.FirstLineIndent = -hanging;
            pdfParagraph.IndentationLeft += hanging;
        }
    }

    /// <summary>
    /// 将 DOCX 高亮颜色映射为 BaseColor
    /// </summary>
    private static BaseColor? HighlightToBaseColor(HighlightColorValues highlight)
    {
        if (highlight.Equals(HighlightColorValues.Yellow)) return new BaseColor(255, 255, 0);
        if (highlight.Equals(HighlightColorValues.Green)) return new BaseColor(0, 255, 0);
        if (highlight.Equals(HighlightColorValues.Cyan)) return new BaseColor(0, 255, 255);
        if (highlight.Equals(HighlightColorValues.Magenta)) return new BaseColor(255, 0, 255);
        if (highlight.Equals(HighlightColorValues.Blue)) return new BaseColor(0, 0, 255);
        if (highlight.Equals(HighlightColorValues.Red)) return new BaseColor(255, 0, 0);
        if (highlight.Equals(HighlightColorValues.DarkBlue)) return new BaseColor(0, 0, 139);
        if (highlight.Equals(HighlightColorValues.DarkCyan)) return new BaseColor(0, 139, 139);
        if (highlight.Equals(HighlightColorValues.DarkGreen)) return new BaseColor(0, 100, 0);
        if (highlight.Equals(HighlightColorValues.DarkMagenta)) return new BaseColor(139, 0, 139);
        if (highlight.Equals(HighlightColorValues.DarkRed)) return new BaseColor(139, 0, 0);
        if (highlight.Equals(HighlightColorValues.DarkYellow)) return new BaseColor(139, 139, 0);
        if (highlight.Equals(HighlightColorValues.DarkGray)) return new BaseColor(169, 169, 169);
        if (highlight.Equals(HighlightColorValues.LightGray)) return new BaseColor(211, 211, 211);
        if (highlight.Equals(HighlightColorValues.Black)) return BaseColor.Black;
        if (highlight.Equals(HighlightColorValues.White)) return BaseColor.White;
        
        return null;
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

    private Indentation? GetStyleIndentation(string? styleId)
    {
        return GetFromStyleChain(styleId, s => s.StyleParagraphProperties?.GetFirstChild<Indentation>());
    }

    private JustificationValues? GetStyleJustification(string? styleId)
    {
        var j = GetFromStyleChain(styleId, s => s.StyleParagraphProperties?.GetFirstChild<Justification>());
        return j?.Val?.Value;
    }

    private string? GetStyleFontSize(string? styleId)
    {
        var sz = GetFromStyleChain(styleId, s => s.StyleRunProperties?.GetFirstChild<FontSize>());
        return sz?.Val?.Value;
    }

    private void ApplyParagraphKeepOptions(iTextParagraph pdfParagraph, ParagraphProperties? paraProps, string? styleId)
    {
        // KeepLinesTogether: 段落内容保持在一起不拆分
        var keepLines = paraProps?.KeepLines ?? GetStyleKeepLines(styleId);
        if (keepLines != null && (keepLines.Val == null || keepLines.Val.Value))
        {
            pdfParagraph.KeepTogether = true;
        }

        // KeepWithNext: 与下一段保持在同一页（iTextSharp 支持有限，主要用于标题）
        // 注意：iTextSharp 不直接支持 KeepWithNext，这里只做标记
        // 实际应用需要在文档级别处理，将连续的 KeepWithNext 段落组合
    }

    private KeepLines? GetStyleKeepLines(string? styleId)
    {
        return GetFromStyleChain(styleId, s => s.StyleParagraphProperties?.GetFirstChild<KeepLines>());
    }

    private static int? GetHeadingLevel(string? styleId)
    {
        if (string.IsNullOrWhiteSpace(styleId)) return null;
        var lower = styleId.ToLowerInvariant();
        if (lower.StartsWith("heading"))
        {
            var numPart = lower.Substring("heading".Length);
            if (int.TryParse(numPart, out var level) && level >= 1 && level <= 9)
                return level;
        }
        return null;
    }


    
    /// <summary>
    /// 处理复杂字段（由 FieldBegin、FieldSeparator、FieldEnd 组成）
    /// </summary>
    private void ProcessComplexField(
        string instruction,
        iTextParagraph pdfParagraph,
        ref bool hasContent,
        bool isHeading,
        float? headingSize,
        DocumentFormat.OpenXml.Wordprocessing.RunProperties? runProps,
        DocumentFormat.OpenXml.Wordprocessing.ParagraphMarkRunProperties? paraRunProps,
        float actualFontSize,
        bool forceBold)
    {
        try
        {
            var instr = instruction.Trim();
            if (string.IsNullOrEmpty(instr)) return;
            
            // 提取字段命令
            var cmd = instr.Split(' ', '\t')[0].ToUpperInvariant();
            
            // 处理常见字段
            if (cmd == "PAGE" || cmd == "NUMPAGES")
            {
                var pageInfo = PageNumberProvider?.Invoke();
                var display = pageInfo.HasValue
                    ? (cmd == "PAGE" ? pageInfo.Value.Current.ToString() : pageInfo.Value.Total.ToString())
                    : "?";
                var fieldFont = isHeading
                    ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                    : _fontHelper.GetFont(runProps, paraRunProps, actualFontSize, forceBold);
                pdfParagraph.Add(new iTextChunk(display, fieldFont));
                hasContent = true;
            }
            else if (cmd == "DATE" || cmd == "TIME")
            {
                // 解析日期/时间格式
                var formatMatch = System.Text.RegularExpressions.Regex.Match(instr, @"\\@\s*""([^""]+)""");
                var format = formatMatch.Success ? formatMatch.Groups[1].Value : (cmd == "DATE" ? "yyyy-MM-dd" : "HH:mm:ss");
                
                var display = System.DateTime.Now.ToString(format);
                var fieldFont = isHeading
                    ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                    : _fontHelper.GetFont(runProps, paraRunProps, actualFontSize, forceBold);
                pdfParagraph.Add(new iTextChunk(display, fieldFont));
                hasContent = true;
            }
            else if (cmd == "AUTHOR" || cmd == "TITLE" || cmd == "SUBJECT" || cmd == "KEYWORDS")
            {
                // 使用字段解析器
                var resolved = FieldResolver?.Invoke(instr);
                if (!string.IsNullOrEmpty(resolved))
                {
                    var fieldFont = isHeading
                        ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                        : _fontHelper.GetFont(runProps, paraRunProps, actualFontSize, forceBold);
                    pdfParagraph.Add(new iTextChunk(resolved, fieldFont));
                    hasContent = true;
                }
            }
            else if (cmd == "MERGEFIELD")
            {
                // 邮件合并字段
                var resolved = FieldResolver?.Invoke(instr);
                if (!string.IsNullOrEmpty(resolved))
                {
                    var fieldFont = isHeading
                        ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                        : _fontHelper.GetFont(runProps, paraRunProps, actualFontSize, forceBold);
                    pdfParagraph.Add(new iTextChunk(resolved, fieldFont));
                    hasContent = true;
                }
            }
            else if (cmd == "REF")
            {
                // 交叉引用
                var resolved = FieldResolver?.Invoke(instr);
                if (!string.IsNullOrEmpty(resolved))
                {
                    var fieldFont = isHeading
                        ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                        : _fontHelper.GetFont(runProps, paraRunProps, actualFontSize, forceBold);
                    pdfParagraph.Add(new iTextChunk(resolved, fieldFont));
                    hasContent = true;
                }
            }
            else
            {
                // 其他字段，尝试使用通用解析器
                var resolved = FieldResolver?.Invoke(instr);
                if (!string.IsNullOrEmpty(resolved))
                {
                    var fieldFont = isHeading
                        ? _fontHelper.GetFont(headingSize ?? 16f, iTextFont.BOLD)
                        : _fontHelper.GetFont(runProps, paraRunProps, actualFontSize, forceBold);
                    pdfParagraph.Add(new iTextChunk(resolved, fieldFont));
                    hasContent = true;
                }
            }
        }
        catch
        {
            // 字段处理失败，忽略
        }
    }



    /// <summary>
    /// 检查 Run 是否属于插入的修订
    /// </summary>
    private static bool IsInsertedRevision(DocumentFormat.OpenXml.Wordprocessing.Run run)
    {
        var runProperties = run.RunProperties;
        if (runProperties == null) return false;
        
        // 检查属性中是否包含修订标记
        foreach (var attr in runProperties.GetAttributes())
        {
            if ((attr.LocalName == "ins" || attr.LocalName == "inserted") && 
                (attr.Value == "1" || attr.Value == "true"))
            {
                return true;
            }
        }
        
        return false;
    }

    /// <summary>
    /// 检查 Run 是否属于删除的修订
    /// </summary>
    private static bool IsDeletedRevision(DocumentFormat.OpenXml.Wordprocessing.Run run)
    {
        var runProperties = run.RunProperties;
        if (runProperties == null) return false;
        
        // 检查是否有删除标记（del 标签）
        var parentDel = run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Deleted>().FirstOrDefault();
        if (parentDel != null)
        {
            return true;
        }
        
        // 检查属性中是否包含修订标记
        foreach (var attr in runProperties.GetAttributes())
        {
            if ((attr.LocalName == "del" || attr.LocalName == "deleted") && 
                (attr.Value == "1" || attr.Value == "true"))
            {
                return true;
            }
        }
        
        return false;
    }

    private string? ResolveHyperlinkUri(Hyperlink hyperlink)
    {
        if (_hyperlinkTargets == null) return null;
        var id = hyperlink.Id?.Value;
        if (string.IsNullOrEmpty(id)) return null;
        if (!_hyperlinkTargets.TryGetValue(id, out var uri)) return null;
        var anchor = hyperlink.Anchor?.Value;
        if (!string.IsNullOrEmpty(anchor))
            uri = uri.Contains('#') ? uri : uri + "#" + anchor;
        return uri;
    }

}
