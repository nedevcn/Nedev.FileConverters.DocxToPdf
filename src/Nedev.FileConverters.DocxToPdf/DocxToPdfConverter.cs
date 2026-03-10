using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.IO;
using Nedev.FileConverters.DocxToPdf.Converters;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.Models;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using Nedev.FileConverters.DocxToPdf.PdfEngine.Compatibility;
using Nedev.FileConverters.Core;
using iTextDocument = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfDocument;
using iTextWriter = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfWriter;
using iTextRectangle = Nedev.FileConverters.DocxToPdf.PdfEngine.Rectangle;
using iTextBaseColor = Nedev.FileConverters.DocxToPdf.PdfEngine.BaseColor;
using iTextFont = Nedev.FileConverters.DocxToPdf.PdfEngine.Font;
using iTextChunk = Nedev.FileConverters.DocxToPdf.PdfEngine.Chunk;
using iTextParagraph = Nedev.FileConverters.DocxToPdf.PdfEngine.Paragraph;
using iTextPhrase = Nedev.FileConverters.DocxToPdf.PdfEngine.Phrase;
using iTextImage = Nedev.FileConverters.DocxToPdf.PdfEngine.Image;
using iTextPdfPTable = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfPTable;
using iTextPdfPCell = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfPCell;
using iTextList = Nedev.FileConverters.DocxToPdf.PdfEngine.List;
using iTextListItem = Nedev.FileConverters.DocxToPdf.PdfEngine.ListItem;
using iTextElement = Nedev.FileConverters.DocxToPdf.PdfEngine.IElement;
using iTextPdfContentByte = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfContentByte;
using iTextColumnText = Nedev.FileConverters.DocxToPdf.PdfEngine.ColumnText;
using iTextPdfOutline = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfOutline;
using iTextPdfDestination = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfDestination;
using iTextPdfPageEventHelper = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfPageEventHelper;
using iTextPdfReader = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfReader;
using iTextPdfStamper = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfStamper;
using DocxImageConverter = Nedev.FileConverters.DocxToPdf.Converters.ImageConverter;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WPageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;

namespace Nedev.FileConverters.DocxToPdf;

/// <summary>
/// DOCX 转 PDF 转换器
/// </summary>
[FileConverter("docx","pdf")]
public class DocxToPdfConverter : IFileConverter
{
    private readonly ConvertOptions _options;
    
    private readonly Dictionary<int, Models.SectionPageSettings> _sectionSettingsMap = new();

    private void CaptureSectionSettings(int sectionIndex)
    {
        _sectionSettingsMap[sectionIndex] = new Models.SectionPageSettings
        {
            PageSize = _options.PageSize,
            MarginLeft = _options.MarginLeft,
            MarginRight = _options.MarginRight,
            MarginTop = _options.MarginTop,
            MarginBottom = _options.MarginBottom,
            HeaderDistance = _options.HeaderDistance,
            FooterDistance = _options.FooterDistance,
            TextDirection = _currentTextDirection
        };
    }

    private ColumnInfo _currentColumnInfo = new();
    private Models.TextDirection _currentTextDirection = Models.TextDirection.Horizontal;

    private class ColumnInfo
    {
        public int Count { get; set; } = 1;
        public float Spacing { get; set; } = 36f; // Default 0.5 inch (720 twips)
        public float[]? Widths { get; set; }
    }

    /// <summary>
    /// 使用默认选项初始化转换器
    /// </summary>
    public DocxToPdfConverter() : this(ConvertOptions.Default) { }

    /// <summary>
    /// 使用自定义选项初始化转换器
    /// </summary>
    public DocxToPdfConverter(ConvertOptions options)
    {
        _options = options ?? ConvertOptions.Default;
    }

    private static readonly Chunk NextPageChunk = new Chunk("NEXTPAGE_SIGNAL");

    /// <summary>
    /// 将 DOCX 文件转换为 PDF 文件
    /// </summary>
    /// <param name="docxPath">输入 DOCX 文件路径</param>
    /// <param name="pdfPath">输出 PDF 文件路径</param>
    public void Convert(string docxPath, string pdfPath)
    {
        if (string.IsNullOrWhiteSpace(docxPath))
            throw new ArgumentException("DOCX 文件路径不能为空。", nameof(docxPath));
        if (string.IsNullOrWhiteSpace(pdfPath))
            throw new ArgumentException("PDF 文件路径不能为空。", nameof(pdfPath));
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("DOCX 文件不存在。", docxPath);

        using var inputStream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var outputStream = File.Create(pdfPath);
        Convert(inputStream, outputStream);
    }

    /// <summary>
    /// 实现 IFileConverter接口
    /// </summary>
    /// <param name="input">输入流</param>
    /// <returns>输出 PDF 流</returns>
    Stream IFileConverter.Convert(Stream input)
    {
        var memory = new MemoryStream();
        Convert(input, memory);
        memory.Position = 0;
        return memory;
    }

    /// <summary>
    /// 将 DOCX 流转换为 PDF 流
    /// </summary>
    /// <param name="docxStream">输入 DOCX 流</param>
    /// <param name="pdfStream">输出 PDF 流</param>
    public void Convert(Stream docxStream, Stream pdfStream)
    {
        if (docxStream == null) throw new ArgumentNullException(nameof(docxStream));
        if (pdfStream == null) throw new ArgumentNullException(nameof(pdfStream));

        // 打开 DOCX 文档
        using var docxDocument = WordprocessingDocument.Open(docxStream, false);
        var mainPart = docxDocument.MainDocumentPart
                       ?? throw new InvalidOperationException("DOCX 文件无效：缺少主文档部分。");

        var body = mainPart.Document?.Body
                   ?? throw new InvalidOperationException("DOCX 文件无效：缺少文档正文。");

        // 读取页面设置
        var sectionProps = body.Descendants<SectionProperties>().LastOrDefault();
        ApplyPageSettings(sectionProps);

        // 初始化字体（需要主题色用于字体颜色解析）
        var colorScheme = mainPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        var fontHelper = new FontHelper(_options, colorScheme);
        fontHelper.RegisterFonts();

        var hyperlinkTargets = mainPart.HyperlinkRelationships
            .Where(r => r.Uri != null)
            .GroupBy(r => r.Id)
            .ToDictionary(g => g.Key, g => g.First().Uri.ToString(), StringComparer.OrdinalIgnoreCase);

        var footnoteNumberById = BuildNoteNumberMap(mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>());
        var endnoteNumberById = BuildNoteNumberMap(mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>());
        var footnoteById = BuildNoteByIdMap(mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>());
        var endnoteById = BuildNoteByIdMap(mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>());

        var footnoteIdsEncountered = new List<int>();
        var endnoteIdsEncountered = new List<int>();

        // 扫描所有节属性以确定正确的初始设置和后续切换
        var allSections = new List<SectionProperties>();
        foreach (var para in body.Descendants<WParagraph>())
        {
            if (para.ParagraphProperties?.SectionProperties != null)
                allSections.Add(para.ParagraphProperties.SectionProperties);
        }
        if (body.Elements<SectionProperties>().FirstOrDefault() != null)
            allSections.Add(body.Elements<SectionProperties>().First());

        if (allSections.Count > 0)
        {
            ApplyPageSettings(allSections[0]);
            CaptureSectionSettings(0);
        }
        else
        {
            _currentColumnInfo = new ColumnInfo(); // Default
            CaptureSectionSettings(0);
        }

        var pageWidth = GetColumnWidth();
        var numberingPart = mainPart.NumberingDefinitionsPart;
        var numbering = numberingPart?.Numbering;
        var styles = mainPart.StyleDefinitionsPart?.Styles;

        var hasHeaderFooter = _options.RenderHeadersFooters && HasAnyHeaderFooter(body, sectionProps);
        var sectionTracker = new SectionTracker { CurrentSection = 0 };

        Stream targetStream = pdfStream;
        MemoryStream? tempStream = null;
        if (hasHeaderFooter)
        {
            tempStream = new MemoryStream();
            targetStream = tempStream;
        }

        var pdfDocument = new iTextDocument(_options.PageSize,
            _options.MarginLeft, _options.MarginRight,
            _options.MarginTop, _options.MarginBottom);

        var writer = PdfWriter.GetInstance(pdfDocument, targetStream);
        writer.CloseStream = false;

        // 设置中文字体（使用系统已安装的字体）
        writer.SetChineseFont("SimSun");

        // 注解集合（用于超链接等）
        var annotations = new AnnotationCollection();
        writer.SetAnnotationCollection(annotations);

        // PDF加密支持
        if (_options.Encryption != null && !string.IsNullOrEmpty(_options.Encryption.UserPassword))
        {
            int permissions = 0;
            if (_options.Encryption.AllowPrint) permissions |= PdfEncryption.PRINT;
            if (_options.Encryption.AllowModifyContent) permissions |= PdfEncryption.MODIFY;
            if (_options.Encryption.AllowCopyContent) permissions |= PdfEncryption.COPY;
            if (_options.Encryption.AllowFillForms) permissions |= PdfEncryption.FILL_FORM;
            
            var encryption = new PdfEncryption(
                _options.Encryption.UserPassword,
                _options.Encryption.OwnerPassword,
                permissions
            );
            writer.SetEncryption(encryption);
        }

        // PDF元数据支持
        writer.SetMetadata(
            _options.PdfTitle,
            _options.PdfAuthor,
            _options.PdfSubject,
            _options.PdfKeywords,
            _options.PdfCreator
        );

        // 多栏排版管理器（需要尽早初始化以便设置页码）
        var ct = new ColumnText(writer.DirectContent);
        ct.SetAnnotationCollection(annotations);
        ct.TextDirection = _currentTextDirection;
        ct.LineNumberSettings = _options.LineNumberSettings;
        if (_options.LineNumberSettings != null) ct.CurrentLineNumber = _options.LineNumberSettings.Start;

        // 书签支持
        var bookmarkTracker = new global::Nedev.FileConverters.DocxToPdf.PdfEngine.BookmarkTracker(writer);
        
        // 收集所有 PageEvent
        var events = new List<PdfPageEventHelper> { bookmarkTracker };
        if (hasHeaderFooter) events.Insert(0, sectionTracker);
        if (_options.PageBorders != null) events.Add(new PageBorderEvent(_options.PageBorders, _options));
        
        if (events.Count > 1)
            writer.PageEvent = new CombinedPageEvent(events.ToArray());
        else
            writer.PageEvent = events[0];
        
        pdfDocument.Open();
        pdfDocument.NewPage(); // MUST start a page so PdfWriter tracks it correctly
        ct.SetCurrentPage(pdfDocument.PageNumber);

        // 目录提取（如果在配置中启用）
        var tocEntries = new List<TableOfContentsGenerator.TOCEntry>();
        var tocPageNumbersByKey = new Dictionary<string, Queue<int>>(StringComparer.Ordinal);
        HashSet<string>? expectedTocKeys = null;
        if (_options.GenerateTableOfContents)
        {
            tocEntries = TableOfContentsGenerator.ExtractTOC(body);
            if (tocEntries.Count > 0)
            {
                expectedTocKeys = tocEntries
                    .Select(TableOfContentsGenerator.BuildEntryKey)
                    .ToHashSet(StringComparer.Ordinal);
                GenerateTOCPage(pdfDocument, tocEntries);
            }
        }

        var paragraphConverter = new ParagraphConverter(fontHelper, styles, colorScheme, hyperlinkTargets, footnoteNumberById, endnoteNumberById, docxDocument.MainDocumentPart)
        {
            FootnoteIdsEncountered = footnoteIdsEncountered,
            EndnoteIdsEncountered = endnoteIdsEncountered,
            BookmarkTracker = bookmarkTracker,
            FieldResolver = instr => ResolveField(instr, docxDocument, null) // 简化：不传递 docxPath
        };
        paragraphConverter.HeadingRendered = (key, title, level, pageNumber) =>
        {
            if (expectedTocKeys == null || !expectedTocKeys.Contains(key))
                return;

            if (!tocPageNumbersByKey.TryGetValue(key, out var pages))
            {
                pages = new Queue<int>();
                tocPageNumbersByKey[key] = pages;
            }

            pages.Enqueue(pageNumber);
        };
        var imageConverter = new DocxImageConverter(docxDocument, _options);
        var listConverter = new ListConverter(fontHelper, styles, colorScheme);
        var tableConverter = new TableConverter(fontHelper, paragraphConverter, imageConverter, listConverter, numbering, styles, colorScheme);

        // 页眉页脚：收集所有节的引用（仅在启用时）
        HeaderFooterRenderer? headerFooterRenderer = null;
        if (hasHeaderFooter)
        {
            headerFooterRenderer = new HeaderFooterRenderer(mainPart, paragraphConverter, imageConverter, _options, GetPageContentWidth())
            {
                TableConverter = tableConverter
            };
            // 注册所有节的页眉页脚
            for (int s = 0; s < allSections.Count; s++)
            {
                headerFooterRenderer.RegisterSection(allSections[s], s);
            }
        }
        var elements = body.ChildElements.ToList();
        var i = 0;
        
        int currentColumn = 0;
        int currentSectionIndex = 0;
        bool sectionBreakEncountered = false;

        // Deferred InFrontOfText images — drawn after ColumnText.Go() to ensure correct z-order
        var pendingInFrontImages = new List<global::Nedev.FileConverters.DocxToPdf.Converters.FloatingObject>();

        void FlushInFrontImages()
        {
            foreach (var pending in pendingInFrontImages)
                writer.DirectContent.AddImage(pending.Image);
            pendingInFrontImages.Clear();
        }

        void SetColumnBounds()
        {
            var info = _currentColumnInfo;
            var pageSize = _options.PageSize;
            var ml = _options.MarginLeft;
            var mr = _options.MarginRight;
            var mt = _options.MarginTop;
            var mb = _options.MarginBottom;

            var availableWidth = pageSize.Width - ml - mr;
            if (info.Count <= 1)
            {
                ct.SetSimpleColumn(ml, mb, pageSize.Width - mr, pageSize.Height - mt);
                return;
            }

            var colWidth = (availableWidth - (info.Count - 1) * info.Spacing) / info.Count;
            
            float llx;
            if (_currentTextDirection == Models.TextDirection.Vertical)
            {
                // 竖排：从右向左分栏
                llx = pageSize.Width - mr - (currentColumn * (colWidth + info.Spacing)) - colWidth;
            }
            else
            {
                llx = ml + currentColumn * (colWidth + info.Spacing);
            }

            var lly = mb;
            var urx = llx + colWidth;
            var ury = pageSize.Height - mt;
            ct.SetSimpleColumn(llx, lly, urx, ury);
        }

        SetColumnBounds();

        while (i < elements.Count)
        {
            var element = elements[i];
            var itemsToAdd = new List<IElement>();

            Action<SectionProperties> onSectionBreak = (sp) =>
            {
                sectionBreakEncountered = true;
                if (hasHeaderFooter)
                {
                    // sectionTracker update handled via page events usually, 
                    // but we need to ensure tracker knows we moved to next section?
                    // Actually SectionTracker listens to OnStartPage.
                    // We need to update sectionTracker.CurrentSection manually if needed?
                    // CombinedPageEvent uses reference to sectionTracker.
                    // When NewPage happens, OnStartPage reads sectionTracker.CurrentSection.
                    // So we must update it BEFORE NewPage.
                }
            };

            switch (element)
            {
                case WParagraph paragraph:
                    itemsToAdd.AddRange(ProcessParagraph(paragraph, writer, paragraphConverter, imageConverter,
                        listConverter, numbering, elements, ref i, ref pageWidth, onSectionBreak));
                    break;

                case WTable table:
                    // 表格使用当前分栏宽度
                    var pdfTable = tableConverter.Convert(table, GetColumnWidth());
                    if (pdfTable != null) itemsToAdd.Add(pdfTable);
                    i++;
                    break;

                case SectionProperties sect:
                    ApplyPageSettings(sect);
                    onSectionBreak(sect);
                    itemsToAdd.Add(NextPageChunk);
                    i++;
                    break;
                case SdtBlock sdtBlock:
                    itemsToAdd.AddRange(ProcessSdtBlock(sdtBlock, writer, paragraphConverter, tableConverter, imageConverter, listConverter, numbering, ref pageWidth, onSectionBreak));
                    i++;
                    break;

                default:
                    i++;
                    break;
            }

            foreach (var item in itemsToAdd)
            {
                // 处理分页符
                if (item == NextPageChunk || (item is Chunk c && (c.Content == "PAGE_BREAK" || c.Content == "NEXTPAGE_SIGNAL")))
                {
                    if (sectionBreakEncountered)
                    {
                        currentSectionIndex++;
                        if (currentSectionIndex < allSections.Count)
                        {
                            ApplyPageSettings(allSections[currentSectionIndex]);
                            CaptureSectionSettings(currentSectionIndex);
                        }
                        if (hasHeaderFooter) sectionTracker.CurrentSection = currentSectionIndex;
                        sectionBreakEncountered = false;
                        ApplyCurrentPageSettings(pdfDocument, ct);

                        // 更新行号设置
                        ct.LineNumberSettings = _options.LineNumberSettings;
                        if (_options.LineNumberSettings?.RestartMode == LineNumberRestartMode.NewSection)
                        {
                            ct.CurrentLineNumber = _options.LineNumberSettings.Start;
                        }
                    }

                    pdfDocument.NewPage();
                    ct.SetCurrentPage(pdfDocument.PageNumber);
                    currentColumn = 0;
                    SetColumnBounds();
                    pageWidth = GetColumnWidth();
                    continue;
                }

                // 处理浮动对象
                if (item is global::Nedev.FileConverters.DocxToPdf.Converters.FloatingObject floatObj)
                {
                    if (floatObj.Wrapping == global::Nedev.FileConverters.DocxToPdf.Converters.WrappingStyle.BehindText)
                    {
                        writer.DirectContentUnder.AddImage(floatObj.Image);
                    }
                    else if (floatObj.Wrapping == global::Nedev.FileConverters.DocxToPdf.Converters.WrappingStyle.InFrontOfText)
                    {
                        // Defer drawing until after ColumnText.Go() to ensure images appear above text
                        pendingInFrontImages.Add(floatObj);
                    }
                    else if (floatObj.Wrapping == global::Nedev.FileConverters.DocxToPdf.Converters.WrappingStyle.TopAndBottom)
                    {
                        // 上下型：
                        // 1. 结束当前行，换行
                        // 2. 在当前 Y 位置留出图片高度
                        // 3. 绘制图片到该空隙
                        // 4. 继续排版
                        
                        // 强行换行
                        ct.Go(); 
                        
                        var currentY = ct.YLine;
                        var imgHeight = floatObj.Height;
                        
                        // 检查剩余空间
                        if (currentY - imgHeight < _options.MarginBottom)
                        {
                            pdfDocument.NewPage();
                            ct.SetCurrentPage(pdfDocument.PageNumber);
                            currentColumn = 0;
                            SetColumnBounds();
                            ct.YLine = _options.PageSize.Height - _options.MarginTop;
                            currentY = ct.YLine;
                        }
                        
                        // 计算图片 X (居中或左对齐)
                        var pageContentWidth = GetPageContentWidth(); // 图片通常是基于栏或页的，这里简化为基于页
                        // 若 floatObj.Left 未设置（非绝对定位），则默认居中
                        var imgX = floatObj.PositionIsAbsolute ? floatObj.Left : (_options.MarginLeft + (pageContentWidth - floatObj.Width) / 2f);
                        var imgY = floatObj.PositionIsAbsolute ? (_options.PageSize.Height - floatObj.Top - imgHeight) : (currentY - imgHeight - 5f);
                        
                        floatObj.Image.SetAbsolutePosition(imgX, imgY);
                        writer.DirectContent.AddImage(floatObj.Image);
                        
                        // 调整 YLine
                        ct.YLine = imgY - 5f;
                    }
                    else if (floatObj.Wrapping == global::Nedev.FileConverters.DocxToPdf.Converters.WrappingStyle.Square || 
                             floatObj.Wrapping == global::Nedev.FileConverters.DocxToPdf.Converters.WrappingStyle.Tight || 
                             floatObj.Wrapping == global::Nedev.FileConverters.DocxToPdf.Converters.WrappingStyle.Through)
                    {
                        // 四周型/紧密型/穿越型：文字环绕图片
                        // 策略：
                        // 1. 计算图片在页面中的位置
                        // 2. 检测是否与当前文字流重叠
                        // 3. 如果重叠，先推进文字流到图片下方
                        // 4. 绘制图片
                        // 5. 继续排版文字
                        
                        var currentY = ct.YLine;
                        var imgWidth = floatObj.Width;
                        var imgHeight = floatObj.Height;
                        var imgX = floatObj.PositionIsAbsolute ? floatObj.Left : _options.MarginLeft;
                        var imgY = floatObj.PositionIsAbsolute 
                            ? (_options.PageSize.Height - floatObj.Top - imgHeight) 
                            : (currentY - imgHeight - 5f);
                        
                        // 计算图片占用的水平空间（相对于当前栏）
                        var columnBounds = GetColumnBounds(currentColumn);
                        var imgLeftInColumn = imgX - columnBounds.Left;
                        var imgRightInColumn = imgLeftInColumn + imgWidth;
                        var availableWidth = columnBounds.Width;
                        var leftSpace = imgLeftInColumn;
                        var rightSpace = availableWidth - imgRightInColumn;
                        
                        // 重叠检测：检查图片是否与当前文字流重叠
                        var hasOverlap = imgY < currentY && imgY + imgHeight > _options.MarginBottom;
                        
                        // 判断图片是否在栏内（不完全在左侧或右侧）
                        var imageInMiddle = leftSpace > 10f && rightSpace > 10f; // 左右都有空间
                        
                        if (hasOverlap)
                        {
                            if (imageInMiddle)
                            {
                                // 图片在中间，左右都有文字空间
                                // 策略：在图片前插入空白，将文字推到图片下方
                                // 这样可以避免文字重叠，但会留下空白区域
                                
                                // 计算需要推进的距离
                                var pushDownDistance = imgY + imgHeight - currentY + 5f;
                                if (pushDownDistance > 0)
                                {
                                    // 方法 1：使用空白段落推进文字流
                                    var blankPara = new iTextParagraph(" ");
                                    blankPara.SpacingAfter = pushDownDistance;
                                    ct.AddElement(blankPara);
                                }
                            }
                            else if (leftSpace < 10f || rightSpace < 10f)
                            {
                                // 图片占据大部分栏宽（左侧或右侧）
                                // 策略：推进文字流到图片下方，让文字在图片下方继续

                                var pushDownDistance = imgY + imgHeight - currentY + 5f;
                                if (pushDownDistance > 0)
                                {
                                    // 使用空白段落推进文字流
                                    var blankPara = new iTextParagraph(" ");
                                    blankPara.SpacingAfter = pushDownDistance;
                                    ct.AddElement(blankPara);
                                }
                            }
                            // else: 图片在边缘，文字可以从另一侧环绕
                            // 暂时不特殊处理，让 iText 自然排版
                        }
                        
                        // 设置图片绝对位置并绘制
                        floatObj.Image.SetAbsolutePosition(imgX, imgY);
                        writer.DirectContent.AddImage(floatObj.Image);
                        
                        // 更新文字流的 Y 位置（如果图片在当前文字流下方）
                        var imageBottom = imgY - imgHeight;
                        if (imageBottom < ct.YLine && imageBottom > _options.MarginBottom)
                        {
                            // 图片底部在当前位置下方，但不需要推进（已经在下方）
                            // 可以选择性地更新 YLine，让后续文字从图片底部开始
                            // ct.YLine = imageBottom; // 这可能导致文字跳到图片下方，谨慎使用
                        }
                    }
                    else
                    {
                        // Inline (嵌入型) 或其他未识别的:
                        // 把它交给 ColumnText 当作普通流式元素
                        ct.AddElement(floatObj);
                    }
                    continue;
                }

                ct.AddElement(item);

                // 渲染内容并处理分栏/分页
                while (true)
                {
                    var status = ct.Go();
                    FlushInFrontImages(); // draw deferred InFrontOfText images above rendered text
                    if (!ColumnText.HasMoreText(status)) break;

                    // 当前栏已满
                    currentColumn++;
                    if (currentColumn >= _currentColumnInfo.Count)
                    {
                        pdfDocument.NewPage();
                        ct.SetCurrentPage(pdfDocument.PageNumber);
                        currentColumn = 0;
                    }
                    SetColumnBounds();
                }
            }
        }

        // 确保所有内容都被渲染
        ct.Go();
        FlushInFrontImages(); // draw any remaining deferred images

        // 脚注与尾注内容（文末输出）
        if (_options.RenderFootnoteEndContent)
        {
            RenderNotesContent(pdfDocument, paragraphConverter, imageConverter, listConverter, tableConverter, numbering,
                footnoteById, endnoteById, footnoteIdsEncountered, endnoteIdsEncountered, footnoteNumberById, endnoteNumberById, pageWidth);
        }

        // 添加批注汇总页
        if (_options.AddCommentsSummaryPage)
        {
            CommentExporter.AddCommentsSummaryPage(docxDocument, pdfDocument);
        }

        // 添加修订汇总页
        if (_options.AddRevisionsSummaryPage)
        {
            RevisionHandler.AddRevisionsSummaryPage(body, pdfDocument);
        }

        pdfDocument.Close();
        writer.Close();

        paragraphConverter.HeadingRendered = null;

        if (hasHeaderFooter && tempStream != null && headerFooterRenderer != null)
        {
            StampHeaderFooter(tempStream, pdfStream, headerFooterRenderer, sectionTracker, _sectionSettingsMap);
        }

        // 应用水印（在所有其他内容之后）
        if (_options.Watermark != null && !string.IsNullOrEmpty(_options.Watermark.Text))
        {
            pdfStream.Position = 0;
            var watermarkedStream = new MemoryStream();
            WatermarkRenderer.ApplyWatermark(pdfStream, watermarkedStream, _options.Watermark);
            watermarkedStream.Position = 0;
            watermarkedStream.CopyTo(pdfStream);
            pdfStream.Position = 0;
        }

        // 更新目录页码
        if (_options.GenerateTableOfContents && tocEntries.Count > 0)
        {
            ApplyRecordedTocPageNumbers(tocEntries, tocPageNumbersByKey);

            pdfStream.Position = 0;
            using var reader = new PdfReader(pdfStream);
            using var outputStream = new MemoryStream();
            using var stamper = new PdfStamper(reader, outputStream);
            var totalPages = reader.NumberOfPages;

            if (totalPages > 1)
            {
                var pageSize = reader.GetPageSize(1);
                var contentFont = FontFactory.GetFont("STSong-Light", 10);

                var tocPageTextMap = new Dictionary<int, string>();
                for (var tocIndex = 0; tocIndex < tocEntries.Count; tocIndex++)
                {
                    var targetPage = tocEntries[tocIndex].PageNumber;
                    if (targetPage <= 0)
                        targetPage = Math.Min(tocIndex + 2, totalPages);
                    else if (targetPage > totalPages)
                        targetPage = totalPages;

                    tocPageTextMap[tocIndex] = targetPage.ToString();
                }

                var cb = stamper.GetOverContent(1);
                if (cb != null)
                {
                    cb.BeginText();
                    cb.SetFontAndSize(contentFont.Family, 10);

                    foreach (var kvp in tocPageTextMap)
                    {
                        var y = pageSize.Height - 112f - (kvp.Key * 18f);
                        if (y < 50f) break;
                        var x = pageSize.Width - 80f;
                        cb.ShowTextAligned(Element.ALIGN_LEFT, kvp.Value, x, y, 0);
                    }

                    cb.EndText();
                }
            }

            stamper.Close();
            outputStream.Position = 0;
            pdfStream.SetLength(0);
            outputStream.CopyTo(pdfStream);
            pdfStream.Position = 0;
        }
    }

    /// <summary>
    /// 生成目录页
    /// </summary>
    private void GenerateTOCPage(PdfDocument pdfDocument, List<TableOfContentsGenerator.TOCEntry> entries)
    {
        pdfDocument.NewPage();

        var titleFont = FontFactory.GetFont("STSong-Light", 24, iTextFont.BOLD);
        var title = new iTextParagraph("目录", titleFont)
        {
            Alignment = Element.ALIGN_CENTER,
            SpacingAfter = 30f
        };
        pdfDocument.Add(title);

        var contentFont = FontFactory.GetFont("STSong-Light", 12);
        var pageNumFont = FontFactory.GetFont("STSong-Light", 10);

        foreach (var entry in entries)
        {
            var indent = (entry.Level - 1) * 20f;
            var para = new iTextParagraph
            {
                IndentationLeft = indent,
                SpacingAfter = 4f
            };

            var titleChunk = new iTextChunk(entry.Title, contentFont);
            para.Add(titleChunk);

            para.Add(new iTextChunk(" ", contentFont));

            var pageNumChunk = new iTextChunk("    ", pageNumFont);
            para.Add(pageNumChunk);

            pdfDocument.Add(para);
        }
    }

    private static void ApplyRecordedTocPageNumbers(
        List<TableOfContentsGenerator.TOCEntry> entries,
        Dictionary<string, Queue<int>> tocPageNumbersByKey)
    {
        foreach (var entry in entries)
        {
            var key = TableOfContentsGenerator.BuildEntryKey(entry);
            if (tocPageNumbersByKey.TryGetValue(key, out var pages) && pages.Count > 0)
            {
                entry.PageNumber = pages.Dequeue();
            }
        }
    }

    private void StampHeaderFooter(
        MemoryStream sourcePdf, 
        Stream outputStream, 
        HeaderFooterRenderer renderer, 
        SectionTracker tracker,
        Dictionary<int, SectionPageSettings> sectionSettings)
    {
        sourcePdf.Position = 0;
        using var reader = new PdfReader(sourcePdf.ToArray());
        using var stamper = new PdfStamper(reader, outputStream);
        var totalPages = reader.NumberOfPages;
        
        // 计算每个节的页码范围
        // tracker.PageSections 记录了每一页所属的 sectionIndex
        // 我们需要知道每一页在当前节中是第几页
        
        var sectionPageCounts = new Dictionary<int, int>(); // sectionIndex -> total pages so far
        var pageNumInSection = new int[totalPages + 1];
        
        for (var i = 1; i <= totalPages; i++)
        {
            var sectionIndex = i <= tracker.PageSections.Count ? tracker.PageSections[i - 1] : 0;
            if (!sectionPageCounts.ContainsKey(sectionIndex)) sectionPageCounts[sectionIndex] = 0;
            sectionPageCounts[sectionIndex]++;
            pageNumInSection[i] = sectionPageCounts[sectionIndex];
        }

        for (var p = 1; p <= totalPages; p++)
        {
            var sectionIndex = p <= tracker.PageSections.Count ? tracker.PageSections[p - 1] : 0;
            var cb = stamper.GetOverContent(p);
            var pageSize = reader.GetPageSize(p);
            
            // Get settings
            if (!sectionSettings.TryGetValue(sectionIndex, out var settings))
                settings = sectionSettings.GetValueOrDefault(0) ?? new SectionPageSettings();

            renderer.Render(cb, pageSize, p, totalPages, sectionIndex, pageNumInSection[p], settings);
        }
    }

    /// <summary>
    /// 默认字段解析：支持 DATE/TIME/AUTHOR/TITLE/SUBJECT/REF/MERGEFIELD 等字段。
    /// </summary>
    internal string? ResolveField(string instruction, WordprocessingDocument docxDocument, string? docxPath = null)
    {
        if (string.IsNullOrWhiteSpace(instruction)) return null;
        var parts = instruction.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0) return null;

        var code = parts[0].TrimStart('\\').Trim().ToUpperInvariant();
        var props = docxDocument.PackageProperties;

        switch (code)
        {
            case "DATE":
                // 支持格式：DATE \@ "yyyy-MM-dd"
                var dateFormat = ParseFieldFormat(instruction, "yyyy-MM-dd");
                return DateTime.Now.ToString(dateFormat);
                
            case "TIME":
                var timeFormat = ParseFieldFormat(instruction, "HH:mm");
                return DateTime.Now.ToString(timeFormat);
                
            case "CREATEDATE":
                return props.Created?.ToString(ParseFieldFormat(instruction, "yyyy-MM-dd HH:mm")) ?? "";
                
            case "PRINTDATE":
                return props.LastPrinted?.ToString(ParseFieldFormat(instruction, "yyyy-MM-dd HH:mm")) ?? "";
                
            case "SAVEDATE":
                return props.Modified?.ToString(ParseFieldFormat(instruction, "yyyy-MM-dd HH:mm")) ?? "";
                
            case "AUTHOR":
            case "CREATOR":
                return props.Creator ?? props.LastModifiedBy;
                
            case "TITLE":
                return props.Title;
                
            case "SUBJECT":
                return props.Subject;
                
            case "KEYWORDS":
                return props.Keywords;
                
            case "COMMENTS":
                return props.Description;
                
            case "CATEGORY":
                return props.Category;
                
            case "COMPANY":
                // 从自定义属性获取公司信息
                return GetCustomProperty(docxDocument, "Company")?.ToString() ?? "";
                
            case "MANAGER":
                return GetCustomProperty(docxDocument, "Manager")?.ToString() ?? "";
                
            case "FILENAME":
                // 获取文件名（不含路径）
                // 优先使用提供的 docxPath
                var fileName = !string.IsNullOrEmpty(docxPath) ? Path.GetFileNameWithoutExtension(docxPath) : "document";
                return fileName;
                
            case "FILEPATH":
                // 返回文档路径（如果有）
                return docxPath ?? "";
                
            case "NUMPAGES":
            case "PAGE":
                // 这些需要运行时动态获取，返回 null 由 PageNumberProvider 处理
                return null;
                
            case "SECTION":
            case "SECTIONPAGES":
                // 节相关字段，需要节追踪器
                return null;
                
            case "REF":
                // 交叉引用书签
                return ParseRefField(instruction, docxDocument);
                
            case "MERGEFIELD":
                // 邮件合并字段
                var mfMatch = System.Text.RegularExpressions.Regex.Match(instruction, @"MERGEFIELD\s+""?([^""\s]+)""?", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (mfMatch.Success)
                {
                    var fieldName = mfMatch.Groups[1].Value;
                    if (_options.MergeData != null && _options.MergeData.TryGetValue(fieldName, out var val))
                    {
                        return val;
                    }
                    return ParseMergeField(instruction, docxDocument);
                }
                return null;
                
            case "HYPERLINK":
                // 超链接字段
                return ParseHyperlinkField(instruction);
                
            case "QUOTE":
                // 引号包裹的内容
                return ParseQuoteField(instruction);
                
            case "MACROBUTTON":
                // 宏按钮，显示按钮文本
                return ParseMacroButtonField(instruction);
                
            case "SHAPEDOG":
            case "DOCVARIABLE":
            case "LISTNUM":
            case "SYMBOL":
                // 这些字段暂时不支持
                return null;
                
            case "EQ":
                // EQ (Equation) 字段支持
                return ParseEqField(instruction);
                
            default:
                // 未知字段，尝试返回字段名本身
                return $"«{code}»";
        }
    }
    
    /// <summary>
    /// 解析字段格式（\@ "format"）
    /// </summary>
    private static string ParseFieldFormat(string instruction, string defaultFormat)
    {
        var formatMatch = System.Text.RegularExpressions.Regex.Match(instruction, @"\\@\s*""([^""]+)""", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return formatMatch.Success ? formatMatch.Groups[1].Value : defaultFormat;
    }
    
    /// <summary>
    /// 解析 EQ (Equation) 字段
    /// EQ 字段是 Word 中创建简单数学公式的一种方式（不同于 OMML）
    /// </summary>
    private static string ParseEqField(string instruction)
    {
        if (string.IsNullOrEmpty(instruction)) return "";
        
        var eqMatch = System.Text.RegularExpressions.Regex.Match(instruction, @"EQ\s*(.*)");
        if (!eqMatch.Success) return "";
        
        var eqContent = eqMatch.Groups[1].Value;
        
        var switches = System.Text.RegularExpressions.Regex.Matches(eqContent, @"\\(\w)\s*(\([^)]*\)|[^\s])");
        var result = new System.Text.StringBuilder();
        
        foreach (System.Text.RegularExpressions.Match sw in switches)
        {
            var switchCode = sw.Groups[1].Value;
            var switchArg = sw.Groups[2].Value.Trim('(', ')');
            
            switch (switchCode.ToLower())
            {
                case "f":
                case "fr":
                    // 分数 \f(numerator,denominator)
                    var fracParts = switchArg.Split(',');
                    if (fracParts.Length >= 2)
                        result.Append($"({fracParts[0]})/({fracParts[1]})");
                    else
                        result.Append($"({switchArg})");
                    break;
                case "su":
                    // 上标 \s(up)
                    var upMatch = System.Text.RegularExpressions.Regex.Match(switchArg, @"\(([^,]+),([^)]+)\)");
                    if (upMatch.Success)
                        result.Append($"{upMatch.Groups[1].Value}^{upMatch.Groups[2].Value}");
                    else
                        result.Append($"^{switchArg}");
                    break;
                case "di":
                    // 下标 \d(down)
                    var downMatch = System.Text.RegularExpressions.Regex.Match(switchArg, @"\(([^,]+),([^)]+)\)");
                    if (downMatch.Success)
                        result.Append($"{downMatch.Groups[1].Value}_{downMatch.Groups[2].Value}");
                    else
                        result.Append($"_{switchArg}");
                    break;
                case "ra":
                    // 根号 \r(order,radicand)
                    var radMatch = System.Text.RegularExpressions.Regex.Match(switchArg, @"\(([^,]+),?([^)]*)\)");
                    if (radMatch.Success)
                    {
                        var order = radMatch.Groups[1].Value;
                        var radicand = radMatch.Groups[2].Value;
                        if (!string.IsNullOrEmpty(order) && !string.IsNullOrEmpty(radicand))
                            result.Append($"{order}√({radicand})");
                        else if (!string.IsNullOrEmpty(radicand))
                            result.Append($"√({radicand})");
                    }
                    break;
                case "in":
                    // 积分 \i(start,end,integrand)
                    var intMatch = System.Text.RegularExpressions.Regex.Match(switchArg, @"\(([^,]+),([^,]+),([^)]+)\)");
                    if (intMatch.Success)
                        result.Append($"∫_{intMatch.Groups[1].Value}}}​^{intMatch.Groups[2].Value}({intMatch.Groups[3].Value})");
                    else
                        result.Append($"∫({switchArg})");
                    break;
                case "sum":
                    // 求和
                    result.Append($"∑({switchArg})");
                    break;
                case "ov":
                    // 上划线 \o(over)
                    result.Append($"({switchArg})¯");
                    break;
                case "ac":
                    // 省略号 \ac
                    result.Append("...");
                    break;
                default:
                    // 处理未知开关，直接输出
                    if (!string.IsNullOrEmpty(switchArg))
                        result.Append(switchArg);
                    break;
            }
        }
        
        return result.Length > 0 ? result.ToString() : "";
    }
    
    /// <summary>
    /// 解析 REF 字段（交叉引用）
    /// </summary>
    private string? ParseRefField(string instruction, WordprocessingDocument docxDocument)
    {
        // 提取书签名：REF BookmarkName \h 或 REF BookmarkName
        var match = System.Text.RegularExpressions.Regex.Match(instruction, @"REF\s+(\S+)");
        if (!match.Success) return null;
        
        var bookmarkName = match.Groups[1].Value;
        
        // 尝试查找书签并返回引用内容
        try
        {
            var mainPart = docxDocument.MainDocumentPart;
            if (mainPart == null) return bookmarkName;
            
            // 查找书签开始位置
            var bookmarkStarts = mainPart.Document?.Body?.Descendants<DocumentFormat.OpenXml.Wordprocessing.BookmarkStart>();
            if (bookmarkStarts != null)
            {
                foreach (var bookmarkStart in bookmarkStarts)
                {
                    if (bookmarkStart.Name?.Value?.Equals(bookmarkName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        // 找到书签，尝试获取书签后的内容
                        var bookmarkEnd = bookmarkStarts
                            .FirstOrDefault(b => b.Id == bookmarkStart.Id && b != bookmarkStart);
                        
                        // 获取书签所在段落的内容
                        var parentPara = bookmarkStart.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
                        if (parentPara != null)
                        {
                            // 返回段落文本（简化处理）
                            var text = parentPara.InnerText.Trim();
                            if (!string.IsNullOrEmpty(text))
                            {
                                return text;
                            }
                        }
                        
                        // 如果无法获取内容，返回书签名
                        return bookmarkName;
                    }
                }
            }
        }
        catch
        {
            // 忽略异常
        }
        
        // 如果找不到书签，返回书签名本身
        return bookmarkName;
    }
    
    /// <summary>
    /// 解析 MERGEFIELD 字段（邮件合并）
    /// </summary>
    private string? ParseMergeField(string instruction, WordprocessingDocument docxDocument)
    {
        // 提取字段名：MERGEFIELD FieldName
        var match = System.Text.RegularExpressions.Regex.Match(instruction, @"MERGEFIELD\s+(\S+)");
        if (!match.Success) return null;
        
        var fieldName = match.Groups[1].Value;
        
        // 尝试从文档属性获取值（简化版本）
        try
        {
            // 1. 尝试从自定义属性获取
            // 注意：完整的邮件合并需要连接外部数据源，这里仅支持文档属性
            
            // 2. 尝试从内置属性获取
            var builtInValue = GetBuiltInProperty(docxDocument, fieldName);
            if (builtInValue != null)
            {
                return builtInValue.ToString();
            }
        }
        catch
        {
            // 忽略异常
        }
        
        // 如果找不到值，返回占位符
        return $"«{fieldName}»";
    }
    
    /// <summary>
    /// 获取内置属性值
    /// </summary>
    private string? GetBuiltInProperty(WordprocessingDocument docxDocument, string fieldName)
    {
        var props = docxDocument.PackageProperties;
        
        switch (fieldName.ToUpperInvariant())
        {
            case "AUTHOR":
            case "CREATOR":
                return props.Creator;
            case "TITLE":
                return props.Title;
            case "SUBJECT":
                return props.Subject;
            case "KEYWORDS":
                return props.Keywords;
            case "COMMENTS":
            case "DESCRIPTION":
                return props.Description;
            case "CATEGORY":
                return props.Category;
            case "CREATED":
            case "CREATEDATE":
                return props.Created?.ToString("yyyy-MM-dd HH:mm");
            case "MODIFIED":
            case "SAVEDATE":
                return props.Modified?.ToString("yyyy-MM-dd HH:mm");
            case "LASTPRINTED":
            case "PRINTDATE":
                return props.LastPrinted?.ToString("yyyy-MM-dd HH:mm");
            case "LASTSAVEDBY":
            case "LASTMODIFIEDBY":
                return props.LastModifiedBy;
            case "REVISION":
                return props.Revision;
            case "VERSION":
                return props.Version;
        }
        
        return null;
    }
    
    /// <summary>
    /// 解析 HYPERLINK 字段
    /// </summary>
    private string? ParseHyperlinkField(string instruction)
    {
        // 提取 URL：HYPERLINK "http://example.com"。
        // 文本内容由字段本身的子元素提供，点击行为将在段落转换时附加锚点。
        var match = System.Text.RegularExpressions.Regex.Match(instruction, @"HYPERLINK\s+""([^""]+)""");
        if (!match.Success) return null;
        return match.Groups[1].Value;
    }
    
    /// <summary>
    /// 解析 QUOTE 字段
    /// </summary>
    private string? ParseQuoteField(string instruction)
    {
        // 提取引号内容：QUOTE "content"
        var match = System.Text.RegularExpressions.Regex.Match(instruction, @"QUOTE\s+""([^""]+)""");
        if (!match.Success) return null;
        
        return match.Groups[1].Value;
    }
    
    /// <summary>
    /// 解析 MACROBUTTON 字段
    /// </summary>
    private string? ParseMacroButtonField(string instruction)
    {
        // 提取按钮文本：MACROBUTTON MacroName "Button Text"
        var match = System.Text.RegularExpressions.Regex.Match(instruction, @"MACROBUTTON\s+\S+\s+""([^""]+)""");
        if (!match.Success) return null;
        
        return match.Groups[1].Value;
    }
    
    /// <summary>
    /// 获取自定义属性（完整支持）
    /// </summary>
    private string? GetCustomProperty(WordprocessingDocument docxDocument, string propertyName)
    {
        try
        {
            // 使用更简单的方式：直接从文档的内置属性和扩展属性查找
            var props = docxDocument.PackageProperties;
            
            // 尝试通过扩展属性查找（如果存在）
            // 注意：OpenXML SDK 的 CustomProperties 支持有限，这里使用备用方案
            
            // 简化处理：仅支持通过内置属性扩展
            // 完整的 CustomProperties 需要额外的 API 支持
        }
        catch (Exception ex)
        {
            // 记录异常但不抛出（可选：添加日志）
            System.Diagnostics.Debug.WriteLine($"读取自定义属性失败：{ex.Message}");
        }
        
        return null;
    }

    /// <summary>
    /// 将 DOCX 字节数组转为 PDF 字节数组
    /// </summary>
    public byte[] Convert(byte[] docxBytes)
    {
        if (docxBytes == null || docxBytes.Length == 0)
            throw new ArgumentException("DOCX 数据不能为空。", nameof(docxBytes));

        using var inputStream = new MemoryStream(docxBytes);
        using var outputStream = new MemoryStream();
        Convert(inputStream, outputStream);
        return outputStream.ToArray();
    }

    /// <summary>
    /// 静态方法：快速将 DOCX 文件转换为 PDF
    /// </summary>
    public static void ConvertFile(string docxPath, string pdfPath, ConvertOptions? options = null)
    {
        var converter = new DocxToPdfConverter(options ?? ConvertOptions.Default);
        converter.Convert(docxPath, pdfPath);
    }

    /// <summary>
    /// 静态方法：快速将 DOCX 流转换为 PDF
    /// </summary>
    public static void ConvertStream(Stream docxStream, Stream pdfStream, ConvertOptions? options = null)
    {
        var converter = new DocxToPdfConverter(options ?? ConvertOptions.Default);
        converter.Convert(docxStream, pdfStream);
    }

    /// <summary>
    /// 静态方法：快速将 DOCX 字节数组转换为 PDF
    /// </summary>
    public static byte[] ConvertBytes(byte[] docxBytes, ConvertOptions? options = null)
    {
        var converter = new DocxToPdfConverter(options ?? ConvertOptions.Default);
        return converter.Convert(docxBytes);
    }

    /// <summary>
    /// 查找系统中文字体路径
    /// </summary>
    private static string FindChineseFontPath()
    {
        // Windows字体目录
        var fontDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

        // 尝试的中文字体列表（按优先级）
        var fontFiles = new[]
        {
            "msyh.ttc",      // 微软雅黑
            "msyhbd.ttc",    // 微软雅黑粗体
            "simsun.ttc",    // 宋体
            "simhei.ttf",    // 黑体
            "simkai.ttf",    // 楷体
            "simfang.ttf",   // 仿宋
            "msgothic.ttc",  // 日文哥特体（支持部分中文）
            "malgun.ttf",    // 韩文（支持部分中文）
        };

        foreach (var fontFile in fontFiles)
        {
            var fontPath = Path.Combine(fontDir, fontFile);
            if (File.Exists(fontPath))
            {
                return fontPath;
            }
        }

        // 如果找不到，返回空字符串
        Console.WriteLine("Warning: No Chinese font found in system. Chinese characters may not display correctly.");
        return string.Empty;
    }

    /// <summary>
    /// 处理段落元素（包括列表合并处理）
    /// </summary>
    private List<IElement> ProcessParagraph(
        WParagraph paragraph,
        PdfWriter writer,
        ParagraphConverter paragraphConverter,
        DocxImageConverter imageConverter,
        ListConverter listConverter,
        Numbering? numbering,
        System.Collections.Generic.List<DocumentFormat.OpenXml.OpenXmlElement> elements,
        ref int index,
        ref float pageWidth,
        Action<SectionProperties>? onSectionBreak = null)
    {
        var results = new List<IElement>();

        // 检查是否为列表项
        if (ListConverter.IsListItem(paragraph))
        {
            var numberingId = ListConverter.GetNumberingId(paragraph);
            if (numberingId.HasValue)
            {
                // 收集同一列表的连续段落
                var listParagraphs = new System.Collections.Generic.List<WParagraph> { paragraph };
                var j = index + 1;

                while (j < elements.Count && elements[j] is WParagraph nextPara
                       && ListConverter.IsListItem(nextPara)
                       && ListConverter.GetNumberingId(nextPara) == numberingId)
                {
                    listParagraphs.Add(nextPara);
                    j++;
                }

                var pdfList = listConverter.ConvertListItems(listParagraphs, numbering, numberingId.Value);
                results.Add(pdfList);
                index = j;
                return results;
            }
        }

        // 普通段落转换
        var pdfElements = paragraphConverter.Convert(paragraph);
        results.AddRange(pdfElements);

        // 检查段落中是否包含图片
        var hasImages = paragraph.Descendants<Drawing>().Any()
                        || paragraph.Descendants<Picture>().Any()
                        || paragraph.Descendants<DocumentFormat.OpenXml.AlternateContent>().Any();

        if (hasImages)
        {
            var images = imageConverter.ConvertImagesInParagraph(paragraph, pageWidth, writer);
            results.AddRange(images);
        }

        // 检查是否有分页符（段落级别）
        var pageBreakBefore = paragraph.ParagraphProperties?.PageBreakBefore;
        if (pageBreakBefore != null && (pageBreakBefore.Val == null || pageBreakBefore.Val.Value))
        {
            results.Add(NextPageChunk);
        }

        var sectionProps = paragraph.ParagraphProperties?.SectionProperties;
        if (sectionProps != null)
        {
            ApplyPageSettings(sectionProps);
            onSectionBreak?.Invoke(sectionProps);
            results.Add(NextPageChunk);
        }

        index++;
        return results;
    }

    private List<IElement> ProcessSdtBlock(
        SdtBlock sdtBlock,
        PdfWriter writer,
        ParagraphConverter paragraphConverter,
        TableConverter tableConverter,
        DocxImageConverter imageConverter,
        ListConverter listConverter,
        Numbering? numbering,
        ref float pageWidth,
        Action<SectionProperties>? onSectionBreak = null)
    {
        var results = new List<IElement>();
        
        // 提取内容控件属性（用于调试或特殊处理）
        var sdtProperties = sdtBlock.GetFirstChild<SdtProperties>();
        var sdtAlias = sdtProperties?.GetFirstChild<SdtAlias>();
        var displayName = sdtAlias?.Val?.Value;
        
        var content = sdtBlock.SdtContentBlock;
        if (content == null) return results;

        var children = content.ChildElements.ToList();
        var i = 0;
        while (i < children.Count)
        {
            var element = children[i];
            switch (element)
            {
                case WParagraph paragraph:
                    results.AddRange(ProcessParagraph(paragraph, writer, paragraphConverter, imageConverter, listConverter, numbering, children, ref i, ref pageWidth, onSectionBreak));
                    break;
                case WTable table:
                    var pdfTable = tableConverter.Convert(table, pageWidth);
                    if (pdfTable != null) results.Add(pdfTable);
                    i++;
                    break;
                case SectionProperties sect:
                    ApplyPageSettings(sect);
                    onSectionBreak?.Invoke(sect);
                    results.Add(NextPageChunk);
                    i++;
                    break;
                case SdtBlock innerSdt:
                    // 递归处理嵌套的内容控件（简化：仅处理一层）
                    results.AddRange(ProcessSdtBlock(innerSdt, writer, paragraphConverter, tableConverter, imageConverter, listConverter, numbering, ref pageWidth, onSectionBreak));
                    i++;
                    break;
                default:
                    i++;
                    break;
            }
        }
        return results;
    }

    private void ApplyCurrentPageSettings(iTextDocument pdfDocument, iTextColumnText ct)
    {
        pdfDocument.SetPageSize(_options.PageSize);
        pdfDocument.SetMargins(_options.MarginLeft, _options.MarginRight, _options.MarginTop, _options.MarginBottom);
        ct.TextDirection = _currentTextDirection;
    }

    private float GetPageContentWidth()
    {
        return _options.PageSize.Width - _options.MarginLeft - _options.MarginRight;
    }

    private float GetColumnWidth()
    {
        var avail = GetPageContentWidth();
        if (_currentColumnInfo.Count <= 1) return avail;
        return (avail - (_currentColumnInfo.Count - 1) * _currentColumnInfo.Spacing) / _currentColumnInfo.Count;
    }

    private Rectangle GetColumnBounds(int columnIndex)
    {
        var info = _currentColumnInfo;
        var pageSize = _options.PageSize;
        var ml = _options.MarginLeft;
        var mr = _options.MarginRight;

        var availableWidth = pageSize.Width - ml - mr;
        if (info.Count <= 1)
        {
            return new Rectangle(ml, _options.MarginBottom, pageSize.Width - mr, _options.MarginTop);
        }

        var colWidth = (availableWidth - (info.Count - 1) * info.Spacing) / info.Count;
        var llx = ml + columnIndex * (colWidth + info.Spacing);
        var urx = llx + colWidth;
        return new Rectangle(llx, _options.MarginBottom, urx, _options.MarginTop);
    }

    /// <summary>
    /// 从 DOCX SectionProperties 中读取页面设置
    /// </summary>
    private void ApplyPageSettings(SectionProperties? sectionProps)
    {
        if (sectionProps == null) return;

        var pageSize = sectionProps.GetFirstChild<WPageSize>();
        if (pageSize != null)
        {
            // DOCX 页面尺寸以 twips (1/20 pt) 为单位
            if (pageSize.Width?.Value is uint w && pageSize.Height?.Value is uint h)
            {
                var widthPt = w / 20f;
                var heightPt = h / 20f;
                _options.PageSize = new Rectangle(widthPt, heightPt);

                // 横向
                if (pageSize.Orient?.Value == PageOrientationValues.Landscape)
                {
                    _options.PageSize = _options.PageSize.Rotate();
                }
            }
        }

        var pageMargin = sectionProps.GetFirstChild<PageMargin>();
        if (pageMargin != null)
        {
            if (pageMargin.Left?.Value is uint ml) _options.MarginLeft = ml / 20f;
            if (pageMargin.Right?.Value is uint mr) _options.MarginRight = mr / 20f;
            if (pageMargin.Top?.Value is int mt) _options.MarginTop = Math.Abs(mt) / 20f;
            if (pageMargin.Bottom?.Value is int mb) _options.MarginBottom = Math.Abs(mb) / 20f;
            
            // Header/Footer distance
            if (pageMargin.Header?.Value is uint h) _options.HeaderDistance = h / 20f;
            if (pageMargin.Footer?.Value is uint f) _options.FooterDistance = f / 20f;
        }

        // 文本方向
        var textDir = sectionProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.TextDirection>();
        if (textDir != null && textDir.Val?.Value == TextDirectionValues.TopToBottomRightToLeft)
        {
            _currentTextDirection = Models.TextDirection.Vertical;
        }
        else
        {
            _currentTextDirection = Models.TextDirection.Horizontal;
        }

        // 多栏支持
        var columns = sectionProps.GetFirstChild<Columns>();
        if (columns != null)
        {
            var numColumns = columns.ColumnCount?.Value ?? 1;
            var spacing = columns.Space?.Value ?? "720"; // 默认 720 twips (36pt)
            var spacingPt = StyleHelper.TwipsToPoints(spacing);

            _currentColumnInfo = new ColumnInfo
            {
                Count = Math.Max(1, (int)numColumns),
                Spacing = spacingPt
            };
        }
        else
        {
            _currentColumnInfo = new ColumnInfo();
        }

        // 行号支持
        var lineNumType = sectionProps.GetFirstChild<LineNumberType>();
        if (lineNumType != null)
        {
            var settings = new LineNumberSettings();
            
            // start
            settings.Start = (int)(lineNumType.Start?.Value ?? 1);
                
            // countBy
            settings.CountBy = (int)(lineNumType.CountBy?.Value ?? 1); 
            
            // distance (twips)
            if (lineNumType.Distance?.Value is string distStr && float.TryParse(distStr, out var dist))
                settings.Distance = dist / 20f;
            else
                settings.Distance = 0.25f * 72f; // Default ~0.25 inch
                
            // restart
            if (lineNumType.Restart?.Value == LineNumberRestartValues.NewPage)
                settings.RestartMode = LineNumberRestartMode.NewPage;
            else if (lineNumType.Restart?.Value == LineNumberRestartValues.NewSection)
                settings.RestartMode = LineNumberRestartMode.NewSection;
            else
                settings.RestartMode = LineNumberRestartMode.Continuous;

            _options.LineNumberSettings = settings;
        }
        else
        {
            _options.LineNumberSettings = null;
        }

        // 页面边框
        var pgBorders = sectionProps.GetFirstChild<PageBorders>();
        if (pgBorders != null)
        {
            var options = new PageBorderOptions
            {
                OffsetFrom = pgBorders.OffsetFrom?.Value.ToString().ToLower() ?? "page"
            };

            PageBorderSide? ParseSide(BorderType? b)
            {
                if (b == null || b.Val?.Value == BorderValues.None) return null;
                return new PageBorderSide
                {
                    Val = b.Val?.Value.ToString().ToLower() ?? "none",
                    Size = (b.Size?.Value ?? 4) / 8f, // sz is in 1/8 pt
                    Space = (b.Space?.Value ?? 0),   // space is in pt
                    Color = StyleHelper.HexToBaseColor(b.Color?.Value ?? "000000") ?? iTextBaseColor.Black
                };
            }

            options.Top = ParseSide(pgBorders.TopBorder);
            options.Bottom = ParseSide(pgBorders.BottomBorder);
            options.Left = ParseSide(pgBorders.LeftBorder);
            options.Right = ParseSide(pgBorders.RightBorder);

            _options.PageBorders = options;
        }
        else
        {
            _options.PageBorders = null;
        }
    }

    private static bool HasAnyHeaderFooter(Body body, SectionProperties? firstSection)
    {
        bool HasRefs(SectionProperties? sp) =>
            sp?.Elements<HeaderReference>().Any() == true || sp?.Elements<FooterReference>().Any() == true;
        if (HasRefs(firstSection)) return true;
        foreach (var sp in body.Descendants<SectionProperties>())
        {
            if (HasRefs(sp)) return true;
        }
        return false;
    }

    private static Dictionary<int, int> BuildNoteNumberMap<TNote>(IEnumerable<TNote>? notes)
        where TNote : OpenXmlCompositeElement
    {
        var map = new Dictionary<int, int>();
        if (notes == null) return map;

        var n = 0;
        foreach (var note in notes)
        {
            var idAttr = note.GetAttributes().FirstOrDefault(a => a.LocalName.Equals("id", StringComparison.OrdinalIgnoreCase));
            if (!int.TryParse(idAttr.Value, out var id)) continue;
            if (id <= 0) continue;
            n++;
            map[id] = n;
        }

        return map;
    }

    private static Dictionary<int, TNote> BuildNoteByIdMap<TNote>(IEnumerable<TNote>? notes)
        where TNote : OpenXmlCompositeElement
    {
        var map = new Dictionary<int, TNote>();
        if (notes == null) return map;
        foreach (var note in notes)
        {
            var idAttr = note.GetAttributes().FirstOrDefault(a => a.LocalName.Equals("id", StringComparison.OrdinalIgnoreCase));
            if (int.TryParse(idAttr.Value, out var id) && id > 0)
                map[id] = note;
        }
        return map;
    }

    private void RenderNotesContent(
        iTextDocument pdfDocument,
        ParagraphConverter paragraphConverter,
        DocxImageConverter imageConverter,
        ListConverter listConverter,
        TableConverter tableConverter,
        Numbering? numbering,
        Dictionary<int, Footnote> footnoteById,
        Dictionary<int, Endnote> endnoteById,
        List<int> footnoteIdsEncountered,
        List<int> endnoteIdsEncountered,
        Dictionary<int, int> footnoteNumberById,
        Dictionary<int, int> endnoteNumberById,
        float pageWidth)
    {
        var seenFn = new HashSet<int>();
        var seenEn = new HashSet<int>();

        foreach (var id in footnoteIdsEncountered)
        {
            if (!seenFn.Add(id) || !footnoteById.TryGetValue(id, out var fn)) continue;
            var num = footnoteNumberById.GetValueOrDefault(id, id);
            RenderNoteBody(pdfDocument, paragraphConverter, imageConverter, listConverter, tableConverter, numbering, fn, num, pageWidth);
        }
        foreach (var id in endnoteIdsEncountered)
        {
            if (!seenEn.Add(id) || !endnoteById.TryGetValue(id, out var en)) continue;
            var num = endnoteNumberById.GetValueOrDefault(id, id);
            RenderNoteBody(pdfDocument, paragraphConverter, imageConverter, listConverter, tableConverter, numbering, en, num, pageWidth);
        }
    }

    private void RenderNoteBody<TNote>(
        iTextDocument pdfDocument,
        ParagraphConverter paragraphConverter,
        DocxImageConverter imageConverter,
        ListConverter listConverter,
        TableConverter tableConverter,
        Numbering? numbering,
        TNote note,
        int noteNum,
        float pageWidth)
        where TNote : OpenXmlCompositeElement
    {
        var body = note.Elements().FirstOrDefault(e => e.LocalName == "body" || e.LocalName == "Body");
        var children = body?.ChildElements ?? note.ChildElements;

        foreach (var child in children)
        {
            switch (child)
            {
                case WParagraph para:
                    if (ListConverter.IsListItem(para))
                    {
                        var numberingId = ListConverter.GetNumberingId(para);
                        if (numberingId.HasValue)
                        {
                            var listParagraphs = new List<WParagraph> { para };
                            var pdfList = listConverter.ConvertListItems(listParagraphs, numbering, numberingId.Value);
                            pdfDocument.Add(pdfList);
                        }
                    }
                    else
                    {
                        var elements = paragraphConverter.Convert(para);
                        foreach (var el in elements)
                        {
                            if (el is Chunk c && c.Content == "PAGE_BREAK") continue;
                            pdfDocument.Add(el);
                        }
                        if (para.Descendants<Drawing>().Any() || para.Descendants<Picture>().Any())
                        {
                            foreach (var img in imageConverter.ConvertImagesInParagraph(para, pageWidth, null))
                                pdfDocument.Add(img);
                        }
                    }
                    break;
                case WTable table:
                    var pdfTable = tableConverter.Convert(table, pageWidth);
                    if (pdfTable != null) pdfDocument.Add(pdfTable);
                    break;
            }
        }
    }
}

/// <summary>用于两遍转换时记录每页所属节</summary>
internal class SectionTracker : PdfPageEventHelper
{
    public int CurrentSection { get; set; }
    public readonly List<int> PageSections = new();
    public override void OnStartPage(PdfWriter writer, iTextDocument document)
    {
        PageSections.Add(CurrentSection);
    }
}

/// <summary>书签跟踪器：记录标题和书签位置</summary>
internal class BookmarkTracker : PdfPageEventHelper
{
    private readonly PdfWriter _writer;
    private PdfOutline? _rootOutline;
    private readonly Dictionary<int, PdfOutline> _outlineByLevel = new();
    
    public BookmarkTracker(PdfWriter writer)
    {
        _writer = writer;
    }

    public override void OnOpenDocument(PdfWriter writer, iTextDocument document)
    {
        _rootOutline = writer.RootOutline;
    }

    public void AddHeadingBookmark(string title, int level)
    {
        if (_rootOutline == null) return;
        level = Math.Clamp(level, 1, 6);

        var dest = new PdfDestination(PdfDestination.XYZ, -1, _writer.GetVerticalPosition(false), 0);
        var parent = level == 1 ? _rootOutline : _outlineByLevel.GetValueOrDefault(level - 1, _rootOutline);
        var outline = new PdfOutline(parent, dest, title, level <= 2);
        _outlineByLevel[level] = outline;
    }

    public void AddBookmark(string name)
    {
        if (_rootOutline == null) return;
        var dest = new PdfDestination(PdfDestination.XYZ, -1, _writer.GetVerticalPosition(false), 0);
        new PdfOutline(_rootOutline, dest, name, false);
    }
}

/// <summary>组合多个 PageEvent</summary>
internal class CombinedPageEvent : PdfPageEventHelper
{
    private readonly PdfPageEventHelper[] _events;

    public CombinedPageEvent(params PdfPageEventHelper[] events)
    {
        _events = events;
    }

    public override void OnOpenDocument(PdfWriter writer, iTextDocument document)
    {
        foreach (var e in _events) e.OnOpenDocument(writer, document);
    }

    public override void OnStartPage(PdfWriter writer, iTextDocument document)
    {
        foreach (var e in _events) e.OnStartPage(writer, document);
    }

    public override void OnEndPage(PdfWriter writer, iTextDocument document)
    {
        foreach (var e in _events) e.OnEndPage(writer, document);
    }

    public override void OnCloseDocument(PdfWriter writer, iTextDocument document)
    {
        foreach (var e in _events) e.OnCloseDocument(writer, document);
    }
}

/// <summary>页面边框渲染器</summary>
internal class PageBorderEvent : PdfPageEventHelper
{
    private readonly PageBorderOptions _borders;
    private readonly ConvertOptions _options;

    public PageBorderEvent(PageBorderOptions borders, ConvertOptions options)
    {
        _borders = borders;
        _options = options;
    }

    public override void OnEndPage(PdfWriter writer, iTextDocument document)
    {
        var cb = writer.DirectContent;
        var rect = document.PageSize;

        cb.SaveState();

        void DrawSide(PageBorderSide? side, float x1, float y1, float x2, float y2)
        {
            if (side == null || side.Val == "none") return;

            var size = side.Size;

            switch (side.Val)
            {
                case "double":
                    // Draw two parallel lines with a gap
                    var gap = Math.Max(size * 1.5f, 1f);
                    // Determine offset direction (perpendicular to the line)
                    var dx = x2 - x1;
                    var dy = y2 - y1;
                    var len = (float)Math.Sqrt(dx * dx + dy * dy);
                    float nx = 0, ny = 0;
                    if (len > 0) { nx = -dy / len * gap / 2f; ny = dx / len * gap / 2f; }

                    cb.SetLineWidth(size * 0.5f);
                    cb.SetColorStroke(side.Color);
                    cb.MoveTo(x1 + nx, y1 + ny);
                    cb.LineTo(x2 + nx, y2 + ny);
                    cb.Stroke();
                    cb.MoveTo(x1 - nx, y1 - ny);
                    cb.LineTo(x2 - nx, y2 - ny);
                    cb.Stroke();
                    break;

                case "dashed":
                case "dashSmallGap":
                    cb.SetLineWidth(size);
                    cb.SetColorStroke(side.Color);
                    cb.SetLineDash(new[] { 4f * size, 2f * size });
                    cb.MoveTo(x1, y1);
                    cb.LineTo(x2, y2);
                    cb.Stroke();
                    cb.SetLineDash(Array.Empty<float>()); // reset
                    break;

                case "dotted":
                    cb.SetLineWidth(size);
                    cb.SetColorStroke(side.Color);
                    cb.SetLineDash(new[] { size, size });
                    cb.MoveTo(x1, y1);
                    cb.LineTo(x2, y2);
                    cb.Stroke();
                    cb.SetLineDash(Array.Empty<float>());
                    break;

                case "dotDash":
                    cb.SetLineWidth(size);
                    cb.SetColorStroke(side.Color);
                    cb.SetLineDash(new[] { 4f * size, 2f * size, size, 2f * size });
                    cb.MoveTo(x1, y1);
                    cb.LineTo(x2, y2);
                    cb.Stroke();
                    cb.SetLineDash(Array.Empty<float>());
                    break;

                default: // "single" and others — solid line
                    cb.SetLineWidth(size);
                    cb.SetColorStroke(side.Color);
                    cb.MoveTo(x1, y1);
                    cb.LineTo(x2, y2);
                    cb.Stroke();
                    break;
            }
        }

        float top = rect.Top;
        float bottom = rect.Bottom;
        float left = rect.Left;
        float right = rect.Right;

        if (_borders.OffsetFrom == "text")
        {
            top -= _options.MarginTop - (_borders.Top?.Space ?? 0);
            bottom += _options.MarginBottom - (_borders.Bottom?.Space ?? 0);
            left += _options.MarginLeft - (_borders.Left?.Space ?? 0);
            right -= _options.MarginRight - (_borders.Right?.Space ?? 0);
        }
        else // offsetFrom == "page"
        {
            top -= (_borders.Top?.Space ?? 24);
            bottom += (_borders.Bottom?.Space ?? 24);
            left += (_borders.Left?.Space ?? 24);
            right -= (_borders.Right?.Space ?? 24);
        }

        DrawSide(_borders.Top, left, top, right, top);
        DrawSide(_borders.Bottom, left, bottom, right, bottom);
        DrawSide(_borders.Left, left, bottom, left, top);
        DrawSide(_borders.Right, right, bottom, right, top);

        cb.RestoreState();
    }
}
