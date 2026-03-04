using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.DocxToPdf.Models;
using Nedev.DocxToPdf.PdfEngine;
using DocxImageConverter = Nedev.DocxToPdf.Converters.ImageConverter;
using iTextParagraph = Nedev.DocxToPdf.PdfEngine.Paragraph;

namespace Nedev.DocxToPdf.Converters;

/// <summary>
/// 在 PDF 每页上渲染页眉页脚
/// </summary>
internal class HeaderFooterRenderer
{
    private readonly MainDocumentPart _mainPart;
    private readonly ParagraphConverter _paragraphConverter;
    private readonly DocxImageConverter _imageConverter;
    private readonly ConvertOptions _options;
    private readonly float _pageWidth;
    private readonly Dictionary<int, SectionInfo> _sectionInfos = new();

    private class SectionInfo
    {
        public OpenXmlElement? HeaderDefault;
        public OpenXmlElement? HeaderFirst;
        public OpenXmlElement? HeaderEven;
        public OpenXmlElement? FooterDefault;
        public OpenXmlElement? FooterFirst;
        public OpenXmlElement? FooterEven;
        public bool TitlePage; // 首页不同
    }

    public HeaderFooterRenderer(
        MainDocumentPart mainPart,
        ParagraphConverter paragraphConverter,
        DocxImageConverter imageConverter,
        ConvertOptions options,
        float pageWidth)
    {
        _mainPart = mainPart;
        _paragraphConverter = paragraphConverter;
        _imageConverter = imageConverter;
        _options = options;
        _pageWidth = pageWidth;
    }

    /// <summary>
    /// 从 SectionProperties 提取当前节的页眉页脚引用并缓存
    /// </summary>
    public void RegisterSection(SectionProperties? sectionProps, int sectionIndex)
    {
        if (sectionProps == null) return;

        var info = new SectionInfo();
        info.TitlePage = sectionProps.GetFirstChild<TitlePage>()?.Val?.Value ?? false;

        foreach (var headerRef in sectionProps.Elements<HeaderReference>())
        {
            var id = headerRef.Id?.Value;
            var type = headerRef.Type?.Value;
            if (string.IsNullOrEmpty(id)) continue;
            var part = _mainPart.GetPartById(id) as HeaderPart;
            var headerBody = part?.Header;
            
            if (type == HeaderFooterValues.Default) info.HeaderDefault = headerBody;
            else if (type == HeaderFooterValues.First) info.HeaderFirst = headerBody;
            else if (type == HeaderFooterValues.Even) info.HeaderEven = headerBody;
        }

        foreach (var footerRef in sectionProps.Elements<FooterReference>())
        {
            var id = footerRef.Id?.Value;
            var type = footerRef.Type?.Value;
            if (string.IsNullOrEmpty(id)) continue;
            var part = _mainPart.GetPartById(id) as FooterPart;
            var footerBody = part?.Footer;

            if (type == HeaderFooterValues.Default) info.FooterDefault = footerBody;
            else if (type == HeaderFooterValues.First) info.FooterFirst = footerBody;
            else if (type == HeaderFooterValues.Even) info.FooterEven = footerBody;
        }

        _sectionInfos[sectionIndex] = info;
    }

    /// <summary>
    /// 获取指定节在特定页面类型下的页眉页脚（支持继承上一节）
    /// </summary>
    public (OpenXmlElement? Header, OpenXmlElement? Footer) GetForPage(int sectionIndex, int pageNumInTotal, int pageNumInSection)
    {
        // 查找当前节或最近的前驱节定义
        SectionInfo? info = null;
        for (var i = sectionIndex; i >= 0; i--)
        {
            if (_sectionInfos.TryGetValue(i, out info))
                break;
        }
        
        if (info == null) return (null, null);

        // 判断页面类型
        var isFirst = info.TitlePage && pageNumInSection == 1;
        var isEven = pageNumInTotal % 2 == 0;

        // 页眉选择逻辑：First -> Even -> Default
        // 注意：Word 中 Even 仅在 Settings.EvenAndOddHeaders 为 true 时生效，这里简化处理：只要有 Even 定义就优先使用
        OpenXmlElement? header = null;
        if (isFirst && info.HeaderFirst != null) header = info.HeaderFirst;
        else if (isEven && info.HeaderEven != null) header = info.HeaderEven;
        else header = info.HeaderDefault;

        // 页脚选择逻辑
        OpenXmlElement? footer = null;
        if (isFirst && info.FooterFirst != null) footer = info.FooterFirst;
        else if (isEven && info.FooterEven != null) footer = info.FooterEven;
        else footer = info.FooterDefault;

        // 如果当前节定义了某种类型但为 null，是否应该回退到上一节？
        // Word 规则：链接到前一节 (LinkToPrevious)。
        // OpenXML 中若无 HeaderReference，则默认继承。
        // 但若显式定义了 HeaderReference 但内容为空，则是空页眉。
        // 由于 RegisterSection 仅解析存在的引用，这里 info 中的 null 意味着未定义，应当继承上一节的对应类型。
        // 简化实现：如果在 info 中没找到，尝试向前查找有该类型定义的节。
        
        if (header == null) header = FindInheritedHeader(sectionIndex, isFirst, isEven);
        if (footer == null) footer = FindInheritedFooter(sectionIndex, isFirst, isEven);

        return (header, footer);
    }

    private OpenXmlElement? FindInheritedHeader(int startIndex, bool isFirst, bool isEven)
    {
        for (var i = startIndex; i >= 0; i--)
        {
            if (!_sectionInfos.TryGetValue(i, out var info)) continue;
            if (isFirst && info.HeaderFirst != null) return info.HeaderFirst;
            if (isEven && info.HeaderEven != null) return info.HeaderEven;
            if (!isFirst && !isEven && info.HeaderDefault != null) return info.HeaderDefault;
            
            // 如果不是 TitlePage，First 页眉回退到 Default？不，First 仅在 TitlePage=true 时生效。
            // 但如果上一节定义了 Default，当前页是 Default 类型，则继承。
        }
        return null;
    }

    private OpenXmlElement? FindInheritedFooter(int startIndex, bool isFirst, bool isEven)
    {
        for (var i = startIndex; i >= 0; i--)
        {
            if (!_sectionInfos.TryGetValue(i, out var info)) continue;
            if (isFirst && info.FooterFirst != null) return info.FooterFirst;
            if (isEven && info.FooterEven != null) return info.FooterEven;
            if (!isFirst && !isEven && info.FooterDefault != null) return info.FooterDefault;
        }
        return null;
    }

    /// <summary>
    /// 在指定页面上绘制页眉页脚（用于第二遍 PdfStamper 叠加）
    /// </summary>
    public void Render(PdfContentByte cb, Rectangle pageSize, int pageNum, int totalPages, int sectionIndex, int pageNumInSection)
    {
        var (headerBody, footerBody) = GetForPage(sectionIndex, pageNum, pageNumInSection);
        if (headerBody == null && footerBody == null) return;

        _paragraphConverter.PageNumberProvider = () => (pageNum, totalPages);

        var pageHeight = pageSize.Height;
        var pageWidth = pageSize.Width;
        var marginL = _options.MarginLeft;
        var marginR = _options.MarginRight;
        var marginT = _options.MarginTop;
        var marginB = _options.MarginBottom;

        if (headerBody != null)
        {
            var headerHeight = MeasureHeight(headerBody, pageWidth - marginL - marginR);
            RenderBody(cb, headerBody,
                marginL, pageHeight - marginT - headerHeight,
                pageWidth - marginR, pageHeight - marginT + 20f);
        }

        if (footerBody != null)
        {
            var footerHeight = MeasureHeight(footerBody, pageWidth - marginL - marginR);
            RenderBody(cb, footerBody,
                marginL, marginB - 20f,
                pageWidth - marginR, marginB + footerHeight);
        }
    }

    private float MeasureHeight(OpenXmlElement body, float width)
    {
        // 简单估算：基于内容行数
        try
        {
            var simCt = new ColumnText(new PdfContentByte());
            simCt.SetSimpleColumn(0, 0, width, 10000);

            AddBodyToColumnText(simCt, body, width);

            simCt.Go(true); // simulate
            var h = 10000 - simCt.YLine;
            return h < 20 ? 20 : h;
        }
        catch
        {
            return 40f; // fallback
        }
    }

    private void RenderBody(PdfContentByte cb, OpenXmlElement body,
        float llx, float lly, float urx, float ury)
    {
        var ct = new ColumnText(cb);
        ct.SetSimpleColumn(llx, lly, urx, ury);

        AddBodyToColumnText(ct, body, urx - llx);

        ct.Go();
    }

    private void AddBodyToColumnText(ColumnText ct, OpenXmlElement body, float width)
    {
        foreach (var child in body.ChildElements)
        {
            switch (child)
            {
                case DocumentFormat.OpenXml.Wordprocessing.Paragraph para:
                    var elements = _paragraphConverter.Convert(para);
                    foreach (var el in elements)
                    {
                        if (el is Chunk c && (c.Content == "PAGE_BREAK" || c.Content == "NEXTPAGE_SIGNAL")) continue;
                        if (el is iTextParagraph p)
                            ct.AddElement(p);
                        else if (el is PdfPTable tbl)
                            ct.AddElement(tbl);
                    }
                    break;
                case DocumentFormat.OpenXml.Wordprocessing.Table table:
                    if (TableConverter != null)
                    {
                        var pdfTable = TableConverter.Convert(table, width);
                        if (pdfTable != null)
                            ct.AddElement(pdfTable);
                    }
                    break;
            }
        }
    }

    /// <summary>
    /// 设置表格转换器（用于页眉页脚中的表格）
    /// </summary>
    public TableConverter? TableConverter { get; set; }
}
