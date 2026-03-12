using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Models;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using DocxImageConverter = Nedev.FileConverters.DocxToPdf.Converters.ImageConverter;
using iTextParagraph = Nedev.FileConverters.DocxToPdf.PdfEngine.Paragraph;

namespace Nedev.FileConverters.DocxToPdf.Converters;

/// <summary>
/// ? PDF ?????????
/// </summary>
public class HeaderFooterRenderer
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
        public bool TitlePage; // ????
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
    /// ? SectionProperties ???????????????
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
    /// ??????????????????(???????)
    /// </summary>
    public (OpenXmlElement? Header, OpenXmlElement? Footer) GetForPage(int sectionIndex, int pageNumInTotal, int pageNumInSection)
    {
        // ??????????????
        SectionInfo? info = null;
        for (var i = sectionIndex; i >= 0; i--)
        {
            if (_sectionInfos.TryGetValue(i, out info))
                break;
        }
        
        if (info == null) return (null, null);

        // ??????
        var isFirst = info.TitlePage && pageNumInSection == 1;
        var isEven = pageNumInTotal % 2 == 0;

        // ??????:First -> Even -> Default
        // ??:Word ? Even ?? Settings.EvenAndOddHeaders ? true ???,??????:??? Even ???????
        OpenXmlElement? header = null;
        if (isFirst && info.HeaderFirst != null) header = info.HeaderFirst;
        else if (isEven && info.HeaderEven != null) header = info.HeaderEven;
        else header = info.HeaderDefault;

        // ??????
        OpenXmlElement? footer = null;
        if (isFirst && info.FooterFirst != null) footer = info.FooterFirst;
        else if (isEven && info.FooterEven != null) footer = info.FooterEven;
        else footer = info.FooterDefault;

        // ?????????????? null,???????????
        // Word ??:?????? (LinkToPrevious)?
        // OpenXML ??? HeaderReference,??????
        // ??????? HeaderReference ?????,??????
        // ?? RegisterSection ????????,?? info ?? null ??????,?????????????
        // ????:??? info ????,???????????????
        
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
            
            // ???? TitlePage,First ????? Default??,First ?? TitlePage=true ????
            // ????????? Default,???? Default ??,????
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
    /// ????????????(????? PdfStamper ??)
    /// </summary>
    public void Render(PdfContentByte cb, Rectangle pageSize, int pageNum, int totalPages, int sectionIndex, int pageNumInSection, int totalPagesInSection, SectionPageSettings settings)
    {
        var (headerBody, footerBody) = GetForPage(sectionIndex, pageNum, pageNumInSection);
        if (headerBody == null && footerBody == null) return;

        _paragraphConverter.PageNumberProvider = () => (pageNum, totalPages);
        _paragraphConverter.SectionInfoProvider = () => (sectionIndex, pageNumInSection, totalPagesInSection);

        var pageHeight = pageSize.Height;
        var pageWidth = pageSize.Width;
        var marginL = settings.MarginLeft;
        var marginR = settings.MarginRight;
        var marginT = settings.MarginTop;
        var marginB = settings.MarginBottom;
        var headerDist = settings.HeaderDistance;
        var footerDist = settings.FooterDistance;

        if (headerBody != null)
        {
            var headerHeight = MeasureHeight(headerBody, pageWidth - marginL - marginR);
            // Header position: usually from top edge downwards by HeaderDistance
            // But we render upwards from bottom? No, PDF coords.
            // Header top is pageHeight - HeaderDistance.
            // Or Header baseline?
            // Usually HeaderDistance is from edge to top of header.
            // So Y = pageHeight - HeaderDistance - headerHeight (if top aligned)
            // Or Y range: pageHeight - HeaderDistance (top) to pageHeight - MarginTop (bottom)?
            // Word: Header is inside the top margin area.
            // Header starts at HeaderDistance from top.
            
            float headerY = pageHeight - headerDist;
            // Ensure header fits within margin?
            // If header is tall, it pushes body down? Not in this simple renderer.
            // We just render it.
            
            RenderBody(cb, headerBody,
                marginL, headerY - headerHeight,
                pageWidth - marginR, headerY);
        }

        if (footerBody != null)
        {
            var footerHeight = MeasureHeight(footerBody, pageWidth - marginL - marginR);
            // Footer position: from bottom edge upwards by FooterDistance
            // Footer starts at FooterDistance from bottom.
            float footerY = footerDist;
            
            RenderBody(cb, footerBody,
                marginL, footerY,
                pageWidth - marginR, footerY + footerHeight);
        }
    }

    private float MeasureHeight(OpenXmlElement body, float width)
    {
        // ????:??????
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
    /// ???????(??????????)
    /// </summary>
    public TableConverter? TableConverter { get; set; }
}
