using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextParagraph = Nedev.FileConverters.DocxToPdf.PdfEngine.Paragraph;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace Nedev.FileConverters.DocxToPdf.Converters;

/// <summary>
/// DOCX 表格转 PDF 表格
/// </summary>
public class TableConverter
{
    private readonly FontHelper _fontHelper;
    private readonly ParagraphConverter _paragraphConverter;
    private readonly ImageConverter? _imageConverter;
    private readonly ListConverter? _listConverter;
    private readonly Numbering? _numbering;
    private readonly Styles? _styles;
    private readonly StyleInheritanceResolver _styleResolver;
    private readonly OpenXmlElement? _colorScheme;

    public TableConverter(
        FontHelper fontHelper,
        ParagraphConverter paragraphConverter,
        ImageConverter? imageConverter = null,
        ListConverter? listConverter = null,
        Numbering? numbering = null,
        Styles? styles = null,
        OpenXmlElement? colorScheme = null)
    {
        _fontHelper = fontHelper;
        _paragraphConverter = paragraphConverter;
        _imageConverter = imageConverter;
        _listConverter = listConverter;
        _numbering = numbering;
        _styles = styles;
        _styleResolver = new StyleInheritanceResolver(styles);
        _colorScheme = colorScheme;
    }

    /// <summary>
    /// ? DOCX Table ?? iTextSharp PdfPTable
    /// </summary>
    public PdfPTable? Convert(WTable docxTable, float pageWidth)
    {
        var rows = docxTable.Elements<TableRow>().ToList();
        if (rows.Count == 0) return null;
        
        // ????
        return BuildTable(docxTable, rows, pageWidth);
    }
    
    /// <summary>
    /// ????????
    /// </summary>
    private PdfPTable? BuildTable(WTable docxTable, List<TableRow> rows, float pageWidth)
    {

        // ????(?? TableGrid,??? GridSpan ????)
        var columnCount = docxTable.TableGrid?.Elements<GridColumn>().Count() ?? 0;
        if (columnCount <= 0)
        {
            columnCount = rows
                .Select(r => r.Elements<TableCell>().Sum(c => (c.TableCellProperties?.GridSpan?.Val?.Value is int s && s > 0) ? s : 1))
                .DefaultIfEmpty(0)
                .Max();
        }
        if (columnCount == 0) return null;

        // ??????
        var tableProps = docxTable.TableProperties;
        var tableStyleId = tableProps?.TableStyle?.Val?.Value;
        var tableDefaultCellMar = tableProps?.GetFirstChild<TableCellMarginDefault>()
                                 ?? GetStyleTableCellMarginDefault(tableStyleId);
        var tableBorders = MergeTableBorders(GetStyleTableBorders(tableStyleId), tableProps?.TableBorders);

        // ????????
        var colWidths = GetColumnWidths(docxTable, columnCount, pageWidth);

        var pdfTable = new PdfPTable(columnCount)
        {
            SpacingBefore = 6f,
            SpacingAfter = 6f
        };

        var tableWidth = tableProps?.TableWidth;
        if (tableWidth?.Type?.Value == TableWidthUnitValues.Pct && tableWidth.Width?.Value != null)
        {
            if (float.TryParse(tableWidth.Width.Value, out var pct))
                pdfTable.WidthPercentage = pct / 50f;
            else
                pdfTable.WidthPercentage = 100;
        }
        else if (tableWidth?.Type?.Value == TableWidthUnitValues.Dxa && tableWidth.Width?.Value != null)
        {
            var widthPt = StyleHelper.DxaToPoints(tableWidth.Width.Value);
            if (widthPt > 0 && widthPt <= pageWidth)
            {
                pdfTable.TotalWidth = widthPt;
                pdfTable.LockedWidth = true;
            }
            else pdfTable.WidthPercentage = 100;
        }
        else
        {
            pdfTable.WidthPercentage = 100;
        }

        // 表格水平对齐 (jc)
        var jc = tableProps?.TableJustification?.Val?.Value;
        if (jc == TableRowAlignmentValues.Center)
            pdfTable.HorizontalAlignment = Element.ALIGN_CENTER;
        else if (jc == TableRowAlignmentValues.Right)
            pdfTable.HorizontalAlignment = Element.ALIGN_RIGHT;

        // ????
        if (colWidths != null && colWidths.Length == columnCount)
        {
            try
            {
                pdfTable.SetWidths(colWidths);
            }
            catch
            {
                // ????????,??????
            }
        }

        // ????
        SetTableBorders(pdfTable, tableBorders);

        // ?? TableLook (????????????????,???????)
        var tblLook = tableProps?.TableLook;
        bool tableHeaderRow = tblLook?.FirstRow?.Value == true;
        bool tableFirstCol = tblLook?.FirstColumn?.Value == true;

        // ??????? null,???? w:val (???????)
        if (tblLook?.Val?.Value != null && uint.TryParse(tblLook.Val.Value, System.Globalization.NumberStyles.HexNumber, null, out var lookVal))
        {
            if ((lookVal & 0x0020) != 0) tableHeaderRow = true; // First Row
            if ((lookVal & 0x0080) != 0) tableFirstCol = true;  // First Column
        }

        var firstRowLookExplicit = tblLook?.FirstRow?.Value;
        var firstColLookExplicit = tblLook?.FirstColumn?.Value;
        var styleFirstRowBold = tableStyleId != null && TableStyleOverrideHasBold(tableStyleId, TableStyleOverrideValues.FirstRow);
        var styleFirstColBold = tableStyleId != null && TableStyleOverrideHasBold(tableStyleId, TableStyleOverrideValues.FirstColumn);

        if (!tableHeaderRow && styleFirstRowBold && firstRowLookExplicit == null)
            tableHeaderRow = true;
        if (!tableFirstCol && styleFirstColBold && firstColLookExplicit == null)
            tableFirstCol = true;

        // ????:?????????
        var headerRowCount = 0;
        for (var i = 0; i < rows.Count; i++)
        {
            var trPr = rows[i].TableRowProperties;
            if (trPr?.GetFirstChild<TableHeader>() != null)
                headerRowCount = i + 1;
            else
                break;
        }
        if (headerRowCount > 0)
            pdfTable.HeaderRows = headerRowCount;

        // ??????(??????)
        pdfTable.KeepTogether = rows.Count <= 3;

        var rowCount = rows.Count;

        var placementsByRow = BuildCellPlacements(rows, columnCount);
        ComputeRowSpans(placementsByRow);

        pdfTable.TableEvent = new OuterBorderTableEvent(tableBorders, _colorScheme);

        var rowspanLeft = new int[columnCount];

        // ???
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var row = rows[rowIndex];
            var trPr = row.TableRowProperties;
            
            // ????
            var trHeight = trPr?.GetFirstChild<TableRowHeight>();
            float? rowHeight = null;
            bool isFixedHeight = false;
            if (trHeight?.Val?.Value is uint heightVal)
            {
                rowHeight = StyleHelper.TwipsToPoints((int)heightVal);
                isFixedHeight = trHeight.HeightType?.Value == HeightRuleValues.Exact;
            }

            pdfTable.RowsList.Add(new PdfPRow());

            if (rowIndex > 0)
            {
                for (var c = 0; c < columnCount; c++)
                {
                    if (rowspanLeft[c] > 0) rowspanLeft[c]--;
                }
            }

            var rowPlacements = placementsByRow[rowIndex]
                .Where(p => !p.IsContinue)
                .GroupBy(p => p.StartCol)
                .ToDictionary(g => g.Key, g => g.First());

            for (var colIndex = 0; colIndex < columnCount;)
            {
                if (rowspanLeft[colIndex] > 0)
                {
                    colIndex++;
                    continue;
                }

                if (!rowPlacements.TryGetValue(colIndex, out var placement))
                {
                    var empty = CreateEmptyCell(tableDefaultCellMar, tableBorders, rowIndex, colIndex, rowCount, columnCount, rowHeight, isFixedHeight);
                    empty.ColIndex = colIndex;
                    pdfTable.RowsList.Last().Cells.Add(empty);
                    colIndex++;
                    continue;
                }

                bool forceBold = (rowIndex == 0 && tableHeaderRow)
                                 || (colIndex == 0 && tableFirstCol);

                // ????:???/?
                var cellShading = GetConditionalShading(tableStyleId, tblLook, rowIndex, colIndex, rowCount, columnCount);

                var cellWidth = EstimateCellWidth(colWidths, pageWidth, columnCount, colIndex, placement.ColSpan);
                var pdfCell = ConvertCell(
                    placement.Cell,
                    tableDefaultCellMar,
                    tableBorders,
                    rowIndex,
                    colIndex,
                    rowCount,
                    columnCount,
                    cellWidth,
                    rowHeight,
                    isFixedHeight,
                    forceBold);

                // ????????
                if (cellShading != null && pdfCell.BackgroundColor == null)
                    pdfCell.BackgroundColor = cellShading;

                if (placement.RowSpan > 1) pdfCell.Rowspan = placement.RowSpan;
                pdfCell.ColIndex = colIndex;

                pdfTable.RowsList.Last().Cells.Add(pdfCell);

                if (placement.RowSpan > 1)
                {
                    for (var c = colIndex; c < Math.Min(columnCount, colIndex + Math.Max(placement.ColSpan, 1)); c++)
                    {
                        rowspanLeft[c] = Math.Max(rowspanLeft[c], placement.RowSpan - 1);
                    }
                }

                colIndex += Math.Max(placement.ColSpan, 1);
            }
        }

        return pdfTable;
    }

    private static float EstimateCellWidth(float[]? colWidths, float pageWidth, int columnCount, int startCol, int colSpan)
    {
        var span = Math.Max(colSpan, 1);
        if (colWidths != null && colWidths.Length == columnCount)
        {
            var sum = 0f;
            for (var i = startCol; i < Math.Min(columnCount, startCol + span); i++)
            {
                sum += colWidths[i];
            }
            if (sum > 0) return sum;
        }

        return pageWidth * (span / (float)Math.Max(columnCount, 1));
    }

    private sealed class CellPlacement
    {
        public required TableCell Cell { get; init; }
        public required int StartCol { get; init; }
        public required int ColSpan { get; init; }
        public required bool IsContinue { get; init; }
        public required bool IsRestart { get; init; }
        public int RowSpan { get; set; } = 1;
    }

    private sealed class OuterBorderTableEvent
    {
        private readonly TableBorders? _borders;
        private readonly OpenXmlElement? _colorScheme;

        public OuterBorderTableEvent(TableBorders? borders, OpenXmlElement? colorScheme)
        {
            _borders = borders;
            _colorScheme = colorScheme;
        }

        public void TableLayout(PdfPTable table, float[][] widths, float[] heights, int headerRows, int rowStart, PdfContentByte[] canvases)
        {
            if (canvases == null || canvases.Length == 0) return;

            var cb = canvases[PdfPTable.LINECANVAS];
            if (cb == null) return;

            if (widths.Length == 0 || heights.Length == 0) return;
            if (widths[0].Length < 2) return;

            var left = widths[0][0];
            var right = widths[0][widths[0].Length - 1];
            var top = heights[0];
            var bottom = heights[heights.Length - 1];

            BorderType? Top()
            {
                BorderType? b = _borders?.TopBorder;
                return b ?? (BorderType?)_borders?.InsideHorizontalBorder;
            }

            BorderType? Bottom()
            {
                BorderType? b = _borders?.BottomBorder;
                return b ?? (BorderType?)_borders?.InsideHorizontalBorder;
            }

            BorderType? Left()
            {
                BorderType? b = _borders?.LeftBorder;
                if (b != null) return b;
                b = _borders?.StartBorder;
                return b ?? (BorderType?)_borders?.InsideVerticalBorder;
            }

            BorderType? Right()
            {
                BorderType? b = _borders?.RightBorder;
                if (b != null) return b;
                b = _borders?.EndBorder;
                return b ?? (BorderType?)_borders?.InsideVerticalBorder;
            }

            var topW = _borders != null ? StyleHelper.GetBorderWidth(Top()) : 0.5f;
            var bottomW = _borders != null ? StyleHelper.GetBorderWidth(Bottom()) : 0.5f;
            var leftW = _borders != null ? StyleHelper.GetBorderWidth(Left()) : 0.5f;
            var rightW = _borders != null ? StyleHelper.GetBorderWidth(Right()) : 0.5f;

            var topC = StyleHelper.ResolveBorderColor(_colorScheme, Top()) ?? new BaseColor(200, 200, 200);
            var bottomC = StyleHelper.ResolveBorderColor(_colorScheme, Bottom()) ?? new BaseColor(200, 200, 200);
            var leftC = StyleHelper.ResolveBorderColor(_colorScheme, Left()) ?? new BaseColor(200, 200, 200);
            var rightC = StyleHelper.ResolveBorderColor(_colorScheme, Right()) ?? new BaseColor(200, 200, 200);

            if (topW > 0)
            {
                cb.SaveState();
                cb.SetLineWidth(topW);
                cb.SetColorStroke(topC);
                cb.MoveTo(left, top);
                cb.LineTo(right, top);
                cb.Stroke();
                cb.RestoreState();
            }

            if (bottomW > 0)
            {
                cb.SaveState();
                cb.SetLineWidth(bottomW);
                cb.SetColorStroke(bottomC);
                cb.MoveTo(left, bottom);
                cb.LineTo(right, bottom);
                cb.Stroke();
                cb.RestoreState();
            }

            if (leftW > 0)
            {
                cb.SaveState();
                cb.SetLineWidth(leftW);
                cb.SetColorStroke(leftC);
                cb.MoveTo(left, bottom);
                cb.LineTo(left, top);
                cb.Stroke();
                cb.RestoreState();
            }

            if (rightW > 0)
            {
                cb.SaveState();
                cb.SetLineWidth(rightW);
                cb.SetColorStroke(rightC);
                cb.MoveTo(right, bottom);
                cb.LineTo(right, top);
                cb.Stroke();
                cb.RestoreState();
            }
        }
    }

    private static List<List<CellPlacement>> BuildCellPlacements(List<TableRow> rows, int columnCount)
    {
        var result = new List<List<CellPlacement>>(rows.Count);

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            var startCol = 0;

            var placements = new List<CellPlacement>();
            foreach (var cell in row.Elements<TableCell>())
            {
                if (startCol >= columnCount) break;

                var cellProps = cell.TableCellProperties;
                var colSpan = (cellProps?.GridSpan?.Val?.Value is int s && s > 0) ? s : 1;

                var vMerge = cellProps?.VerticalMerge;
                var isRestart = vMerge != null && vMerge.Val?.Value == MergedCellValues.Restart;
                var isContinue = vMerge != null && (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Continue);

                placements.Add(new CellPlacement
                {
                    Cell = cell,
                    StartCol = startCol,
                    ColSpan = colSpan,
                    IsContinue = isContinue,
                    IsRestart = isRestart
                });

                startCol = Math.Min(columnCount, startCol + Math.Max(colSpan, 1));
            }

            result.Add(placements);
        }

        return result;
    }

    private static void ComputeRowSpans(List<List<CellPlacement>> placementsByRow)
    {
        var lookups = placementsByRow
            .Select(r => r.GroupBy(p => p.StartCol).ToDictionary(g => g.Key, g => g.First()))
            .ToList();

        for (var rowIndex = 0; rowIndex < placementsByRow.Count; rowIndex++)
        {
            foreach (var p in placementsByRow[rowIndex])
            {
                if (!p.IsRestart) continue;

                var span = 1;
                for (var r = rowIndex + 1; r < placementsByRow.Count; r++)
                {
                    if (!lookups[r].TryGetValue(p.StartCol, out var down)) break;
                    if (!down.IsContinue) break;
                    if (down.ColSpan != p.ColSpan) break;
                    span++;
                }

                p.RowSpan = span;
            }
        }
    }

    private PdfPCell CreateEmptyCell(
        TableCellMarginDefault? tableDefaultCellMar,
        TableBorders? tableBorders,
        int rowIndex,
        int colIndex,
        int rowCount,
        int columnCount,
        float? height = null,
        bool isFixedHeight = false)
    {
        var pdfCell = new PdfPCell
        {
            UseAscender = true,
            UseDescender = true,
            VerticalAlignment = Element.ALIGN_MIDDLE
        };

        if (height.HasValue)
        {
            if (isFixedHeight) pdfCell.FixedHeight = height.Value;
            else pdfCell.MinimumHeight = height.Value;
        }

        if (tableDefaultCellMar != null)
        {
            var top = GetTableCellMarginSidePoints(tableDefaultCellMar, "top");
            var bottom = GetTableCellMarginSidePoints(tableDefaultCellMar, "bottom");
            var left = GetTableCellMarginSidePoints(tableDefaultCellMar, "left");
            var right = GetTableCellMarginSidePoints(tableDefaultCellMar, "right");

            pdfCell.PaddingTop = top ?? 2f;
            pdfCell.PaddingBottom = bottom ?? 2f;
            pdfCell.PaddingLeft = left ?? 5.4f;
            pdfCell.PaddingRight = right ?? 5.4f;
        }
        else
        {
            pdfCell.PaddingTop = 2f;
            pdfCell.PaddingBottom = 2f;
            pdfCell.PaddingLeft = 5.4f;
            pdfCell.PaddingRight = 5.4f;
        }

        SetCellBorders(pdfCell, null, tableBorders, _colorScheme, rowIndex, colIndex, rowCount, columnCount, 1);

        pdfCell.Phrase = new Phrase(" ");
        return pdfCell;
    }

    private TableCellMarginDefault? GetStyleTableCellMarginDefault(string? styleId)
    {
        return GetFromStyleChain(styleId, s => s.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableCellMarginDefault>());
    }

    private TableBorders? GetStyleTableBorders(string? styleId)
    {
        return GetFromStyleChain(styleId, s => s.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableBorders>());
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

    private static TableBorders? MergeTableBorders(TableBorders? baseBorders, TableBorders? overrideBorders)
    {
        if (baseBorders == null) return overrideBorders;
        if (overrideBorders == null) return baseBorders;

        static bool IsExplicit(BorderType? b)
        {
            if (b == null) return false;
            if (b.Val != null) return true;
            if (b.Size != null) return true;
            if (b.Color != null) return true;
            if (b.ThemeColor != null) return true;
            if (b.ThemeTint != null) return true;
            if (b.ThemeShade != null) return true;
            return false;
        }

        static TBorder? Pick<TBorder>(TBorder? baseB, TBorder? overrideB) where TBorder : BorderType
        {
            if (IsExplicit(overrideB)) return overrideB;
            return baseB;
        }

        var merged = new TableBorders
        {
            TopBorder = (TopBorder?)(Pick(baseBorders.TopBorder, overrideBorders.TopBorder)?.CloneNode(true)),
            BottomBorder = (BottomBorder?)(Pick(baseBorders.BottomBorder, overrideBorders.BottomBorder)?.CloneNode(true)),
            LeftBorder = (LeftBorder?)(Pick(baseBorders.LeftBorder, overrideBorders.LeftBorder)?.CloneNode(true)),
            RightBorder = (RightBorder?)(Pick(baseBorders.RightBorder, overrideBorders.RightBorder)?.CloneNode(true)),
            InsideHorizontalBorder = (InsideHorizontalBorder?)(Pick(baseBorders.InsideHorizontalBorder, overrideBorders.InsideHorizontalBorder)?.CloneNode(true)),
            InsideVerticalBorder = (InsideVerticalBorder?)(Pick(baseBorders.InsideVerticalBorder, overrideBorders.InsideVerticalBorder)?.CloneNode(true)),
            StartBorder = (StartBorder?)(Pick(baseBorders.StartBorder, overrideBorders.StartBorder)?.CloneNode(true)),
            EndBorder = (EndBorder?)(Pick(baseBorders.EndBorder, overrideBorders.EndBorder)?.CloneNode(true))
        };

        if (merged.LeftBorder == null && merged.StartBorder != null)
            merged.LeftBorder = (LeftBorder)merged.StartBorder.CloneNode(true);
        if (merged.RightBorder == null && merged.EndBorder != null)
            merged.RightBorder = (RightBorder)merged.EndBorder.CloneNode(true);

        return merged;
    }

    /// <summary>
    /// ????
    /// </summary>
    private float[]? GetColumnWidths(WTable docxTable, int columnCount, float pageWidth)
    {
        var grid = docxTable.TableGrid;
        if (grid == null) return null;

        var gridColumns = grid.Elements<GridColumn>().ToList();
        if (gridColumns.Count != columnCount) return null;

        var widths = new float[columnCount];
        float totalWidth = 0;

        for (var i = 0; i < columnCount; i++)
        {
            var w = StyleHelper.DxaToPoints(gridColumns[i].Width?.Value);
            widths[i] = w > 0 ? w : 1f;
            totalWidth += widths[i];
        }

        // ???
        if (totalWidth > 0)
        {
            for (var i = 0; i < columnCount; i++)
            {
                widths[i] = widths[i] / totalWidth * pageWidth;
            }
        }

        return widths;
    }

    /// <summary>
    /// ?????
    /// </summary>
    private PdfPCell ConvertCell(
        TableCell docxCell,
        TableCellMarginDefault? tableDefaultCellMar,
        TableBorders? tableBorders,
        int rowIndex,
        int colIndex,
        int rowCount,
        int columnCount,
        float cellWidth,
        float? height = null,
        bool isFixedHeight = false,
        bool forceBold = false)
    {
        var pdfCell = new PdfPCell
        {
            UseAscender = true,
            UseDescender = true,
            VerticalAlignment = Element.ALIGN_MIDDLE // ????
        };

        if (height.HasValue)
        {
            if (isFixedHeight) pdfCell.FixedHeight = height.Value;
            else pdfCell.MinimumHeight = height.Value;
        }

        var cellProps = docxCell.TableCellProperties;

        // ?????????
        var cellMar = cellProps?.TableCellMargin;
        if (cellMar != null)
        {
            if (cellMar.TopMargin?.Width?.Value is string topM) pdfCell.PaddingTop = StyleHelper.DxaToPoints(topM);
            if (cellMar.BottomMargin?.Width?.Value is string bottomM) pdfCell.PaddingBottom = StyleHelper.DxaToPoints(bottomM);
            if (cellMar.LeftMargin?.Width?.Value is string leftM) pdfCell.PaddingLeft = StyleHelper.DxaToPoints(leftM);
            if (cellMar.RightMargin?.Width?.Value is string rightM) pdfCell.PaddingRight = StyleHelper.DxaToPoints(rightM);
        }
        else if (tableDefaultCellMar != null)
        {
            var top = GetTableCellMarginSidePoints(tableDefaultCellMar, "top");
            var bottom = GetTableCellMarginSidePoints(tableDefaultCellMar, "bottom");
            var left = GetTableCellMarginSidePoints(tableDefaultCellMar, "left");
            var right = GetTableCellMarginSidePoints(tableDefaultCellMar, "right");

            if (top.HasValue) pdfCell.PaddingTop = top.Value;
            if (bottom.HasValue) pdfCell.PaddingBottom = bottom.Value;
            if (left.HasValue) pdfCell.PaddingLeft = left.Value;
            if (right.HasValue) pdfCell.PaddingRight = right.Value;
        }
        else
        {
            pdfCell.PaddingTop = 2f;
            pdfCell.PaddingBottom = 2f;
            pdfCell.PaddingLeft = 5.4f;
            pdfCell.PaddingRight = 5.4f;
        }
        
        // ??????????????,???????????????
        if (forceBold)
        {
             pdfCell.PaddingTop = Math.Max(pdfCell.PaddingTop, 3f);
             pdfCell.PaddingBottom = Math.Max(pdfCell.PaddingBottom, 3f);
        }

        // ????
        var colSpan = 1;
        if (cellProps?.GridSpan?.Val?.Value is int gridSpan && gridSpan > 1)
        {
            pdfCell.Colspan = gridSpan;
            colSpan = gridSpan;
        }

        // ????
        if (cellProps?.VerticalMerge != null)
        {
            var mergeVal = cellProps.VerticalMerge.Val;
            if (mergeVal == null || mergeVal.Value == MergedCellValues.Continue)
            {
                pdfCell.Phrase = new Phrase(" ");
                pdfCell.Border = Rectangle.NO_BORDER;
                pdfCell.FixedHeight = 0.1f;
                return pdfCell;
            }
        }

        // ??????
        var bgColor = StyleHelper.ResolveShadingFill(_colorScheme, cellProps?.Shading);
        if (bgColor != null) pdfCell.BackgroundColor = bgColor;

        if (cellProps?.TableCellVerticalAlignment?.Val?.Value is TableVerticalAlignmentValues vAlign)
        {
            if (vAlign.Equals(TableVerticalAlignmentValues.Center))
                pdfCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            else if (vAlign.Equals(TableVerticalAlignmentValues.Bottom))
                pdfCell.VerticalAlignment = Element.ALIGN_BOTTOM;
            else
                pdfCell.VerticalAlignment = Element.ALIGN_TOP;
        }

        // ?????
        SetCellBorders(pdfCell, cellProps, tableBorders, _colorScheme, rowIndex, colIndex, rowCount, columnCount, colSpan);

        // ????? � ??(???????)
        void AddParagraph(WParagraph para)
        {
            var pdfElements = _paragraphConverter.Convert(para, forceBold);
            foreach (var element in pdfElements)
            {
                if (element is Chunk c && c.Content == "PAGE_BREAK") continue;
                pdfCell.AddElement(element);
            }

            if (_imageConverter != null)
            {
                var images = _imageConverter.ConvertImagesInParagraph(para, cellWidth, null);
                foreach (var img in images)
                {
                    pdfCell.AddElement(img);
                }
            }
        }

        var children = docxCell.ChildElements.ToList();
        var i = 0;
        while (i < children.Count)
        {
            var child = children[i];
            switch (child)
            {
                case WParagraph para:
                    if (_listConverter != null && _numbering != null && ListConverter.IsListItem(para))
                    {
                        var numberingId = ListConverter.GetNumberingId(para);
                        if (numberingId.HasValue)
                        {
                            var listParagraphs = new System.Collections.Generic.List<WParagraph> { para };
                            var j = i + 1;
                            while (j < children.Count && children[j] is WParagraph nextPara
                                   && ListConverter.IsListItem(nextPara)
                                   && ListConverter.GetNumberingId(nextPara) == numberingId)
                            {
                                listParagraphs.Add(nextPara);
                                j++;
                            }

                            var pdfList = _listConverter.ConvertListItems(listParagraphs, _numbering, numberingId.Value);
                            pdfCell.AddElement(pdfList);
                            i = j;
                            continue;
                        }
                    }

                    AddParagraph(para);
                    i++;
                    break;
                case SdtBlock sdt:
                    var content = sdt.SdtContentBlock;
                    if (content == null) { i++; break; }
                    foreach (var inner in content.ChildElements)
                    {
                        if (inner is WParagraph innerPara) AddParagraph(innerPara);
                        else if (inner is WTable innerTable)
                        {
                            var nestedAvailableWidth = Math.Max(10f, cellWidth - pdfCell.PaddingLeft - pdfCell.PaddingRight);
                            var nested = Convert(innerTable, nestedAvailableWidth);
                            if (nested != null)
                            {
                                nested.WidthPercentage = 100;
                                nested.SpacingBefore = 2f;
                                nested.SpacingAfter = 2f;
                                pdfCell.AddElement(nested);
                            }
                        }
                    }
                    i++;
                    break;
                case WTable table:
                    // ??????:????,?????
                    var availableWidth = Math.Max(10f, cellWidth - pdfCell.PaddingLeft - pdfCell.PaddingRight);
                    var innerPdfTable = Convert(table, availableWidth);
                    if (innerPdfTable != null)
                    {
                        // ???????????? 100%,????????
                        innerPdfTable.WidthPercentage = 100;
                        // ????????????
                        innerPdfTable.SpacingBefore = 2f;
                        innerPdfTable.SpacingAfter = 2f;
                        pdfCell.AddElement(innerPdfTable);
                    }
                    i++;
                    break;
                default:
                    i++;
                    break;
            }
        }

        // ???????
        if (!docxCell.HasChildren || !docxCell.Descendants<WParagraph>().Any())
        {
            pdfCell.Phrase = new Phrase(" ");
        }

        return pdfCell;
    }

    private static float? GetTableCellMarginSidePoints(OpenXmlCompositeElement marElem, string sideLocalName)
    {
        var side = marElem.ChildElements.FirstOrDefault(e =>
            e.LocalName.Equals(sideLocalName, StringComparison.OrdinalIgnoreCase));
        if (side == null) return null;

        var w = side.GetAttributes().FirstOrDefault(a =>
            a.LocalName.Equals("w", StringComparison.OrdinalIgnoreCase) ||
            a.LocalName.Equals("width", StringComparison.OrdinalIgnoreCase));

        return w.Value != null ? StyleHelper.DxaToPoints(w.Value) : null;
    }

    private bool TableStyleOverrideHasBold(string styleId, TableStyleOverrideValues overrideType)
    {
        var style = _styles?
            .Elements<Style>()
            .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));

        if (style == null) return false;

        var overrideProps = style
            .Elements<TableStyleProperties>()
            .FirstOrDefault(p => p.Type?.Value != null && p.Type.Value.Equals(overrideType));

        var rPr = overrideProps?.GetFirstChild<RunProperties>();
        return rPr?.Bold != null || rPr?.BoldComplexScript != null;
    }

    private BaseColor? GetConditionalShading(string? styleId, TableLook? tblLook, int rowIndex, int colIndex, int rowCount, int columnCount)
    {
        if (string.IsNullOrEmpty(styleId)) return null;

        var style = _styles?
            .Elements<Style>()
            .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
        if (style == null) return null;

        // ??????????
        TableStyleOverrideValues? overrideType = null;
        if (rowIndex == 0 && tblLook?.FirstRow?.Value == true)
            overrideType = TableStyleOverrideValues.FirstRow;
        else if (rowIndex == rowCount - 1 && tblLook?.LastRow?.Value == true)
            overrideType = TableStyleOverrideValues.LastRow;
        else if (colIndex == 0 && tblLook?.FirstColumn?.Value == true)
            overrideType = TableStyleOverrideValues.FirstColumn;
        else if (colIndex == columnCount - 1 && tblLook?.LastColumn?.Value == true)
            overrideType = TableStyleOverrideValues.LastColumn;
        else if (tblLook?.NoHorizontalBand?.Value != true && rowIndex % 2 == 1)
            overrideType = TableStyleOverrideValues.Band1Horizontal;
        else if (tblLook?.NoVerticalBand?.Value != true && colIndex % 2 == 1)
            overrideType = TableStyleOverrideValues.Band1Vertical;

        if (overrideType == null) return null;

        var overrideProps = style
            .Elements<TableStyleProperties>()
            .FirstOrDefault(p => p.Type?.Value != null && p.Type.Value.Equals(overrideType.Value));

        var tcPr = overrideProps?.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>();
        var shading = tcPr?.GetFirstChild<Shading>();
        return shading != null ? StyleHelper.ResolveShadingFill(_colorScheme, shading) : null;
    }

    /// <summary>
    /// ????????
    /// </summary>
    private static void SetTableBorders(PdfPTable pdfTable, TableBorders? borders)
    {
        // ????
        pdfTable.DefaultCell.BorderWidth = 0.5f;
        pdfTable.DefaultCell.BorderColor = new BaseColor(200, 200, 200);

        if (borders == null) return;

        // ??:PdfPTable ??????,? Border ?????????????
        // ??????? DefaultCell ?????
        if (borders.TopBorder?.Val?.Value is BorderValues topVal)
            pdfTable.DefaultCell.BorderWidthTop = StyleHelper.GetBorderWidth(topVal);
        if (borders.BottomBorder?.Val?.Value is BorderValues botVal)
            pdfTable.DefaultCell.BorderWidthBottom = StyleHelper.GetBorderWidth(botVal);
        if (borders.LeftBorder?.Val?.Value is BorderValues leftVal)
            pdfTable.DefaultCell.BorderWidthLeft = StyleHelper.GetBorderWidth(leftVal);
        if (borders.RightBorder?.Val?.Value is BorderValues rightVal)
            pdfTable.DefaultCell.BorderWidthRight = StyleHelper.GetBorderWidth(rightVal);
    }

    /// <summary>
    /// ???????
    /// </summary>
    private static void SetCellBorders(PdfPCell pdfCell, TableCellProperties? cellProps, TableBorders? tableBorders, OpenXmlElement? colorScheme, int rowIndex, int colIndex, int rowCount, int columnCount, int colSpan)
    {
        var borders = cellProps?.TableCellBorders;

        // ????????????????,????????
        if (borders == null && tableBorders == null)
        {
            pdfCell.BorderWidth = 0.5f;
            pdfCell.BorderColor = new BaseColor(200, 200, 200);
            return;
        }

        float GetW(BorderType? b) => StyleHelper.GetBorderWidth(b);
        BaseColor? GetC(BorderType? b) => StyleHelper.ResolveBorderColor(colorScheme, b);

        static bool IsExplicit(BorderType? b)
        {
            if (b == null) return false;
            if (b.Val != null) return true;
            if (b.Size != null) return true;
            if (b.Color != null) return true;
            if (b.ThemeColor != null) return true;
            if (b.ThemeTint != null) return true;
            if (b.ThemeShade != null) return true;
            return false;
        }

        var defaultColor = StyleHelper.ResolveBorderColor(colorScheme, tableBorders?.InsideHorizontalBorder)
                           ?? StyleHelper.ResolveBorderColor(colorScheme, tableBorders?.InsideVerticalBorder)
                           ?? StyleHelper.ResolveBorderColor(colorScheme, tableBorders?.TopBorder)
                           ?? StyleHelper.ResolveBorderColor(colorScheme, tableBorders?.LeftBorder)
                           ?? StyleHelper.ResolveBorderColor(colorScheme, tableBorders?.StartBorder)
                           ?? new BaseColor(200, 200, 200);
        var isLastRow = rowIndex >= rowCount - 1;
        var isLastCol = colIndex + Math.Max(colSpan, 1) - 1 >= columnCount - 1;

        BorderType? TableTop()
        {
            if (rowIndex != 0) return tableBorders?.InsideHorizontalBorder;
            return (BorderType?)tableBorders?.TopBorder ?? tableBorders?.InsideHorizontalBorder;
        }

        BorderType? TableBottom()
        {
            if (!isLastRow) return tableBorders?.InsideHorizontalBorder;
            return (BorderType?)tableBorders?.BottomBorder ?? tableBorders?.InsideHorizontalBorder;
        }

        BorderType? EffectiveLeft()
        {
            BorderType? b = tableBorders?.LeftBorder;
            return b ?? (BorderType?)tableBorders?.StartBorder;
        }

        BorderType? EffectiveRight()
        {
            BorderType? b = tableBorders?.RightBorder;
            return b ?? (BorderType?)tableBorders?.EndBorder;
        }
        BorderType? TableLeft()
        {
            if (colIndex != 0) return tableBorders?.InsideVerticalBorder;
            return EffectiveLeft() ?? tableBorders?.InsideVerticalBorder;
        }

        BorderType? TableRight()
        {
            if (!isLastCol) return tableBorders?.InsideVerticalBorder;
            return EffectiveRight() ?? tableBorders?.InsideVerticalBorder;
        }
        BorderType? topBorder = IsExplicit(borders?.TopBorder) ? borders?.TopBorder : TableTop();
        BorderType? bottomBorder = IsExplicit(borders?.BottomBorder) ? borders?.BottomBorder : TableBottom();
        BorderType? leftBorder = IsExplicit(borders?.LeftBorder) ? borders?.LeftBorder : TableLeft();
        BorderType? rightBorder = IsExplicit(borders?.RightBorder) ? borders?.RightBorder : TableRight();

        pdfCell.BorderWidthTop = GetW(topBorder);
        pdfCell.BorderWidthBottom = GetW(bottomBorder);
        pdfCell.BorderWidthLeft = GetW(leftBorder);
        pdfCell.BorderWidthRight = GetW(rightBorder);

        var ct = GetC(topBorder);
        var cb = GetC(bottomBorder);
        var cl = GetC(leftBorder);
        var cr = GetC(rightBorder);

        if (pdfCell.BorderWidthTop > 0) pdfCell.BorderColorTop = ct ?? defaultColor;
        if (pdfCell.BorderWidthBottom > 0) pdfCell.BorderColorBottom = cb ?? defaultColor;
        if (pdfCell.BorderWidthLeft > 0) pdfCell.BorderColorLeft = cl ?? defaultColor;
        if (pdfCell.BorderWidthRight > 0) pdfCell.BorderColorRight = cr ?? defaultColor;

        // Border collapsing is handled by PdfPTable logic. 
        // Explicit BorderValues.None checks follow below.

        // ????:?????????,??????0,??????????(iTextSharp ??)
        // ??????????????????0
        if (pdfCell.BackgroundColor != null)
        {
            // ?????? "None",???? 0
            if (topBorder?.Val?.Value == BorderValues.None) pdfCell.BorderWidthTop = 0;
            if (bottomBorder?.Val?.Value == BorderValues.None) pdfCell.BorderWidthBottom = 0;
            if (leftBorder?.Val?.Value == BorderValues.None) pdfCell.BorderWidthLeft = 0;
            if (rightBorder?.Val?.Value == BorderValues.None) pdfCell.BorderWidthRight = 0;
        }
    }
}
