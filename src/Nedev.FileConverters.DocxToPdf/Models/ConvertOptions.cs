using Nedev.FileConverters.DocxToPdf.PdfEngine;

namespace Nedev.FileConverters.DocxToPdf.Models;

/// <summary>
/// DOCX 转 PDF 的配置选项
/// </summary>
public class ConvertOptions
{
    /// <summary>
    /// 页面大小，默认 A4
    /// </summary>
    public Rectangle PageSize { get; set; } = Rectangle.A4;

    /// <summary>
    /// 左边距（pt），默认 72pt = 1 inch
    /// </summary>
    public float MarginLeft { get; set; } = 72f;

    /// <summary>
    /// 右边距（pt），默认 72pt
    /// </summary>
    public float MarginRight { get; set; } = 72f;

    /// <summary>
    /// 上边距（pt），默认 72pt
    /// </summary>
    public float MarginTop { get; set; } = 72f;

    /// <summary>
    /// 下边距（pt），默认 72pt
    /// </summary>
    public float MarginBottom { get; set; } = 72f;

    /// <summary>
    /// 页眉距边界距离（pt），默认 36pt (0.5 inch)
    /// </summary>
    public float HeaderDistance { get; set; } = 36f;

    /// <summary>
    /// 页脚距边界距离（pt），默认 36pt (0.5 inch)
    /// </summary>
    public float FooterDistance { get; set; } = 36f;

    /// <summary>
    /// 默认字体名称（用于中文推荐使用 "STSong-Light" 或系统已注册的字体）
    /// </summary>
    public string DefaultFontName { get; set; } = "STSong-Light";

    /// <summary>
    /// 默认字号（pt）
    /// </summary>
    public float DefaultFontSize { get; set; } = 12f;

    /// <summary>
    /// 额外字体目录列表（用于注册自定义字体）
    /// </summary>
    public List<string> ExtraFontDirectories { get; set; } = [];

    /// <summary>
    /// 是否嵌入图片，默认 true
    /// </summary>
    public bool EmbedImages { get; set; } = true;

    /// <summary>
    /// 是否渲染页眉页脚，默认 true
    /// </summary>
    public bool RenderHeadersFooters { get; set; } = true;

    /// <summary>
    /// 是否在文末输出脚注/尾注内容，默认 true
    /// </summary>
    public bool RenderFootnoteEndContent { get; set; } = true;

    /// <summary>
    /// 是否栅格化图表（Chart），默认 true
    /// </summary>
    public bool RasterizeCharts { get; set; } = true;

    /// <summary>
    /// 是否栅格化 SmartArt，默认 true
    /// </summary>
    public bool RasterizeSmartArt { get; set; } = true;

    /// <summary>
    /// 是否栅格化 DrawingML 形状/文本框，默认 true
    /// </summary>
    public bool RasterizeShapes { get; set; } = true;

    /// <summary>
    /// 是否在文末添加批注汇总页，默认 false
    /// </summary>
    public bool AddCommentsSummaryPage { get; set; } = false;

    /// <summary>
    /// 是否启用性能优化模式（减少内存占用，适合大文档），默认 false
    /// </summary>
    public bool EnablePerformanceMode { get; set; } = false;

    /// <summary>
    /// 图片压缩质量（1-100），仅在性能模式下生效，默认 75
    /// </summary>
    public int ImageCompressionQuality { get; set; } = 75;

    /// <summary>
    /// 是否跳过未使用的字体注册，默认 false
    /// </summary>
    public bool SkipUnusedFonts { get; set; } = false;

    /// <summary>
    /// 水印配置
    /// </summary>
    public WatermarkOptions? Watermark { get; set; }

    /// <summary>
    /// 是否生成目录页，默认 false
    /// </summary>
    public bool GenerateTableOfContents { get; set; } = false;

    /// <summary>
    /// 是否在文末添加修订记录页，默认 false
    /// </summary>
    public bool AddRevisionsSummaryPage { get; set; } = false;

    /// <summary>
    /// 脚注编号格式（"arabic", "roman", "alpha", "chinese"）
    /// </summary>
    public string FootnoteNumberFormat { get; set; } = "arabic";

    /// <summary>
    /// PDF加密/密码保护设置
    /// </summary>
    public PdfEncryptionOptions? Encryption { get; set; }

    /// <summary>
    /// PDF元数据 - 标题
    /// </summary>
    public string? PdfTitle { get; set; }

    /// <summary>
    /// PDF元数据 - 作者
    /// </summary>
    public string? PdfAuthor { get; set; }

    /// <summary>
    /// PDF元数据 - 主题
    /// </summary>
    public string? PdfSubject { get; set; }

    /// <summary>
    /// PDF元数据 - 关键词（逗号分隔）
    /// </summary>
    public string? PdfKeywords { get; set; }

    /// <summary>
    /// PDF元数据 - 创建者（默认"Nedev.FileConverters.DocxToPdf"）
    /// </summary>
    public string PdfCreator { get; set; } = "Nedev.FileConverters.DocxToPdf";

    /// <summary>
    /// 默认选项
    /// </summary>
    public static ConvertOptions Default => new();

    /// <summary>
    /// 当前节的行号设置
    /// </summary>
    public LineNumberSettings? LineNumberSettings { get; set; }
}

/// <summary>
/// 行号设置
/// </summary>
public class LineNumberSettings
{
    /// <summary>起始编号</summary>
    public int Start { get; set; } = 1;
    /// <summary>行号增量（每隔几行显示一次），默认 1</summary>
    public int CountBy { get; set; } = 1;
    /// <summary>距正文距离（pt），默认 0（自动）</summary>
    public float Distance { get; set; } = 0; // twips / 20 -> pt
    /// <summary>重置模式：continuous, newPage, newSection</summary>
    public LineNumberRestartMode RestartMode { get; set; } = LineNumberRestartMode.Continuous;
}

/// <summary>
/// 文本流向
/// </summary>
public enum TextDirection
{
    Horizontal, // lrTb
    Vertical    // tbRl
}

public enum LineNumberRestartMode
{
    Continuous, // 连续
    NewPage,    // 每页重置
    NewSection  // 每节重置
}

/// <summary>
/// PDF加密选项
/// </summary>
public class PdfEncryptionOptions
{
    /// <summary>
    /// 用户密码（打开PDF需要）
    /// </summary>
    public string? UserPassword { get; set; }

    /// <summary>
    /// 所有者密码（用于完全权限）
    /// </summary>
    public string? OwnerPassword { get; set; }

    /// <summary>
    /// 是否允许打印，默认 true
    /// </summary>
    public bool AllowPrint { get; set; } = true;

    /// <summary>
    /// 是否允许修改内容，默认 true
    /// </summary>
    public bool AllowModifyContent { get; set; } = true;

    /// <summary>
    /// 是否允许复制内容，默认 true
    /// </summary>
    public bool AllowCopyContent { get; set; } = true;

    /// <summary>
    /// 是否允许填写表单，默认 true
    /// </summary>
    public bool AllowFillForms { get; set; } = true;
}

/// <summary>
/// 水印配置选项
/// </summary>
public class WatermarkOptions
{
    /// <summary>
    /// 水印文本（如果设置，将使用文本水印而非图片）
    /// </summary>
    public string? Text { get; set; }

    /// <summary>
    /// 水印图片路径（如果设置，将使用图片水印）
    /// </summary>
    public string? ImagePath { get; set; }

    /// <summary>
    /// 水印字体大小（仅文本水印），默认 48pt
    /// </summary>
    public float FontSize { get; set; } = 48f;

    /// <summary>
    /// 水印颜色（RGB），默认浅灰色
    /// </summary>
    public BaseColor? Color { get; set; } = BaseColor.LightGray;

    /// <summary>
    /// 水印透明度（0-1），默认 0.5
    /// </summary>
    public float Opacity { get; set; } = 0.5f;

    /// <summary>
    /// 旋转角度（度），默认 -45 度
    /// </summary>
    public float Rotation { get; set; } = -45f;

    /// <summary>
    /// 水印位置，默认居中
    /// </summary>
    public WatermarkPosition Position { get; set; } = WatermarkPosition.Center;

    /// <summary>
    /// 横向间距（pt），用于平铺模式，默认 200pt
    /// </summary>
    public float HorizontalSpacing { get; set; } = 200f;

    /// <summary>
    /// 纵向间距（pt），用于平铺模式，默认 200pt
    /// </summary>
    public float VerticalSpacing { get; set; } = 200f;

    /// <summary>
    /// 是否平铺水印，默认 false（仅单个水印）
    /// </summary>
    public bool Tiled { get; set; } = false;
}

/// <summary>
/// 水印位置枚举
/// </summary>
public enum WatermarkPosition
{
    Center,
    TopLeft,
    TopRight,
    BottomLeft,
    BottomRight
}
