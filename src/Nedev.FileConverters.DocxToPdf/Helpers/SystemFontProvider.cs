using SkiaSharp;
using System.Collections.Concurrent;

namespace Nedev.FileConverters.DocxToPdf.Helpers;

/// <summary>
/// 跨平台系统字体提供者，基于 SkiaSharp 扫描字体
/// </summary>
public static class SystemFontProvider
{
    private static readonly ConcurrentDictionary<string, string> _fontFamilyToPathMap = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<string, string> _fontFileCache = new(StringComparer.OrdinalIgnoreCase);
    private static bool _initialized = false;
    private static readonly object _lock = new();

    /// <summary>
    /// 初始化系统字体映射
    /// </summary>
    public static void Initialize()
    {
        if (_initialized) return;
        lock (_lock)
        {
            if (_initialized) return;
            
            try 
            {
                var fontManager = SKFontManager.Default;
                var families = fontManager.FontFamilies;

                foreach (var family in families)
                {
                    // 获取该家族的第一个字体样式来定位文件路径
                    // 注意：SkiaSharp 可能不直接暴露文件路径，但在很多平台上
                    // 我们可以通过 MatchCharacter 或 MatchStyle 得到 SKTypeface
                    // 但 SKTypeface 不一定暴露 Stream 或 Path。
                    // 
                    // 这里我们采用一种混合策略：
                    // 1. 优先使用 SkiaSharp 确认家族存在
                    // 2. 结合常见系统路径扫描来建立 Family -> Path 映射
                    
                    // 简单起见，我们先扫描常见目录建立缓存
                }
                
                ScanSystemFonts();
                _initialized = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SystemFontProvider] Error initializing fonts: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 尝试获取字体文件路径
    /// </summary>
    public static string? GetFontPath(string fontFamily)
    {
        if (!_initialized) Initialize();

        // 1. 直接匹配
        if (_fontFamilyToPathMap.TryGetValue(fontFamily, out var path))
            return path;
            
        // 2. 移除空格匹配
        if (_fontFamilyToPathMap.TryGetValue(fontFamily.Replace(" ", ""), out path))
            return path;

        // 3. 常见别名映射
        var alias = GetFontAlias(fontFamily);
        if (alias != null && _fontFamilyToPathMap.TryGetValue(alias, out path))
            return path;

        return null;
    }

    private static string? GetFontAlias(string family)
    {
        return family.ToLowerInvariant() switch
        {
            "宋体" => "SimSun",
            "黑体" => "SimHei",
            "微软雅黑" => "Microsoft YaHei",
            "楷体" => "KaiTi",
            "仿宋" => "FangSong",
            "arial" => "Arial",
            "times new roman" => "Times New Roman",
            _ => null
        };
    }

    private static void ScanSystemFonts()
    {
        var fontDirs = new List<string>();

        if (OperatingSystem.IsWindows())
        {
            fontDirs.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"));
            fontDirs.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Windows\\Fonts"));
        }
        else if (OperatingSystem.IsLinux())
        {
            fontDirs.Add("/usr/share/fonts");
            fontDirs.Add("/usr/local/share/fonts");
            fontDirs.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts"));
        }
        else if (OperatingSystem.IsMacOS())
        {
            fontDirs.Add("/System/Library/Fonts");
            fontDirs.Add("/Library/Fonts");
            fontDirs.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library/Fonts"));
        }

        foreach (var dir in fontDirs)
        {
            if (Directory.Exists(dir))
            {
                ScanDirectory(dir);
            }
        }
    }

    private static void ScanDirectory(string dir)
    {
        try
        {
            foreach (var file in Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories))
            {
                var ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext != ".ttf" && ext != ".ttc" && ext != ".otf") continue;

                // 记录文件名映射 (如 simsun.ttc -> simsun)
                var fileName = Path.GetFileNameWithoutExtension(file);
                _fontFileCache.TryAdd(fileName, file);
                
                // 尝试解析字体名称（这里为了性能暂时只用文件名做简单映射）
                // 理想情况是解析 TTF 的 Name 表，但这比较耗时
                // 我们先建立基础映射：
                // simsun -> SimSun
                // msyh -> Microsoft YaHei (需要硬编码一些常见字体文件名)
                
                MapCommonFontFiles(fileName, file);
            }
        }
        catch { /* 忽略访问权限问题 */ }
    }

    private static void MapCommonFontFiles(string fileName, string filePath)
    {
        var name = fileName.ToLowerInvariant();
        
        if (name == "simsun" || name == "simsunb") _fontFamilyToPathMap["SimSun"] = filePath;
        else if (name == "nsimsun") _fontFamilyToPathMap["NSimSun"] = filePath;
        else if (name == "simhei") _fontFamilyToPathMap["SimHei"] = filePath;
        else if (name == "simkai") _fontFamilyToPathMap["KaiTi"] = filePath;
        else if (name == "simfang") _fontFamilyToPathMap["FangSong"] = filePath;
        else if (name == "msyh" || name == "msyhbd") _fontFamilyToPathMap["Microsoft YaHei"] = filePath;
        else if (name == "arial" || name == "arialbd") _fontFamilyToPathMap["Arial"] = filePath;
        else if (name == "times" || name == "timesbd") _fontFamilyToPathMap["Times New Roman"] = filePath;
        else if (name == "calibri" || name == "calibrib") _fontFamilyToPathMap["Calibri"] = filePath;
        else if (name == "consola" || name == "consolab") _fontFamilyToPathMap["Consolas"] = filePath;
        
        // 通用兜底：文件名作为家族名
        _fontFamilyToPathMap.TryAdd(fileName, filePath);
    }
}
