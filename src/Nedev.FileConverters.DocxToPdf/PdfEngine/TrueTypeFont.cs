using System.Text;

namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// TrueType????? - ?????????????
/// </summary>
public class TrueTypeFont
{
    private byte[] _data;
    private int _pos;
    private bool _isTTC = false;
    private uint _ttcOffset = 0;

    public string FontName { get; private set; } = "SimSun";
    public int UnitsPerEm { get; private set; } = 1000;
    public short Ascent { get; private set; } = 800;
    public short Descent { get; private set; } = -200;
    public short CapHeight { get; private set; } = 700;
    public short StemV { get; private set; } = 80;
    public int Flags { get; private set; } = 32;
    public short ItalicAngle { get; private set; } = 0;
    public short XMin { get; private set; } = -100;
    public short YMin { get; private set; } = -200;
    public short XMax { get; private set; } = 1000;
    public short YMax { get; private set; } = 800;

    // ?????Unicode???
    public Dictionary<int, char> GlyphToUnicode { get; private set; } = new();
    public Dictionary<char, int> UnicodeToGlyph { get; private set; } = new();

    public int[]? GlyphWidths { get; private set; }
    private int _numberOfHMetrics;

    public TrueTypeFont(byte[] data)
    {
        _data = data;
        Parse();
    }

    private void Parse()
    {
        _pos = 0;

        // ?????TTC??
        var tag = ReadTag();
        if (tag == "ttcf")
        {
            _isTTC = true;
            ParseTTCHeader();
        }
        else
        {
            // ??TTF??,????
            _pos = 0;
        }

        // ?????
        var sfntVersion = ReadUInt32();
        var numTables = ReadUInt16();
        var searchRange = ReadUInt16();
        var entrySelector = ReadUInt16();
        var rangeShift = ReadUInt16();

        // ?????
        var tables = new Dictionary<string, (uint offset, uint length)>();
        for (int i = 0; i < numTables; i++)
        {
            var tableTag = ReadTag();
            var checksum = ReadUInt32();
            var offset = ReadUInt32();
            var length = ReadUInt32();
            // Offset is from beginning of file, NOT relative to TTC header
            tables[tableTag] = (offset, length);
        }

        // ??head?
        if (tables.TryGetValue("head", out var headInfo))
        {
            ParseHeadTable(headInfo.offset);
        }

        // ?? hmtx ???
        uint hmtxOffset = 0;
        if (tables.TryGetValue("hmtx", out var hmtxInfo))
        {
            hmtxOffset = hmtxInfo.offset;
        }

        // ??hhea?
        if (tables.TryGetValue("hhea", out var hheaInfo))
        {
            ParseHheaTable(hheaInfo.offset);
        }

        // ??hmtx?
        if (hmtxOffset > 0 && _numberOfHMetrics > 0)
        {
            ParseHmtxTable(hmtxOffset);
        }

        // ??name???????
        if (tables.TryGetValue("name", out var nameInfo))
        {
            ParseNameTable(nameInfo.offset);
        }

        // ??cmap???????
        if (tables.TryGetValue("cmap", out var cmapInfo))
        {
            ParseCmapTable(cmapInfo.offset);
        }

        // ????????,?????
        if (string.IsNullOrEmpty(FontName))
        {
            FontName = "SimSun";
        }
    }

    private void ParseTTCHeader()
    {
        // TTC Header
        var ttcVersion = ReadUInt32(); // Version
        var numFonts = ReadUInt32();   // Number of fonts in collection

        // ???????????
        if (numFonts > 0)
        {
            _ttcOffset = ReadUInt32();
            _pos = (int)_ttcOffset;
        }
    }

    private void ParseHeadTable(uint offset)
    {
        _pos = (int)offset;
        var version = ReadUInt32();
        var fontRevision = ReadUInt32();
        var checkSumAdjustment = ReadUInt32();
        var magicNumber = ReadUInt32();
        var flags = ReadUInt16();
        UnitsPerEm = ReadUInt16();
        var created = ReadUInt64();
        var modified = ReadUInt64();
        XMin = ReadInt16();
        YMin = ReadInt16();
        XMax = ReadInt16();
        YMax = ReadInt16();
        var macStyle = ReadUInt16();
        var lowestRecPPEM = ReadUInt16();
        var fontDirectionHint = ReadInt16();
        var indexToLocFormat = ReadInt16();
        var glyphDataFormat = ReadInt16();
    }

    private void ParseHheaTable(uint offset)
    {
        _pos = (int)offset;
        var version = ReadUInt32();
        Ascent = ReadInt16();
        Descent = ReadInt16();
        var lineGap = ReadInt16();
        var advanceWidthMax = ReadUInt16();
        var minLeftSideBearing = ReadInt16();
        var minRightSideBearing = ReadInt16();
        var xMaxExtent = ReadInt16();
        var caretSlopeRise = ReadInt16();
        var caretSlopeRun = ReadInt16();
        var caretOffset = ReadInt16();
        var reserved1 = ReadInt16();
        var reserved2 = ReadInt16();
        var reserved3 = ReadInt16();
        var reserved4 = ReadInt16();
        var metricDataFormat = ReadInt16();
        _numberOfHMetrics = ReadUInt16();
    }

    private void ParseHmtxTable(uint offset)
    {
        _pos = (int)offset;
        GlyphWidths = new int[_numberOfHMetrics];
        for (int i = 0; i < _numberOfHMetrics; i++)
        {
            GlyphWidths[i] = ReadUInt16();
            ReadInt16(); // Left Side Bearing
        }
    }

    private void ParseNameTable(uint offset)
    {
        _pos = (int)offset;
        var format = ReadUInt16();
        var count = ReadUInt16();
        var stringOffset = ReadUInt16();

        for (int i = 0; i < count; i++)
        {
            var platformID = ReadUInt16();
            var encodingID = ReadUInt16();
            var languageID = ReadUInt16();
            var nameID = ReadUInt16();
            var length = ReadUInt16();
            var nameOffset = ReadUInt16();

            // ??ID 4???????
            if (nameID == 4 && string.IsNullOrEmpty(FontName))
            {
                var savePos = _pos;
                _pos = (int)offset + stringOffset + nameOffset;
                var nameBytes = ReadBytes(length);

                try
                {
                    if (platformID == 3) // Windows??
                    {
                        if (encodingID == 1 || encodingID == 10) // Unicode
                        {
                            FontName = Encoding.BigEndianUnicode.GetString(nameBytes);
                        }
                    }
                    else if (platformID == 1) // Mac??
                    {
                        FontName = Encoding.ASCII.GetString(nameBytes);
                    }
                }
                catch
                {
                    // ??????????
                }

                _pos = savePos;
            }
        }
    }

    private void ParseCmapTable(uint offset)
    {
        _pos = (int)offset;
        var version = ReadUInt16();
        var numTables = ReadUInt16();

        var tables = new List<(ushort platformID, ushort encodingID, uint offset)>();

        for (int i = 0; i < numTables; i++)
        {
            var platformID = ReadUInt16();
            var encodingID = ReadUInt16();
            var subtableOffset = ReadUInt32();
            tables.Add((platformID, encodingID, offset + subtableOffset));
        }

        // 优先选择 Windows Unicode (3, 1) 或 (3, 10)
        // 其次选择 Unicode (0, 3) 或 (0, 4)
        var bestTable = tables
            .OrderByDescending(t => GetCmapPriority(t.platformID, t.encodingID))
            .FirstOrDefault(t => GetCmapPriority(t.platformID, t.encodingID) > 0);

        if (bestTable.offset > 0)
        {
            Console.WriteLine($"Selected CMap: Platform {bestTable.platformID}, Encoding {bestTable.encodingID}");
            ParseCmapSubtable(bestTable.offset);
        }
        else
        {
            Console.WriteLine("No suitable CMap found!");
        }
    }

    private int GetCmapPriority(ushort platformID, ushort encodingID)
    {
        if (platformID == 3 && encodingID == 10) return 100; // Windows Unicode Full
        if (platformID == 3 && encodingID == 1) return 90;   // Windows Unicode BMP
        if (platformID == 0 && encodingID == 4) return 80;   // Unicode 2.0+
        if (platformID == 0 && encodingID == 3) return 70;   // Unicode 2.0 BMP
        if (platformID == 0 && encodingID == 1) return 60;   // Unicode 1.1
        if (platformID == 3 && encodingID == 0) return 10;   // Windows Symbol (Last resort)
        return 0;
    }

    private void ParseCmapSubtable(uint offset)
    {
        _pos = (int)offset;
        var format = ReadUInt16();

        // 格式 4 (Segment mapping to delta values)
        if (format == 4) 
        {
            var length = ReadUInt16();
            var language = ReadUInt16();
            var segCountX2 = ReadUInt16();
            var segCount = segCountX2 / 2;
            var searchRange = ReadUInt16();
            var entrySelector = ReadUInt16();
            var rangeShift = ReadUInt16();

            var endCodes = new ushort[segCount];
            for (int i = 0; i < segCount; i++) endCodes[i] = ReadUInt16();

            var reservedPad = ReadUInt16();

            var startCodes = new ushort[segCount];
            for (int i = 0; i < segCount; i++) startCodes[i] = ReadUInt16();

            var idDeltas = new short[segCount];
            for (int i = 0; i < segCount; i++) idDeltas[i] = ReadInt16();

            // idRangeOffsets array position
            var idRangeOffsetsStartPos = _pos; 
            var idRangeOffsets = new ushort[segCount];
            for (int i = 0; i < segCount; i++) idRangeOffsets[i] = ReadUInt16();

            // 映射逻辑
            for (int seg = 0; seg < segCount; seg++)
            {
                // Safety check for start/end codes
                if (startCodes[seg] > endCodes[seg]) continue;

                for (int c = startCodes[seg]; c <= endCodes[seg]; c++)
                {
                    // 0xFFFF is usually reserved
                    if (c == 0xFFFF) continue;

                    int glyphId;
                    if (idRangeOffsets[seg] == 0)
                    {
                        // idRangeOffset为0，直接使用idDelta
                        glyphId = (c + idDeltas[seg]) & 0xFFFF;
                    }
                    else
                    {
                        // idRangeOffset不为0，依靠偏移量查找glyphIdArray
                        // Address = &idRangeOffsets[seg] + idRangeOffsets[seg] + 2 * (c - startCodes[seg])
                        var currentRangeOffsetAddr = idRangeOffsetsStartPos + seg * 2;
                        var glyphIdAddr = currentRangeOffsetAddr + idRangeOffsets[seg] + 2 * (c - startCodes[seg]);
                        
                        if (glyphIdAddr >= 0 && glyphIdAddr + 2 <= _data.Length)
                        {
                            glyphId = ReadUInt16At((int)glyphIdAddr);
                            if (glyphId != 0)
                            {
                                // 如果查找到的glyphId不为0，再加上idDelta
                                // 注意：这是根据TrueType规范，有些实现可能已经加过了，但规范说要加。
                                // 通常: *(idRangeOffsets[i] + 2*(c-startCode[i]) + &idRangeOffsets[i])
                                // If the value obtained from the indexing operation is not 0 (which indicates missingGlyph), 
                                // idDelta[i] is added to it to get the glyph index.
                                glyphId = (glyphId + idDeltas[seg]) & 0xFFFF;
                            }
                        }
                        else
                        {
                            glyphId = 0;
                        }
                    }

                    if (glyphId != 0)
                    {
                        var ch = (char)c;
                        // 覆盖旧值（如果有）
                        GlyphToUnicode[glyphId] = ch;
                        UnicodeToGlyph[ch] = glyphId;
                    }
                }
            }
        }
        else if (format == 12) // 32?Unicode
        {
            var reserved = ReadUInt16();
            var length = ReadUInt32();
            var language = ReadUInt32();
            var numGroups = ReadUInt32();

            for (uint i = 0; i < numGroups; i++)
            {
                var startCharCode = ReadUInt32();
                var endCharCode = ReadUInt32();
                var startGlyphID = ReadUInt32();

                for (uint c = startCharCode; c <= endCharCode; c++)
                {
                    var glyphId = (int)(startGlyphID + (c - startCharCode));
                    var ch = (char)c;
                    GlyphToUnicode[glyphId] = ch;
                    UnicodeToGlyph[ch] = glyphId;
                }
            }
        }
    }

    // ???????ID
    public int GetGlyphId(char c)
    {
        return UnicodeToGlyph.TryGetValue(c, out var glyphId) ? glyphId : 0;
    }

    // ??????
    public int GetGlyphWidth(int glyphId)
    {
        if (GlyphWidths == null || GlyphWidths.Length == 0) return UnitsPerEm / 2;
        if (glyphId < GlyphWidths.Length) return GlyphWidths[glyphId];
        return GlyphWidths[^1]; // Last metric applies to all subsequent glyphs
    }

    // ??????
    private byte[] ReadBytes(int count)
    {
        var result = new byte[count];
        Array.Copy(_data, _pos, result, 0, count);
        _pos += count;
        return result;
    }

    private byte ReadByte() => _data[_pos++];
    private ushort ReadUInt16() => (ushort)((ReadByte() << 8) | ReadByte());
    private short ReadInt16() => (short)ReadUInt16();
    private uint ReadUInt32() => ((uint)ReadUInt16() << 16) | ReadUInt16();
    private ulong ReadUInt64() => ((ulong)ReadUInt32() << 32) | ReadUInt32();

    private ushort ReadUInt16At(int position)
    {
        return (ushort)((_data[position] << 8) | _data[position + 1]);
    }

    private string ReadTag()
    {
        var bytes = ReadBytes(4);
        return Encoding.ASCII.GetString(bytes);
    }
}
