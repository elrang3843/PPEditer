using System.IO;
using System.Windows.Media;
using Microsoft.Win32;

namespace PPEditer.Services;

public enum FontLicense
{
    Embeddable,  // fsType 0(무제한) 또는 8(편집 가능)  — 필터 통과
    System,      // C:\Windows\Fonts 기본 제공 글꼴     — 항상 통과
    Restricted,  // fsType 2/4 등 임베딩 제한           — 필터 차단
    Unknown,     // 파일 못 찾거나 읽기 실패             — 필터 차단
}

public sealed record FontInfo(string Name, FontLicense License)
{
    public string LicenseLabel => License switch
    {
        FontLicense.Embeddable => "허용",
        FontLicense.System     => "Windows",
        FontLicense.Restricted => "제한",
        _                      => "?",
    };
}

/// <summary>
/// Resolves installed font families to FontInfo by reading the OS/2 fsType
/// embedding flag directly from the font binary (TTF/OTF/TTC).
/// </summary>
public static class FontService
{
    private static readonly string WinFontsDir =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

    private static readonly string UserFontsDir =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                     @"Microsoft\Windows\Fonts");

    private static List<FontInfo>? _allCache;
    private static List<FontInfo>? _filteredCache;

    // ── Public API ─────────────────────────────────────────────────────

    /// <summary>All installed font families with license classification.</summary>
    public static IReadOnlyList<FontInfo> GetAllFonts()
    {
        if (_allCache is not null) return _allCache;

        var pathMap = BuildFontPathMap();
        _allCache = Fonts.SystemFontFamilies
            .Select(f => f.Source)
            .Where(n => !string.IsNullOrEmpty(n))
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .Select(name => new FontInfo(name, ClassifyFont(name, pathMap)))
            .ToList();

        return _allCache;
    }

    /// <summary>
    /// Fonts allowed by the license filter:
    /// Windows built-in fonts + fonts with fsType == 0 or 8.
    /// </summary>
    public static IReadOnlyList<FontInfo> GetLicenseFilteredFonts()
    {
        if (_filteredCache is not null) return _filteredCache;

        _filteredCache = GetAllFonts()
            .Where(f => f.License is FontLicense.Embeddable or FontLicense.System)
            .ToList();

        return _filteredCache;
    }

    // ── Classification ─────────────────────────────────────────────────

    private static FontLicense ClassifyFont(string familyName,
                                             Dictionary<string, string> pathMap)
    {
        var path = ResolveFontPath(familyName, pathMap);
        if (path is null) return FontLicense.Unknown;

        // Windows 기본 글꼴 폴더에 있으면 무조건 허용
        if (path.StartsWith(WinFontsDir, StringComparison.OrdinalIgnoreCase))
            return FontLicense.System;

        // 그 외: fsType 플래그 확인
        ushort fsType = ReadFsType(path);
        // 하위 4비트가 0(무제한) 또는 8(편집 허용)이면 통과
        return (fsType & 0x000F) is 0 or 8
            ? FontLicense.Embeddable
            : FontLicense.Restricted;
    }

    // ── Font path resolution (registry) ───────────────────────────────

    private static Dictionary<string, string>? _pathMapCache;

    private static Dictionary<string, string> BuildFontPathMap()
    {
        if (_pathMapCache is not null) return _pathMapCache;

        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // 시스템 글꼴 레지스트리
        const string regPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";
        AddRegistryFonts(map, Registry.LocalMachine, regPath, WinFontsDir);

        // 사용자 설치 글꼴 레지스트리
        AddRegistryFonts(map, Registry.CurrentUser, regPath, UserFontsDir);

        _pathMapCache = map;
        return map;
    }

    private static void AddRegistryFonts(Dictionary<string, string> map,
                                          RegistryKey root, string subKey,
                                          string defaultDir)
    {
        try
        {
            using var key = root.OpenSubKey(subKey);
            if (key is null) return;
            foreach (var valueName in key.GetValueNames())
            {
                if (key.GetValue(valueName) is not string fileName) continue;

                var fullPath = Path.IsPathRooted(fileName)
                    ? fileName
                    : Path.Combine(defaultDir, fileName);

                // 레지스트리 값 이름에서 "(TrueType)", "(OpenType)" 등 제거
                var baseName = StripParenthetical(valueName);
                map.TryAdd(baseName, fullPath);
                map.TryAdd(valueName, fullPath);
            }
        }
        catch { }
    }

    private static string? ResolveFontPath(string familyName,
                                            Dictionary<string, string> map)
    {
        // 1. 정확히 일치
        if (map.TryGetValue(familyName, out var p)) return p;

        // 2. 패밀리명으로 시작하는 항목 (예: "Arial Bold (TrueType)" → "Arial")
        foreach (var (key, val) in map)
        {
            if (key.StartsWith(familyName, StringComparison.OrdinalIgnoreCase))
                return val;
        }

        // 3. 직접 파일 탐색 (Windows Fonts 폴더)
        foreach (var ext in new[] { ".ttf", ".otf", ".ttc", ".otc" })
        {
            var candidate = Path.Combine(WinFontsDir,
                familyName.Replace(" ", "") + ext);
            if (File.Exists(candidate)) return candidate;
        }

        return null;
    }

    private static string StripParenthetical(string name)
    {
        var idx = name.LastIndexOf('(');
        return idx > 0 ? name[..idx].Trim() : name.Trim();
    }

    // ── fsType binary reader (OS/2 table) ─────────────────────────────

    private static ushort ReadFsType(string path)
    {
        try
        {
            using var fs = new FileStream(path, FileMode.Open,
                                          FileAccess.Read, FileShare.ReadWrite);
            if (fs.Length < 12) return 0;
            using var br = new BinaryReader(fs);

            uint sig = ReadU32(br);

            if (sig == 0x74746366) // 'ttcf' — TrueType Collection
            {
                ReadU32(br);                    // version
                uint nFonts = ReadU32(br);
                if (nFonts == 0) return 0;
                uint offset = ReadU32(br);      // offset of first font
                fs.Seek(offset, SeekOrigin.Begin);
                ReadU32(br);                    // sfVersion of first font
            }
            // 여기서부터 numTables 위치

            ushort numTables = ReadU16(br);
            br.ReadBytes(6);                    // searchRange + entrySelector + rangeShift

            for (int i = 0; i < numTables; i++)
            {
                byte[] tTag  = br.ReadBytes(4);
                br.ReadBytes(4);                // checkSum
                uint tOffset = ReadU32(br);
                br.ReadBytes(4);                // length

                // OS/2 테이블 찾기
                if (tTag[0] == 'O' && tTag[1] == 'S' &&
                    tTag[2] == '/' && tTag[3] == '2')
                {
                    // OS/2: version(2) xAvgCharWidth(2) usWeightClass(2)
                    //        usWidthClass(2) fsType(2)  → offset + 8
                    fs.Seek(tOffset + 8, SeekOrigin.Begin);
                    return ReadU16(br);
                }
            }
        }
        catch { }

        return 0; // 읽기 실패 → 무제한으로 처리
    }

    private static ushort ReadU16(BinaryReader r)
    {
        var b = r.ReadBytes(2);
        return (ushort)((b[0] << 8) | b[1]);
    }

    private static uint ReadU32(BinaryReader r)
    {
        var b = r.ReadBytes(4);
        return (uint)((b[0] << 24) | (b[1] << 16) | (b[2] << 8) | b[3]);
    }
}
