using System.Windows.Media;

namespace PPEditer.Services;

public enum FontLicense
{
    SilOfl,         // SIL Open Font License
    Apache2,        // Apache 2.0
    UbuntuFont,     // Ubuntu Font Licence
    BitstreamVera,  // Bitstream Vera License (DejaVu 기반)
    MIT,
    Free,           // 무료 사용 허용 (상업용 포함)
    System,         // Windows/OS 내장 — 문서 작성 사용 가능, 재배포 불가
}

public sealed record FontInfo(string Name, FontLicense License)
{
    public bool IsOpenSource => License is not FontLicense.System;

    public string LicenseLabel => License switch
    {
        FontLicense.SilOfl        => "OFL",
        FontLicense.Apache2       => "Apache 2.0",
        FontLicense.UbuntuFont    => "Ubuntu",
        FontLicense.BitstreamVera => "Bitstream Vera",
        FontLicense.MIT           => "MIT",
        FontLicense.Free          => "무료",
        FontLicense.System        => "시스템",
        _                         => "?",
    };
}

/// <summary>
/// Provides a font list filtered to license-compatible (open-source / free) fonts,
/// cross-referenced against the fonts actually installed on the system.
/// </summary>
public static class FontService
{
    // ── Known open-source / free fonts ────────────────────────────────
    // Key = font family name as reported by WPF (case-insensitive).
    // The list intentionally covers both Latin and Korean open-source families.
    private static readonly Dictionary<string, FontLicense> KnownOpenFonts =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // ── Korean ────────────────────────────────────────────────
            { "나눔고딕",              FontLicense.SilOfl },
            { "NanumGothic",          FontLicense.SilOfl },
            { "나눔명조",              FontLicense.SilOfl },
            { "NanumMyeongjo",        FontLicense.SilOfl },
            { "나눔바른고딕",          FontLicense.SilOfl },
            { "NanumBarunGothic",     FontLicense.SilOfl },
            { "나눔고딕코딩",          FontLicense.SilOfl },
            { "NanumGothicCoding",    FontLicense.SilOfl },
            { "나눔바른펜",            FontLicense.SilOfl },
            { "나눔손글씨 붓",         FontLicense.SilOfl },
            { "나눔스퀘어",            FontLicense.SilOfl },
            { "나눔스퀘어라운드",       FontLicense.SilOfl },

            { "Noto Sans KR",         FontLicense.SilOfl },
            { "Noto Serif KR",        FontLicense.SilOfl },
            { "Noto Sans CJK KR",     FontLicense.SilOfl },
            { "Noto Serif CJK KR",    FontLicense.SilOfl },

            { "본고딕",                FontLicense.SilOfl },
            { "본명조",                FontLicense.SilOfl },
            { "Source Han Sans K",    FontLicense.SilOfl },
            { "Source Han Serif K",   FontLicense.SilOfl },
            { "Source Han Sans",      FontLicense.SilOfl },
            { "Source Han Serif",     FontLicense.SilOfl },

            { "KoPub돋움체",           FontLicense.Free },
            { "KoPub바탕체",           FontLicense.Free },
            { "KoPubWorld돋움체",      FontLicense.Free },
            { "KoPubWorld바탕체",      FontLicense.Free },

            { "Spoqa Han Sans",       FontLicense.SilOfl },
            { "Spoqa Han Sans Neo",   FontLicense.SilOfl },
            { "Pretendard",           FontLicense.SilOfl },
            { "Pretendard Variable",  FontLicense.SilOfl },
            { "IBM Plex Sans KR",     FontLicense.SilOfl },
            { "Gmarket Sans",         FontLicense.Free },
            { "고도체",                FontLicense.Free },
            { "고도 마음체",            FontLicense.Free },
            { "S-Core Dream",         FontLicense.Free },
            { "에스코어 드림",          FontLicense.Free },
            { "부산체",                FontLicense.Free },
            { "여기어때 잘난체",        FontLicense.Free },
            { "Sandoll 삼립호빵체",    FontLicense.Free },
            { "BMDOHYEON",            FontLicense.Free },
            { "BMJUA",                FontLicense.Free },
            { "BMKIRANGHAERANG",      FontLicense.Free },
            { "BMEULJIROTTF",         FontLicense.Free },
            { "경기천년제목V",          FontLicense.Free },
            { "경기천년제목M",          FontLicense.Free },
            { "경기천년바탕V",          FontLicense.Free },

            // ── Latin open-source ─────────────────────────────────────
            { "Liberation Sans",      FontLicense.SilOfl },
            { "Liberation Serif",     FontLicense.SilOfl },
            { "Liberation Mono",      FontLicense.SilOfl },

            { "DejaVu Sans",          FontLicense.BitstreamVera },
            { "DejaVu Serif",         FontLicense.BitstreamVera },
            { "DejaVu Sans Mono",     FontLicense.BitstreamVera },
            { "DejaVu Sans Condensed",FontLicense.BitstreamVera },

            { "Open Sans",            FontLicense.Apache2 },
            { "Roboto",               FontLicense.Apache2 },
            { "Roboto Condensed",     FontLicense.Apache2 },
            { "Roboto Mono",          FontLicense.Apache2 },
            { "Roboto Slab",          FontLicense.Apache2 },

            { "Lato",                 FontLicense.SilOfl },
            { "Montserrat",           FontLicense.SilOfl },
            { "Raleway",              FontLicense.SilOfl },
            { "Oswald",               FontLicense.SilOfl },
            { "Nunito",               FontLicense.SilOfl },
            { "Nunito Sans",          FontLicense.SilOfl },
            { "Poppins",              FontLicense.SilOfl },
            { "Inter",                FontLicense.SilOfl },
            { "Merriweather",         FontLicense.SilOfl },
            { "Playfair Display",     FontLicense.SilOfl },
            { "Crimson Text",         FontLicense.SilOfl },
            { "Inconsolata",          FontLicense.SilOfl },
            { "Fira Sans",            FontLicense.SilOfl },
            { "Fira Code",            FontLicense.SilOfl },
            { "Fira Mono",            FontLicense.SilOfl },
            { "PT Sans",              FontLicense.SilOfl },
            { "PT Serif",             FontLicense.SilOfl },
            { "PT Mono",              FontLicense.SilOfl },

            { "Source Sans Pro",      FontLicense.SilOfl },
            { "Source Serif Pro",     FontLicense.SilOfl },
            { "Source Code Pro",      FontLicense.SilOfl },
            { "Source Sans 3",        FontLicense.SilOfl },
            { "Source Serif 4",       FontLicense.SilOfl },

            { "IBM Plex Sans",        FontLicense.SilOfl },
            { "IBM Plex Serif",       FontLicense.SilOfl },
            { "IBM Plex Mono",        FontLicense.SilOfl },

            { "Noto Sans",            FontLicense.SilOfl },
            { "Noto Serif",           FontLicense.SilOfl },
            { "Noto Mono",            FontLicense.SilOfl },

            { "Ubuntu",               FontLicense.UbuntuFont },
            { "Ubuntu Mono",          FontLicense.UbuntuFont },
            { "Ubuntu Condensed",     FontLicense.UbuntuFont },

            { "Anonymous Pro",        FontLicense.SilOfl },
            { "Overpass",             FontLicense.SilOfl },
            { "Barlow",               FontLicense.SilOfl },
            { "Cabin",                FontLicense.SilOfl },
            { "Josefin Sans",         FontLicense.SilOfl },
            { "Exo 2",                FontLicense.SilOfl },
            { "Work Sans",            FontLicense.SilOfl },
            { "DM Sans",              FontLicense.SilOfl },
            { "Space Grotesk",        FontLicense.SilOfl },
            { "Plus Jakarta Sans",    FontLicense.SilOfl },
            { "Readex Pro",           FontLicense.SilOfl },

            // Monospace / code
            { "Hack",                 FontLicense.MIT },
            { "JetBrains Mono",       FontLicense.SilOfl },
            { "Cascadia Code",        FontLicense.SilOfl },
            { "Cascadia Mono",        FontLicense.SilOfl },
        };

    // ── Public API ─────────────────────────────────────────────────────

    private static List<FontInfo>? _allCache;

    /// <summary>Returns all system fonts, tagged as open-source or System.</summary>
    public static IReadOnlyList<FontInfo> GetAllFonts()
    {
        if (_allCache is not null) return _allCache;

        _allCache = Fonts.SystemFontFamilies
            .Select(f => f.Source)
            .Where(n => !string.IsNullOrEmpty(n))    // anonymous composite fonts may have null Source
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .Select(name => KnownOpenFonts.TryGetValue(name, out var lic)
                ? new FontInfo(name, lic)
                : new FontInfo(name, FontLicense.System))
            .ToList();

        return _allCache;
    }

    /// <summary>Returns only open-source / free fonts installed on the system.</summary>
    public static IReadOnlyList<FontInfo> GetOpenSourceFonts()
        => GetAllFonts().Where(f => f.IsOpenSource).ToList();

    /// <summary>
    /// Returns open-source fonts if any are installed, otherwise falls back to all fonts.
    /// This prevents showing an empty list on systems with no open-source fonts.
    /// </summary>
    public static IReadOnlyList<FontInfo> GetRecommendedFonts()
    {
        var open = GetOpenSourceFonts();
        return open.Count > 0 ? open : GetAllFonts();
    }
}
