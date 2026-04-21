using System.IO;
using System.Text.Json;
using System.Windows;

namespace PPEditer.Services;

public sealed class AppSettings
{
    private static readonly string ConfigDir =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PPEditer");
    private static readonly string ConfigFile = Path.Combine(ConfigDir, "settings.json");

    private static AppSettings? _instance;
    public static AppSettings Current => _instance ??= Load();

    // ── Persisted properties ───────────────────────────────────────────
    public bool   FontLicenseFilterEnabled { get; set; } = true;
    public string Language                  { get; set; } = "ko";
    public string Theme                     { get; set; } = "Light";

    // ── Load / Save ────────────────────────────────────────────────────
    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(ConfigFile))
            {
                var json = File.ReadAllText(ConfigFile);
                return JsonSerializer.Deserialize<AppSettings>(json) ?? new();
            }
        }
        catch { }
        return new();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(ConfigDir);
            File.WriteAllText(ConfigFile,
                JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }
}

// ── Theme switcher ─────────────────────────────────────────────────────────

public static class AppTheme
{
    private const string UriPrefix = "pack://application:,,,/Resources/Theme.";

    public static void Apply(string theme)
    {
        var uri   = new Uri($"{UriPrefix}{theme}.xaml", UriKind.Absolute);
        var dicts = Application.Current.Resources.MergedDictionaries;
        var old   = dicts.FirstOrDefault(d => d.Source?.ToString().Contains("/Theme.") ?? false);
        int idx   = old is null ? 0 : dicts.IndexOf(old);
        if (old is not null) dicts.Remove(old);
        dicts.Insert(Math.Min(idx, dicts.Count), new ResourceDictionary { Source = uri });
    }
}

// ── Language switcher ──────────────────────────────────────────────────────

public static class AppLanguage
{
    private const string UriPrefix = "pack://application:,,,/Resources/Strings.";

    public static void Apply(string lang)
    {
        var uri   = new Uri($"{UriPrefix}{lang}.xaml", UriKind.Absolute);
        var dicts = Application.Current.Resources.MergedDictionaries;
        var old   = dicts.FirstOrDefault(d => d.Source?.ToString().Contains("/Strings.") ?? false);
        int idx   = old is null ? dicts.Count : dicts.IndexOf(old);
        if (old is not null) dicts.Remove(old);
        dicts.Insert(Math.Min(idx, dicts.Count), new ResourceDictionary { Source = uri });
    }
}
