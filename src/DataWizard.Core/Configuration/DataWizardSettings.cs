using System.Text.Json;
using System.Text.Json.Serialization;

namespace DataWizard.Core.Configuration;

/// <summary>
/// The complete application configuration. Persisted as JSON so that the nested
/// rule lists stay readable and hand-editable.
/// </summary>
public sealed class DataWizardSettings
{
    /// <summary>Schema version, so future releases can migrate old files.</summary>
    public int Version { get; set; } = 1;

    /// <summary>Colour scheme of the application window.</summary>
    public ThemeMode Theme { get; set; } = ThemeMode.System;

    /// <summary>
    /// Whether the analysis panel on the Convert tab starts open.
    /// </summary>
    /// <remarks>
    /// Closed by default. The panel explains every detection decision, which is
    /// exactly what is wanted when a file reads wrongly and pure noise the rest of
    /// the time. It opens by itself when the analysis is ambiguous.
    /// </remarks>
    public bool ShowAnalysisPanel { get; set; }

    /// <summary>Width of the analysis panel in pixels, so a resize is remembered.</summary>
    public double AnalysisPanelWidth { get; set; } = 500d;

    /// <summary>How CSV files are interpreted.</summary>
    public DetectionSettings Detection { get; set; } = new();

    /// <summary>How results are written.</summary>
    public ConversionSettings Conversion { get; set; } = new();

    /// <summary>Folder monitoring rules.</summary>
    public WatcherSettings Watcher { get; set; } = new();

    /// <summary>Profile bookkeeping of the Diff tab.</summary>
    public DiffSettings Diff { get; set; } = new();

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>
    /// Standard location: <c>%APPDATA%\DataWizards\settings.json</c>.
    /// </summary>
    public static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "DataWizards",
        "settings.json");

    /// <summary>
    /// A configuration with the built-in patterns filled in.
    /// </summary>
    public static DataWizardSettings CreateDefault()
    {
        var settings = new DataWizardSettings();
        settings.Detection.KnownFieldNames = DefaultPatterns.CreateKnownFieldNames();
        settings.Detection.FieldRules = DefaultPatterns.CreateFieldRules();
        return settings;
    }

    /// <summary>
    /// Loads settings from disk, falling back to the defaults when the file is
    /// missing or unreadable. A broken settings file must never stop the
    /// application from starting.
    /// </summary>
    public static DataWizardSettings Load(string? path = null)
    {
        path ??= DefaultPath;

        if (!File.Exists(path))
            return CreateDefault();

        try
        {
            var json = File.ReadAllText(path);
            var settings = JsonSerializer.Deserialize<DataWizardSettings>(json, JsonOptions);

            if (settings is null)
                return CreateDefault();

            settings.Detection ??= new DetectionSettings();
            settings.Conversion ??= new ConversionSettings();
            settings.Watcher ??= new WatcherSettings();
            settings.Diff ??= new DiffSettings();
            settings.Normalize();
            return settings;
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            return CreateDefault();
        }
    }

    /// <summary>
    /// Loads settings and reports why the defaults were used, for cases where
    /// silently discarding a broken file would be confusing.
    /// </summary>
    public static DataWizardSettings Load(string? path, out string? warning)
    {
        path ??= DefaultPath;
        warning = null;

        if (!File.Exists(path))
            return CreateDefault();

        try
        {
            var json = File.ReadAllText(path);
            var settings = JsonSerializer.Deserialize<DataWizardSettings>(json, JsonOptions);

            if (settings is null)
            {
                warning = $"Settings file '{path}' was empty; defaults loaded.";
                return CreateDefault();
            }

            settings.Detection ??= new DetectionSettings();
            settings.Conversion ??= new ConversionSettings();
            settings.Watcher ??= new WatcherSettings();
            settings.Diff ??= new DiffSettings();
            settings.Normalize();
            return settings;
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            warning = $"Settings file '{path}' could not be read ({ex.Message}); defaults loaded.";
            return CreateDefault();
        }
    }

    /// <summary>
    /// Writes settings to disk. The file is written to a temporary path first and
    /// then moved into place, so an interrupted save cannot leave a half-written
    /// configuration behind.
    /// </summary>
    public void Save(string? path = null)
    {
        path ??= DefaultPath;

        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        Normalize();

        var json = JsonSerializer.Serialize(this, JsonOptions);
        var temp = path + ".tmp";

        File.WriteAllText(temp, json);
        File.Move(temp, path, overwrite: true);
    }

    /// <summary>Clamps every section into a usable range.</summary>
    public void Normalize()
    {
        Detection.Normalize();
        Conversion.Normalize();
        Watcher.Rules ??= [];
        Diff.Normalize();
        AnalysisPanelWidth = Math.Clamp(AnalysisPanelWidth, 320d, 900d);
    }

    /// <summary>Creates a deep copy, so the settings UI can cancel out of edits.</summary>
    public DataWizardSettings Clone() => new()
    {
        Version = Version,
        Theme = Theme,
        ShowAnalysisPanel = ShowAnalysisPanel,
        AnalysisPanelWidth = AnalysisPanelWidth,
        Detection = Detection.Clone(),
        Conversion = Conversion.Clone(),
        Watcher = Watcher.Clone(),
        Diff = Diff.Clone()
    };
}
