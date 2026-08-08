namespace DataWizard.Core.Configuration;

/// <summary>
/// A single watched folder and what happens to the files that appear in it.
/// </summary>
public sealed class WatchRule
{
    /// <summary>Display name, used in the UI and in log lines.</summary>
    public string Name { get; set; } = "New rule";

    /// <summary>Whether this rule is currently active.</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>Folder that is monitored for new or changed files.</summary>
    public string InputFolder { get; set; } = string.Empty;

    /// <summary>Where results are written. Empty means next to the source file.</summary>
    public string OutputFolder { get; set; } = string.Empty;

    /// <summary>
    /// Semicolon-separated file masks, for example <c>*.csv;*.txt</c>. Only files
    /// matching one of them are picked up.
    /// </summary>
    public string FilePatterns { get; set; } = "*.csv;*.txt;*.xlsx";

    /// <summary>Also watch folders below <see cref="InputFolder"/>.</summary>
    public bool IncludeSubfolders { get; set; }

    /// <summary>What happens to the source file once it has been converted.</summary>
    public SourceFileAction SourceAction { get; set; } = SourceFileAction.Keep;

    /// <summary>Destination for <see cref="SourceFileAction.Move"/>.</summary>
    public string ProcessedFolder { get; set; } = string.Empty;

    /// <summary>Where a file is moved when its conversion fails. Empty means leave it in place.</summary>
    public string ErrorFolder { get; set; } = string.Empty;

    /// <summary>
    /// How long a file must stop changing before it is converted. Large files
    /// arrive in pieces, and reading one mid-copy produces garbage.
    /// </summary>
    public int StabilizationDelayMs { get; set; } = 1500;

    /// <summary>How often a failed conversion is retried before it is given up on.</summary>
    public int MaxRetries { get; set; } = 3;

    /// <summary>Delay between retries.</summary>
    public int RetryDelayMs { get; set; } = 2000;

    /// <summary>Convert files already present in the folder when the watcher starts.</summary>
    public bool ProcessExistingFiles { get; set; }

    /// <summary>Splits <see cref="FilePatterns"/> into individual masks.</summary>
    public string[] ResolveFilePatterns()
    {
        if (string.IsNullOrWhiteSpace(FilePatterns))
            return ["*.*"];

        var patterns = FilePatterns
            .Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(p => p.Length > 0)
            .ToArray();

        return patterns.Length == 0 ? ["*.*"] : patterns;
    }

    /// <summary>Reports whether the rule is complete enough to run.</summary>
    public bool Validate(out string? error)
    {
        error = null;

        if (string.IsNullOrWhiteSpace(InputFolder))
        {
            error = "Input folder is required.";
            return false;
        }

        if (!Directory.Exists(InputFolder))
        {
            error = $"Input folder does not exist: {InputFolder}";
            return false;
        }

        if (SourceAction == SourceFileAction.Move && string.IsNullOrWhiteSpace(ProcessedFolder))
        {
            error = "A processed folder is required when the source action is Move.";
            return false;
        }

        if (StabilizationDelayMs is < 0 or > 600_000)
        {
            error = "Stabilization delay must be between 0 and 600000 ms.";
            return false;
        }

        return true;
    }

    /// <summary>Creates a copy for editing.</summary>
    public WatchRule Clone() => (WatchRule)MemberwiseClone();

    public override string ToString() => Name;
}

/// <summary>
/// The watcher as a whole: whether it runs, and the rules it runs.
/// </summary>
public sealed class WatcherSettings
{
    /// <summary>Start watching as soon as the application launches.</summary>
    public bool AutoStart { get; set; }

    /// <summary>The configured folder rules.</summary>
    public List<WatchRule> Rules { get; set; } = [];

    /// <summary>Creates a copy for editing.</summary>
    public WatcherSettings Clone() => new()
    {
        AutoStart = AutoStart,
        Rules = Rules.Select(r => r.Clone()).ToList()
    };
}
