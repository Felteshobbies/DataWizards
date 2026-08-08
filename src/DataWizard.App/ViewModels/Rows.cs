using System;
using System.Collections.Generic;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;

namespace DataWizard.App.ViewModels;

/// <summary>
/// An editable row for a header field name pattern.
/// </summary>
/// <remarks>
/// The configuration classes are plain data and raise no change notifications, so
/// the grids bind to these wrappers instead and the values are written back when
/// the settings are saved.
/// </remarks>
public partial class NamePatternRow : ViewModelBase
{
    [ObservableProperty]
    public partial bool Enabled { get; set; }

    [ObservableProperty]
    public partial string Pattern { get; set; }

    [ObservableProperty]
    public partial MatchMode Match { get; set; }

    [ObservableProperty]
    public partial bool CaseSensitive { get; set; }

    [ObservableProperty]
    public partial string? Comment { get; set; }

    public NamePatternRow()
    {
        Enabled = true;
        Pattern = string.Empty;
        Match = MatchMode.Contains;
    }

    public NamePatternRow(NamePattern source)
    {
        Enabled = source.Enabled;
        Pattern = source.Pattern;
        Match = source.Match;
        CaseSensitive = source.CaseSensitive;
        Comment = source.Comment;
    }

    /// <summary>All match modes, for the grid's drop-down.</summary>
    public static IReadOnlyList<MatchMode> MatchModes { get; } = Enum.GetValues<MatchMode>();

    /// <summary>Whether the pattern compiles, so invalid regular expressions show up.</summary>
    public bool IsValid => ToModel().IsPatternValid(out _);

    /// <summary>The validation message, or <c>null</c> when the pattern is fine.</summary>
    public string? ValidationError
    {
        get
        {
            ToModel().IsPatternValid(out var error);
            return error;
        }
    }

    public NamePattern ToModel() => new()
    {
        Enabled = Enabled,
        Pattern = Pattern ?? string.Empty,
        Match = Match,
        CaseSensitive = CaseSensitive,
        Comment = Comment
    };
}

/// <summary>An editable row for a data type rule.</summary>
public partial class FieldRuleRow : ViewModelBase
{
    [ObservableProperty]
    public partial bool Enabled { get; set; }

    [ObservableProperty]
    public partial string Pattern { get; set; }

    [ObservableProperty]
    public partial MatchMode Match { get; set; }

    [ObservableProperty]
    public partial bool CaseSensitive { get; set; }

    [ObservableProperty]
    public partial FieldDataType DataType { get; set; }

    [ObservableProperty]
    public partial int? ColumnIndex { get; set; }

    [ObservableProperty]
    public partial string? DateFormat { get; set; }

    [ObservableProperty]
    public partial string? Comment { get; set; }

    public FieldRuleRow()
    {
        Enabled = true;
        Pattern = string.Empty;
        Match = MatchMode.Contains;
        DataType = FieldDataType.Text;
    }

    public FieldRuleRow(FieldRule source)
    {
        Enabled = source.Enabled;
        Pattern = source.Pattern;
        Match = source.Match;
        CaseSensitive = source.CaseSensitive;
        DataType = source.DataType;
        ColumnIndex = source.ColumnIndex;
        DateFormat = source.DateFormat;
        Comment = source.Comment;
    }

    /// <summary>All match modes, for the grid's drop-down.</summary>
    public static IReadOnlyList<MatchMode> MatchModes { get; } = Enum.GetValues<MatchMode>();

    /// <summary>All data types, for the grid's drop-down.</summary>
    public static IReadOnlyList<FieldDataType> DataTypes { get; } = Enum.GetValues<FieldDataType>();

    /// <summary>Whether the pattern compiles.</summary>
    public bool IsValid => ToModel().IsPatternValid(out _);

    public FieldRule ToModel() => new()
    {
        Enabled = Enabled,
        Pattern = Pattern ?? string.Empty,
        Match = Match,
        CaseSensitive = CaseSensitive,
        DataType = DataType,
        ColumnIndex = ColumnIndex,
        DateFormat = string.IsNullOrWhiteSpace(DateFormat) ? null : DateFormat,
        Comment = Comment
    };
}

/// <summary>The state of one file queued for conversion.</summary>
public enum FileState
{
    Pending,
    Running,
    Done,
    Failed
}

/// <summary>A file in the convert list.</summary>
public partial class FileRow : ViewModelBase
{
    [ObservableProperty]
    public partial FileState State { get; set; }

    [ObservableProperty]
    public partial string? Detail { get; set; }

    public FileRow(string path)
    {
        Path = path;
        State = FileState.Pending;
        Direction = ConversionService.DirectionFor(path);
    }

    /// <summary>Full path of the file.</summary>
    public string Path { get; }

    /// <summary>File name, shown in the list.</summary>
    public string FileName => System.IO.Path.GetFileName(Path);

    /// <summary>The folder the file sits in.</summary>
    public string Folder => System.IO.Path.GetDirectoryName(Path) ?? string.Empty;

    /// <summary>Which way this file will be converted.</summary>
    public ConversionDirection? Direction { get; }

    /// <summary>A short description of the conversion, for the list.</summary>
    public string DirectionText => Direction switch
    {
        ConversionDirection.CsvToExcel => "CSV to Excel",
        ConversionDirection.ExcelToCsv => "Excel to CSV",
        _ => "Unsupported"
    };

    /// <summary>A short marker showing where the file got to.</summary>
    public string StateGlyph => State switch
    {
        FileState.Running => "...",
        FileState.Done => "OK",
        FileState.Failed => "!",
        _ => "-"
    };

    /// <summary>
    /// Severity flags for the view. Avalonia toggles a style class from a boolean
    /// binding, which keeps the colours in the theme where they belong instead of
    /// resolving brushes in the view model.
    /// </summary>
    public bool IsDone => State == FileState.Done;

    /// <inheritdoc cref="IsDone"/>
    public bool IsFailed => State == FileState.Failed;

    partial void OnStateChanged(FileState value)
    {
        OnPropertyChanged(nameof(StateGlyph));
        OnPropertyChanged(nameof(IsDone));
        OnPropertyChanged(nameof(IsFailed));
    }

    /// <summary>File size in a readable form.</summary>
    public string SizeText
    {
        get
        {
            try
            {
                var bytes = new FileInfo(Path).Length;
                return bytes switch
                {
                    < 1024 => $"{bytes} B",
                    < 1024 * 1024 => $"{bytes / 1024.0:F1} KB",
                    < 1024L * 1024 * 1024 => $"{bytes / (1024.0 * 1024):F1} MB",
                    _ => $"{bytes / (1024.0 * 1024 * 1024):F1} GB"
                };
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return "unknown";
            }
        }
    }
}

/// <summary>An editable row for a watch rule.</summary>
public partial class WatchRuleRow : ViewModelBase
{
    [ObservableProperty]
    public partial string Name { get; set; }

    [ObservableProperty]
    public partial bool Enabled { get; set; }

    [ObservableProperty]
    public partial string InputFolder { get; set; }

    [ObservableProperty]
    public partial string OutputFolder { get; set; }

    [ObservableProperty]
    public partial string FilePatterns { get; set; }

    [ObservableProperty]
    public partial bool IncludeSubfolders { get; set; }

    [ObservableProperty]
    public partial SourceFileAction SourceAction { get; set; }

    [ObservableProperty]
    public partial string ProcessedFolder { get; set; }

    [ObservableProperty]
    public partial string ErrorFolder { get; set; }

    [ObservableProperty]
    public partial int StabilizationDelayMs { get; set; }

    [ObservableProperty]
    public partial int MaxRetries { get; set; }

    [ObservableProperty]
    public partial int RetryDelayMs { get; set; }

    [ObservableProperty]
    public partial bool ProcessExistingFiles { get; set; }

    [ObservableProperty]
    public partial string? Status { get; set; }

    public WatchRuleRow()
        : this(new WatchRule())
    {
    }

    public WatchRuleRow(WatchRule source)
    {
        Name = source.Name;
        Enabled = source.Enabled;
        InputFolder = source.InputFolder;
        OutputFolder = source.OutputFolder;
        FilePatterns = source.FilePatterns;
        IncludeSubfolders = source.IncludeSubfolders;
        SourceAction = source.SourceAction;
        ProcessedFolder = source.ProcessedFolder;
        ErrorFolder = source.ErrorFolder;
        StabilizationDelayMs = source.StabilizationDelayMs;
        MaxRetries = source.MaxRetries;
        RetryDelayMs = source.RetryDelayMs;
        ProcessExistingFiles = source.ProcessExistingFiles;
    }

    /// <summary>All source actions, for the drop-down.</summary>
    public static IReadOnlyList<SourceFileAction> SourceActions { get; } = Enum.GetValues<SourceFileAction>();

    /// <summary>Whether a processed folder needs to be supplied.</summary>
    public bool NeedsProcessedFolder => SourceAction == SourceFileAction.Move;

    partial void OnSourceActionChanged(SourceFileAction value) =>
        OnPropertyChanged(nameof(NeedsProcessedFolder));

    partial void OnNameChanged(string value) => OnPropertyChanged(nameof(DisplayName));

    /// <summary>Name shown in the rule list, falling back to the folder.</summary>
    public string DisplayName =>
        string.IsNullOrWhiteSpace(Name)
            ? (string.IsNullOrWhiteSpace(InputFolder) ? "New rule" : System.IO.Path.GetFileName(InputFolder.TrimEnd('\\', '/')))
            : Name;

    public WatchRule ToModel() => new()
    {
        Name = Name ?? string.Empty,
        Enabled = Enabled,
        InputFolder = InputFolder ?? string.Empty,
        OutputFolder = OutputFolder ?? string.Empty,
        FilePatterns = FilePatterns ?? "*.csv",
        IncludeSubfolders = IncludeSubfolders,
        SourceAction = SourceAction,
        ProcessedFolder = ProcessedFolder ?? string.Empty,
        ErrorFolder = ErrorFolder ?? string.Empty,
        StabilizationDelayMs = StabilizationDelayMs,
        MaxRetries = MaxRetries,
        RetryDelayMs = RetryDelayMs,
        ProcessExistingFiles = ProcessExistingFiles
    };
}
