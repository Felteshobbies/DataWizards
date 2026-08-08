using DataWizard.Core.Csv;

namespace DataWizard.Core.Conversion;

/// <summary>Which way a file is being converted.</summary>
public enum ConversionDirection
{
    /// <summary>A delimited text file becomes a workbook.</summary>
    CsvToExcel,
    /// <summary>A workbook becomes one or more delimited text files.</summary>
    ExcelToCsv
}

/// <summary>Severity of a log entry.</summary>
public enum LogLevel
{
    Detail,
    Info,
    Success,
    Warning,
    Error
}

/// <summary>One line in the activity log.</summary>
/// <param name="Timestamp">When it happened.</param>
/// <param name="Level">How important it is.</param>
/// <param name="Message">What happened.</param>
/// <param name="FilePath">The file it refers to, when there is one.</param>
public readonly record struct LogEntry(
    DateTimeOffset Timestamp,
    LogLevel Level,
    string Message,
    string? FilePath = null)
{
    /// <summary>Renders the entry the way the log view shows it.</summary>
    public override string ToString() => $"[{Timestamp:HH:mm:ss}] {Message}";
}

/// <summary>Progress of a single file conversion.</summary>
/// <param name="FilePath">The file being converted.</param>
/// <param name="RowsProcessed">Rows handled so far.</param>
/// <param name="TotalRows">Expected total, or <c>null</c> when not yet known.</param>
public readonly record struct ConversionProgress(string FilePath, int RowsProcessed, long? TotalRows)
{
    /// <summary>Completion between 0 and 1, or <c>null</c> when the total is unknown.</summary>
    public double? Fraction =>
        TotalRows is > 0 ? Math.Clamp(RowsProcessed / (double)TotalRows.Value, 0d, 1d) : null;
}

/// <summary>The outcome of converting one file.</summary>
public sealed class ConversionResult
{
    /// <summary>The file that was converted.</summary>
    public required string SourcePath { get; init; }

    /// <summary>Which direction was used.</summary>
    public required ConversionDirection Direction { get; init; }

    /// <summary>Whether the conversion completed.</summary>
    public required bool Success { get; init; }

    /// <summary>The files that were produced.</summary>
    public IReadOnlyList<string> OutputPaths { get; init; } = [];

    /// <summary>Why it failed, when it did.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>The exception behind a failure, for the detailed log.</summary>
    public Exception? Exception { get; init; }

    /// <summary>Analysis of the source, present for CSV sources.</summary>
    public CsvAnalysis? Analysis { get; init; }

    /// <summary>Rows written.</summary>
    public int RowCount { get; init; }

    /// <summary>How long it took.</summary>
    public TimeSpan Duration { get; init; }

    /// <summary>True when the conversion was cancelled rather than failing.</summary>
    public bool WasCancelled { get; init; }

    /// <summary>A one-line summary for the log.</summary>
    public string Describe()
    {
        var name = Path.GetFileName(SourcePath);

        if (WasCancelled)
            return $"{name}: cancelled";

        if (!Success)
            return $"{name}: failed - {ErrorMessage}";

        var outputs = OutputPaths.Count == 1
            ? Path.GetFileName(OutputPaths[0])
            : $"{OutputPaths.Count} files";

        return $"{name} -> {outputs} ({RowCount} rows, {Duration.TotalSeconds:F1}s)";
    }
}
