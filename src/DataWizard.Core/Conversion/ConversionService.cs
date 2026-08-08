using System.Diagnostics;
using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;
using DataWizard.Core.Excel;

namespace DataWizard.Core.Conversion;

/// <summary>
/// Runs conversions in both directions, choosing the direction from the file
/// extension and reporting everything it does through <see cref="Log"/>.
/// </summary>
public sealed class ConversionService
{
    /// <summary>Extensions treated as delimited text.</summary>
    public static readonly string[] CsvExtensions = [".csv", ".txt", ".log", ".tsv", ".dat"];

    /// <summary>Extensions treated as workbooks.</summary>
    public static readonly string[] ExcelExtensions = [".xlsx", ".xlsm"];

    private DataWizardSettings _settings;

    public ConversionService(DataWizardSettings settings)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
    }

    /// <summary>Raised for every log line produced during a conversion.</summary>
    public event EventHandler<LogEntry>? Log;

    /// <summary>The settings used for subsequent conversions.</summary>
    public DataWizardSettings Settings
    {
        get => _settings;
        set => _settings = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>All extensions the application accepts.</summary>
    public static IEnumerable<string> SupportedExtensions => CsvExtensions.Concat(ExcelExtensions);

    /// <summary>Whether a path looks like something that can be converted.</summary>
    public static bool CanConvert(string path) => DirectionFor(path) is not null;

    /// <summary>
    /// The direction implied by a file extension, or <c>null</c> when the
    /// extension is not supported.
    /// </summary>
    public static ConversionDirection? DirectionFor(string path)
    {
        var extension = Path.GetExtension(path);

        if (CsvExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            return ConversionDirection.CsvToExcel;

        if (ExcelExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            return ConversionDirection.ExcelToCsv;

        return null;
    }

    /// <summary>
    /// Converts one file.
    /// </summary>
    /// <param name="sourcePath">The file to convert.</param>
    /// <param name="outputFolderOverride">
    /// Target folder for this call, overriding
    /// <see cref="ConversionSettings.OutputFolder"/>. Used by watch rules.
    /// </param>
    /// <param name="progress">Reports rows handled.</param>
    /// <param name="cancellationToken">Cancels the conversion.</param>
    public ConversionResult Convert(
        string sourcePath,
        string? outputFolderOverride = null,
        IProgress<ConversionProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();
        var direction = DirectionFor(sourcePath);

        if (direction is null)
        {
            return Failure(
                sourcePath,
                ConversionDirection.CsvToExcel,
                $"Unsupported file type '{Path.GetExtension(sourcePath)}'.",
                null,
                stopwatch.Elapsed);
        }

        if (!File.Exists(sourcePath))
        {
            return Failure(sourcePath, direction.Value, "File not found.", null, stopwatch.Elapsed);
        }

        try
        {
            return direction.Value == ConversionDirection.CsvToExcel
                ? ConvertCsvToExcel(sourcePath, outputFolderOverride, progress, stopwatch, cancellationToken)
                : ConvertExcelToCsv(sourcePath, outputFolderOverride, progress, stopwatch, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            WriteLog(LogLevel.Warning, $"Cancelled: {Path.GetFileName(sourcePath)}", sourcePath);

            return new ConversionResult
            {
                SourcePath = sourcePath,
                Direction = direction.Value,
                Success = false,
                WasCancelled = true,
                ErrorMessage = "Cancelled.",
                Duration = stopwatch.Elapsed
            };
        }
        catch (Exception ex)
        {
            return Failure(sourcePath, direction.Value, ex.Message, ex, stopwatch.Elapsed);
        }
    }

    /// <summary>Runs <see cref="Convert"/> on a background thread.</summary>
    public Task<ConversionResult> ConvertAsync(
        string sourcePath,
        string? outputFolderOverride = null,
        IProgress<ConversionProgress>? progress = null,
        CancellationToken cancellationToken = default) =>
        Task.Run(
            () => Convert(sourcePath, outputFolderOverride, progress, cancellationToken),
            cancellationToken);

    /// <summary>
    /// Analyses a CSV file without converting it. Used by the preview panel.
    /// </summary>
    public CsvAnalysis Analyze(string sourcePath) =>
        new CsvAnalyzer(_settings.Detection).Analyze(sourcePath);

    private ConversionResult ConvertCsvToExcel(
        string sourcePath,
        string? outputFolderOverride,
        IProgress<ConversionProgress>? progress,
        Stopwatch stopwatch,
        CancellationToken cancellationToken)
    {
        WriteLog(LogLevel.Info, $"Analysing {Path.GetFileName(sourcePath)}", sourcePath);

        var analysis = new CsvAnalyzer(_settings.Detection).Analyze(sourcePath);

        foreach (var line in analysis.ToLogLines())
            WriteLog(LogLevel.Detail, line, sourcePath);

        foreach (var warning in analysis.Warnings)
            WriteLog(LogLevel.Warning, warning, sourcePath);

        var outputPath = BuildOutputPath(sourcePath, ".xlsx", outputFolderOverride);

        var rowProgress = progress is null
            ? null
            : new Progress<int>(rows => progress.Report(
                new ConversionProgress(sourcePath, rows, analysis.TotalLineCount)));

        var writer = new ExcelWriter(_settings.Conversion, _settings.Detection);
        var result = writer.Write(analysis, outputPath, rowProgress, cancellationToken);

        WriteLog(LogLevel.Success, $"Wrote {Path.GetFileName(outputPath)} ({result.RowCount} rows)", outputPath);

        return new ConversionResult
        {
            SourcePath = sourcePath,
            Direction = ConversionDirection.CsvToExcel,
            Success = true,
            OutputPaths = [outputPath],
            Analysis = analysis,
            RowCount = result.RowCount,
            Duration = stopwatch.Elapsed
        };
    }

    private ConversionResult ConvertExcelToCsv(
        string sourcePath,
        string? outputFolderOverride,
        IProgress<ConversionProgress>? progress,
        Stopwatch stopwatch,
        CancellationToken cancellationToken)
    {
        WriteLog(LogLevel.Info, $"Reading {Path.GetFileName(sourcePath)}", sourcePath);

        var outputPath = BuildOutputPath(sourcePath, ".csv", outputFolderOverride);

        var rowProgress = progress is null
            ? null
            : new Progress<int>(rows => progress.Report(new ConversionProgress(sourcePath, rows, null)));

        var reader = new ExcelReader(_settings.Conversion);
        var exports = reader.ToCsv(sourcePath, outputPath, rowProgress, cancellationToken);

        foreach (var export in exports)
        {
            WriteLog(
                LogLevel.Success,
                $"Wrote {Path.GetFileName(export.Path)} from sheet '{export.SheetName}' ({export.RowCount} rows)",
                export.Path);
        }

        return new ConversionResult
        {
            SourcePath = sourcePath,
            Direction = ConversionDirection.ExcelToCsv,
            Success = true,
            OutputPaths = exports.Select(e => e.Path).ToArray(),
            RowCount = exports.Sum(e => e.RowCount),
            Duration = stopwatch.Elapsed
        };
    }

    /// <summary>
    /// Works out where a result goes. Without overwrite enabled a numeric suffix
    /// is added rather than replacing an existing file.
    /// </summary>
    public string BuildOutputPath(string sourcePath, string extension, string? outputFolderOverride = null)
    {
        var folder = outputFolderOverride;

        if (string.IsNullOrWhiteSpace(folder))
            folder = _settings.Conversion.OutputFolder;

        if (string.IsNullOrWhiteSpace(folder))
            folder = Path.GetDirectoryName(sourcePath);

        if (string.IsNullOrWhiteSpace(folder))
            folder = Directory.GetCurrentDirectory();

        var stem = Path.GetFileNameWithoutExtension(sourcePath);
        var candidate = Path.Combine(folder, stem + extension);

        if (_settings.Conversion.Overwrite || !File.Exists(candidate))
            return candidate;

        // Bounded so a folder full of collisions cannot spin indefinitely.
        for (var counter = 1; counter < 10_000; counter++)
        {
            candidate = Path.Combine(folder, $"{stem}_{counter}{extension}");

            if (!File.Exists(candidate))
                return candidate;
        }

        throw new IOException(
            $"Could not find a free output name for '{stem}{extension}' in '{folder}' after 10000 attempts.");
    }

    /// <summary>
    /// Works out where a conversion would write, without performing it.
    /// </summary>
    /// <remarks>
    /// The folder watcher uses this to recognise its own results before they exist.
    /// A workbook exported sheet by sheet produces one file per worksheet, so the
    /// sheet names are read to predict them; if that fails the base name is
    /// returned rather than throwing, since this is only ever advisory.
    /// </remarks>
    public IReadOnlyList<string> PredictOutputPaths(string sourcePath, string? outputFolderOverride = null)
    {
        var direction = DirectionFor(sourcePath);

        if (direction is null)
            return [];

        try
        {
            if (direction == ConversionDirection.CsvToExcel)
                return [BuildOutputPath(sourcePath, ".xlsx", outputFolderOverride)];

            var basePath = BuildOutputPath(sourcePath, ".csv", outputFolderOverride);

            if (!_settings.Conversion.ExportAllSheets)
                return [basePath];

            var sheets = ExcelReader.ListSheetNames(sourcePath);

            return sheets.Count == 0
                ? [basePath]
                : sheets.Select(name => ExcelReader.InsertSheetName(basePath, name)).ToArray();
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException)
        {
            return [];
        }
    }

    private ConversionResult Failure(
        string sourcePath,
        ConversionDirection direction,
        string message,
        Exception? exception,
        TimeSpan duration)
    {
        WriteLog(LogLevel.Error, $"{Path.GetFileName(sourcePath)}: {message}", sourcePath);

        return new ConversionResult
        {
            SourcePath = sourcePath,
            Direction = direction,
            Success = false,
            ErrorMessage = message,
            Exception = exception,
            Duration = duration
        };
    }

    private void WriteLog(LogLevel level, string message, string? filePath = null) =>
        Log?.Invoke(this, new LogEntry(DateTimeOffset.Now, level, message, filePath));
}
