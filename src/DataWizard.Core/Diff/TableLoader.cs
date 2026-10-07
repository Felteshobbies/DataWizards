using System.Text;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;
using DataWizard.Core.Csv;
using DataWizard.Core.Excel;

namespace DataWizard.Core.Diff;

/// <summary>
/// Reads a CSV or XLSX file into a <see cref="LoadedTable"/>, the shape the diff
/// engine compares.
/// </summary>
/// <remarks>
/// Delimited text goes through the same analysis pipeline as the converter, so
/// encoding, separator and header detection behave exactly as the user expects
/// from the rest of the application. Workbooks are read with
/// <see cref="ExcelReader.ReadSheet"/>, whose normalised values make a workbook
/// and a CSV written in another culture comparable.
/// </remarks>
public sealed class TableLoader
{
    /// <summary>How many leading worksheet rows are scored for a header.</summary>
    private const int ExcelHeaderSampleRows = 50;

    private readonly DetectionSettings _settings;

    public TableLoader(DetectionSettings settings)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _settings.Normalize();
    }

    /// <summary>
    /// Loads a file, choosing the reader from its extension.
    /// </summary>
    /// <param name="path">A CSV or XLSX file.</param>
    /// <param name="sheetIndex">Which worksheet of a workbook to load.</param>
    public LoadedTable Load(string path, int sheetIndex = 0)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("File not found.", path);

        var extension = Path.GetExtension(path);

        return ConversionService.ExcelExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase)
            ? LoadExcel(path, sheetIndex)
            : LoadCsv(path);
    }

    private LoadedTable LoadCsv(string path)
    {
        var analysis = new CsvAnalyzer(_settings).Analyze(path);
        var records = ReadAllRecords(path, analysis);

        var hasHeader = analysis.HeaderResult.HasHeader;
        var dataRecords = hasHeader ? records.Skip(1).ToList() : records;
        var columnCount = records.Count == 0 ? 0 : records.Max(r => r.FieldCount);

        var header = new string[columnCount];

        for (var i = 0; i < columnCount; i++)
        {
            header[i] = hasHeader && i < analysis.HeaderResult.FieldNames.Length
                ? analysis.HeaderResult.FieldNames[i]
                : $"Column {i + 1}";
        }

        var rows = new string[dataRecords.Count][];

        for (var i = 0; i < dataRecords.Count; i++)
        {
            var record = dataRecords[i];
            var row = new string[columnCount];

            for (var column = 0; column < columnCount; column++)
                row[column] = record.ValueAt(column);

            rows[i] = row;
        }

        return new LoadedTable
        {
            SourcePath = path,
            SheetName = null,
            HasHeader = hasHeader,
            Header = header,
            Rows = rows,
            Warnings = analysis.Warnings
        };
    }

    private LoadedTable LoadExcel(string path, int sheetIndex)
    {
        var grid = ExcelReader.ReadSheet(path, sheetIndex);
        var warnings = new List<string>();

        if (grid.Rows.Length == 0)
            warnings.Add("The worksheet contains no rows.");

        var sample = grid.Rows
            .Take(ExcelHeaderSampleRows)
            .Select((row, index) => new CsvRecord
            {
                Fields = row.Select(value => new CsvField(value, false)).ToArray(),
                StartLine = index + 1,
                LineCount = 1
            })
            .ToList();

        var headerResult = new HeaderDetector(_settings).Detect(sample);

        var hasHeader = headerResult.HasHeader;
        var dataRows = hasHeader ? grid.Rows.Skip(1).ToArray() : grid.Rows;
        var columnCount = grid.Rows.Length == 0 ? 0 : grid.Rows.Max(r => r.Length);

        var header = new string[columnCount];

        for (var i = 0; i < columnCount; i++)
        {
            header[i] = hasHeader && i < headerResult.FieldNames.Length
                ? headerResult.FieldNames[i]
                : $"Column {i + 1}";
        }

        var rows = new string[dataRows.Length][];

        for (var i = 0; i < dataRows.Length; i++)
        {
            var source = dataRows[i];
            var row = new string[columnCount];

            for (var column = 0; column < columnCount; column++)
                row[column] = column < source.Length ? source[column] : string.Empty;

            rows[i] = row;
        }

        return new LoadedTable
        {
            SourcePath = path,
            SheetName = grid.SheetName,
            HasHeader = hasHeader,
            Header = header,
            Rows = rows,
            Warnings = warnings
        };
    }

    /// <summary>
    /// Reads every record of the file with the encoding and separator the
    /// analysis chose, applying the same skipping rules the analysis used.
    /// </summary>
    private List<CsvRecord> ReadAllRecords(string path, CsvAnalysis analysis)
    {
        using var stream = new FileStream(
            path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 64 * 1024, FileOptions.SequentialScan);
        using var streamReader = new StreamReader(stream, analysis.Encoding, detectEncodingFromByteOrderMarks: true);
        using var reader = new CsvRecordReader(
            streamReader, analysis.Separator, _settings.QuoteChar, _settings.TrimWhitespace, leaveOpen: true);

        var records = new List<CsvRecord>();

        while (reader.ReadRecord() is { } record)
        {
            // Honour the preamble the analyser skipped, so the table starts
            // where the data starts.
            if (record.StartLine < analysis.FirstContentLine)
                continue;

            if (_settings.SkipEmptyLines && record.IsBlank)
                continue;

            records.Add(record);
        }

        return records;
    }
}
