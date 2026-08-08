using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DataWizard.Core.Configuration;

namespace DataWizard.Core.Excel;

/// <summary>What kind of value a worksheet cell held.</summary>
public enum ExcelCellType
{
    Empty,
    Text,
    Number,
    Date,
    Boolean,
    Error
}

/// <summary>One CSV file produced from a worksheet.</summary>
/// <param name="Path">The file that was written.</param>
/// <param name="SheetName">The worksheet it came from.</param>
/// <param name="RowCount">How many rows were written.</param>
public readonly record struct ExcelSheetExport(string Path, string SheetName, int RowCount);

/// <summary>
/// Reads an XLSX workbook and writes its worksheets out as CSV.
/// </summary>
/// <remarks>
/// Worksheets are streamed row by row, and the shared string table is read once
/// into an array. Looking each string up by walking the table - as the previous
/// implementation did - turns the export into quadratic work and makes large
/// workbooks appear to hang.
/// </remarks>
public sealed class ExcelReader
{
    private readonly ConversionSettings _settings;

    public ExcelReader(ConversionSettings settings)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        EncodingDetectorBootstrap();
    }

    private static void EncodingDetectorBootstrap() => Csv.EncodingDetector.EnsureCodePagesRegistered();

    /// <summary>
    /// Converts a workbook to CSV. One file is written per worksheet when
    /// <see cref="ConversionSettings.ExportAllSheets"/> is set, otherwise only the
    /// selected worksheet is exported.
    /// </summary>
    /// <param name="xlsxPath">The workbook to read.</param>
    /// <param name="csvPath">
    /// Destination file. When exporting all sheets the worksheet name is inserted
    /// before the extension.
    /// </param>
    /// <param name="progress">Reports rows written for the current worksheet.</param>
    /// <param name="cancellationToken">Cancels the export.</param>
    public List<ExcelSheetExport> ToCsv(
        string xlsxPath,
        string csvPath,
        IProgress<int>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (!File.Exists(xlsxPath))
            throw new FileNotFoundException("Excel file not found.", xlsxPath);

        var results = new List<ExcelSheetExport>();

        using var document = SpreadsheetDocument.Open(xlsxPath, false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidDataException($"'{xlsxPath}' has no workbook part and is not a valid XLSX file.");

        var sheets = workbookPart.Workbook?.Descendants<Sheet>().ToList() ?? [];

        if (sheets.Count == 0)
            throw new InvalidDataException($"'{xlsxPath}' contains no worksheets.");

        var sharedStrings = ReadSharedStrings(workbookPart);
        var styles = ReadStyleFormats(workbookPart);

        if (_settings.ExportAllSheets)
        {
            foreach (var sheet in sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var sheetName = sheet.Name?.Value ?? "Sheet";
                var outputPath = InsertSheetName(csvPath, sheetName);
                var rows = ExportSheet(workbookPart, sheet, outputPath, sharedStrings, styles, progress, cancellationToken);

                results.Add(new ExcelSheetExport(outputPath, sheetName, rows));
            }
        }
        else
        {
            var sheet = sheets.ElementAtOrDefault(_settings.WorksheetIndex)
                ?? throw new ArgumentOutOfRangeException(
                    nameof(csvPath),
                    $"Worksheet index {_settings.WorksheetIndex} is out of range; the workbook has {sheets.Count}.");

            var sheetName = sheet.Name?.Value ?? "Sheet";
            var rows = ExportSheet(workbookPart, sheet, csvPath, sharedStrings, styles, progress, cancellationToken);

            results.Add(new ExcelSheetExport(csvPath, sheetName, rows));
        }

        return results;
    }

    /// <summary>Lists the worksheet names of a workbook, for the UI selector.</summary>
    public static IReadOnlyList<string> ListSheetNames(string xlsxPath)
    {
        using var document = SpreadsheetDocument.Open(xlsxPath, false);
        var workbookPart = document.WorkbookPart;

        if (workbookPart?.Workbook is null)
            return [];

        return workbookPart.Workbook.Descendants<Sheet>()
            .Select(s => s.Name?.Value ?? "Sheet")
            .ToArray();
    }

    private int ExportSheet(
        WorkbookPart workbookPart,
        Sheet sheet,
        string csvPath,
        string[] sharedStrings,
        StyleFormat[] styles,
        IProgress<int>? progress,
        CancellationToken cancellationToken)
    {
        if (sheet.Id?.Value is null)
            return 0;

        if (File.Exists(csvPath) && !_settings.Overwrite)
            throw new IOException($"Output file already exists: {csvPath}");

        var directory = Path.GetDirectoryName(csvPath);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var separator = _settings.ResolveCsvSeparator();
        var culture = _settings.ResolveNumberOutputCulture();

        using var stream = new FileStream(csvPath, FileMode.Create, FileAccess.Write, FileShare.None, 64 * 1024);
        using var writer = new StreamWriter(stream, _settings.ResolveCsvEncoding())
        {
            NewLine = _settings.ResolveLineEnding()
        };

        var rowsWritten = 0;
        var expectedRowIndex = 1u;

        using var reader = OpenXmlReader.Create(worksheetPart);

        while (reader.Read())
        {
            if (reader.ElementType != typeof(Row) || !reader.IsStartElement)
                continue;

            cancellationToken.ThrowIfCancellationRequested();

            var row = (Row)reader.LoadCurrentElement()!;
            var rowIndex = row.RowIndex?.Value ?? expectedRowIndex;

            // Worksheets omit rows that hold nothing. Emitting blank lines for
            // them keeps the CSV aligned with what the sheet looks like.
            if (_settings.KeepEmptyRows)
            {
                while (expectedRowIndex < rowIndex)
                {
                    writer.WriteLine();
                    rowsWritten++;
                    expectedRowIndex++;
                }
            }

            expectedRowIndex = rowIndex + 1;

            var line = BuildLine(row, sharedStrings, styles, separator, culture);
            writer.WriteLine(line);
            rowsWritten++;

            if (rowsWritten % 1000 == 0)
                progress?.Report(rowsWritten);
        }

        progress?.Report(rowsWritten);
        return rowsWritten;
    }

    private string BuildLine(
        Row row,
        string[] sharedStrings,
        StyleFormat[] styles,
        char separator,
        CultureInfo culture)
    {
        var cells = row.Elements<Cell>().ToList();

        if (cells.Count == 0)
            return string.Empty;

        // Index the row by column so gaps are filled rather than shifting the
        // remaining values left.
        var byColumn = new Dictionary<int, Cell>(cells.Count);
        var maxColumn = -1;

        for (var i = 0; i < cells.Count; i++)
        {
            var cell = cells[i];
            var column = cell.CellReference?.Value is { } reference
                ? ColumnIndex(reference)
                : i;

            byColumn[column] = cell;
            maxColumn = Math.Max(maxColumn, column);
        }

        var builder = new StringBuilder();

        for (var column = 0; column <= maxColumn; column++)
        {
            if (column > 0)
                builder.Append(separator);

            var value = string.Empty;
            var type = ExcelCellType.Empty;

            if (byColumn.TryGetValue(column, out var cell))
                value = ReadCellValue(cell, sharedStrings, styles, culture, out type);

            builder.Append(FormatField(value, separator, type));
        }

        return builder.ToString();
    }

    private string ReadCellValue(
        Cell cell,
        string[] sharedStrings,
        StyleFormat[] styles,
        CultureInfo culture,
        out ExcelCellType type)
    {
        var dataType = cell.DataType?.Value;

        if (dataType == CellValues.InlineString)
        {
            type = ExcelCellType.Text;
            return cell.InlineString?.Text?.Text ?? string.Empty;
        }

        if (dataType == CellValues.SharedString)
        {
            type = ExcelCellType.Text;
            var raw = cell.CellValue?.InnerText;

            if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                && index >= 0 && index < sharedStrings.Length)
            {
                return sharedStrings[index];
            }

            return raw ?? string.Empty;
        }

        var text = cell.CellValue?.InnerText;

        if (string.IsNullOrEmpty(text))
        {
            type = ExcelCellType.Empty;
            return string.Empty;
        }

        if (dataType == CellValues.Boolean)
        {
            type = ExcelCellType.Boolean;
            return text == "1" ? "TRUE" : "FALSE";
        }

        if (dataType == CellValues.Error)
        {
            type = ExcelCellType.Error;
            return text;
        }

        if (dataType == CellValues.String)
        {
            // Cached formula result, and what DataWizard itself used to write.
            type = ExcelCellType.Text;
            return text;
        }

        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            type = ExcelCellType.Text;
            return text;
        }

        if (IsDateStyled(cell, styles))
        {
            try
            {
                type = ExcelCellType.Date;
                return DateTime.FromOADate(number).ToString(_settings.DateOutputFormat, culture);
            }
            catch (ArgumentException)
            {
                // Out of the range Excel dates can express - keep it as a number.
            }
        }

        type = ExcelCellType.Number;
        return FormatNumber(number, culture);
    }

    /// <summary>
    /// Formats a number for the CSV output.
    /// </summary>
    /// <remarks>
    /// The default keeps up to fifteen decimal places, which is everything a
    /// double can carry. The previous implementation formatted with <c>0.##</c>,
    /// so every value with more than two decimals was silently rounded on the way
    /// out and a round trip lost data.
    /// </remarks>
    private string FormatNumber(double value, CultureInfo culture)
    {
        if (value == Math.Truncate(value) && Math.Abs(value) < 1e15)
            return value.ToString("0", culture);

        var decimals = Math.Clamp(_settings.MaxDecimalPlaces, 0, 15);

        if (decimals == 0)
            return Math.Round(value, MidpointRounding.AwayFromZero).ToString("0", culture);

        return value.ToString("0." + new string('#', decimals), culture);
    }

    /// <summary>
    /// Wraps a value in quotes when the mode or the content requires it.
    /// </summary>
    private string FormatField(string value, char separator, ExcelCellType type)
    {
        var mustQuote =
            value.Contains(separator) ||
            value.Contains('"') ||
            value.Contains('\n') ||
            value.Contains('\r') ||
            (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])));

        var quote = _settings.QuoteMode switch
        {
            QuoteMode.All => true,
            QuoteMode.TextFields => mustQuote || type == ExcelCellType.Text,
            _ => mustQuote
        };

        if (!quote)
            return value;

        return '"' + value.Replace("\"", "\"\"") + '"';
    }

    /// <summary>
    /// Reads the shared string table into an array in one pass.
    /// </summary>
    private static string[] ReadSharedStrings(WorkbookPart workbookPart)
    {
        var part = workbookPart.SharedStringTablePart;
        if (part?.SharedStringTable is null)
            return [];

        return part.SharedStringTable
            .Elements<SharedStringItem>()
            .Select(item => item.Text?.Text ?? item.InnerText)
            .ToArray();
    }

    /// <summary>The number format behind each cell format index.</summary>
    private readonly record struct StyleFormat(uint NumberFormatId, string? FormatCode);

    private static StyleFormat[] ReadStyleFormats(WorkbookPart workbookPart)
    {
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        if (stylesheet?.CellFormats is null)
            return [];

        var customCodes = new Dictionary<uint, string>();

        if (stylesheet.NumberingFormats is not null)
        {
            foreach (var format in stylesheet.NumberingFormats.Elements<NumberingFormat>())
            {
                if (format.NumberFormatId?.Value is { } id && format.FormatCode?.Value is { } code)
                    customCodes[id] = code;
            }
        }

        return stylesheet.CellFormats
            .Elements<CellFormat>()
            .Select(format =>
            {
                var id = format.NumberFormatId?.Value ?? 0u;
                return new StyleFormat(id, customCodes.GetValueOrDefault(id));
            })
            .ToArray();
    }

    private static bool IsDateStyled(Cell cell, StyleFormat[] styles)
    {
        if (cell.StyleIndex?.Value is not { } index || index >= styles.Length)
            return false;

        var style = styles[index];
        return ExcelStyles.IsDateFormatId(style.NumberFormatId)
               || ExcelStyles.IsDateFormatCode(style.FormatCode);
    }

    /// <summary>Converts a cell reference such as <c>BC12</c> to a zero-based column index.</summary>
    public static int ColumnIndex(string cellReference)
    {
        var index = 0;
        var any = false;

        foreach (var c in cellReference)
        {
            if (c is >= 'A' and <= 'Z')
            {
                index = index * 26 + (c - 'A' + 1);
                any = true;
            }
            else if (c is >= 'a' and <= 'z')
            {
                index = index * 26 + (c - 'a' + 1);
                any = true;
            }
            else if (any)
            {
                break;
            }
        }

        return any ? index - 1 : 0;
    }

    /// <summary>
    /// Inserts a worksheet name before the extension, for example
    /// <c>data.csv</c> and <c>Customers</c> become <c>data_Customers.csv</c>.
    /// </summary>
    public static string InsertSheetName(string csvPath, string sheetName)
    {
        var directory = Path.GetDirectoryName(csvPath);
        var stem = Path.GetFileNameWithoutExtension(csvPath);
        var extension = Path.GetExtension(csvPath);
        var fileName = $"{stem}_{SanitizeFileName(sheetName)}{extension}";

        return string.IsNullOrEmpty(directory) ? fileName : Path.Combine(directory, fileName);
    }

    /// <summary>Replaces characters a file name cannot contain.</summary>
    public static string SanitizeFileName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "Sheet";

        var invalid = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder(name.Length);

        foreach (var c in name)
            builder.Append(invalid.Contains(c) || c == ' ' ? '_' : c);

        var sanitized = builder.ToString().Trim('.', '_');
        return sanitized.Length == 0 ? "Sheet" : sanitized;
    }
}
