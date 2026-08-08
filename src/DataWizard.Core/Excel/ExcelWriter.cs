using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;

namespace DataWizard.Core.Excel;

/// <summary>The outcome of writing a workbook.</summary>
/// <param name="Path">The file that was written.</param>
/// <param name="RowCount">Rows written, including the header.</param>
/// <param name="ColumnCount">Widest row written.</param>
public readonly record struct ExcelWriteResult(string Path, int RowCount, int ColumnCount);

/// <summary>
/// Writes a CSV file into an XLSX workbook, applying the column types worked out
/// by analysis.
/// </summary>
/// <remarks>
/// Rows are streamed straight to the package with <see cref="OpenXmlWriter"/>
/// rather than assembled in memory first, so file size is limited by disk rather
/// than by RAM.
/// </remarks>
public sealed class ExcelWriter
{
    private readonly ConversionSettings _conversion;
    private readonly DetectionSettings _detection;
    private readonly ValueTypeDetector _typeDetector;

    public ExcelWriter(ConversionSettings conversion, DetectionSettings detection)
    {
        _conversion = conversion ?? throw new ArgumentNullException(nameof(conversion));
        _detection = detection ?? throw new ArgumentNullException(nameof(detection));
        _typeDetector = new ValueTypeDetector(detection);
    }

    /// <summary>
    /// Converts the file described by <paramref name="analysis"/> into
    /// <paramref name="xlsxPath"/>.
    /// </summary>
    /// <param name="analysis">Analysis of the source CSV.</param>
    /// <param name="xlsxPath">Destination workbook.</param>
    /// <param name="progress">Reports the number of rows written so far.</param>
    /// <param name="cancellationToken">Cancels the conversion.</param>
    public ExcelWriteResult Write(
        CsvAnalysis analysis,
        string xlsxPath,
        IProgress<int>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(analysis);

        if (File.Exists(xlsxPath))
        {
            if (!_conversion.Overwrite)
                throw new IOException($"Output file already exists: {xlsxPath}");

            File.Delete(xlsxPath);
        }

        var directory = Path.GetDirectoryName(xlsxPath);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        var columnTypes = analysis.Columns.Select(c => c.EffectiveType).ToArray();
        var columnRules = analysis.Columns.Select(c => c.AppliedRule).ToArray();
        var widths = EstimateColumnWidths(analysis);

        var rowCount = 0;
        var columnCount = 0;

        using (var document = SpreadsheetDocument.Create(xlsxPath, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = ExcelStyles.Create(_conversion);
            stylesPart.Stylesheet.Save();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            using (var writer = OpenXmlWriter.Create(worksheetPart))
            {
                writer.WriteStartElement(new Worksheet());

                if (_conversion.FreezeHeaderRow && analysis.HasHeader)
                    writer.WriteElement(CreateFrozenHeaderView());

                if (_conversion.AutoFitColumns && widths.Length > 0)
                    writer.WriteElement(CreateColumns(widths));

                writer.WriteStartElement(new SheetData());

                foreach (var record in ReadSourceRecords(analysis, cancellationToken))
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    rowCount++;
                    columnCount = Math.Max(columnCount, record.FieldCount);

                    var isHeaderRow = rowCount == 1 && analysis.HasHeader;
                    writer.WriteElement(CreateRow(record, (uint)rowCount, isHeaderRow, columnTypes, columnRules));

                    if (rowCount % 1000 == 0)
                        progress?.Report(rowCount);
                }

                writer.WriteEndElement(); // sheetData

                if (_conversion.AutoFilterHeaderRow && analysis.HasHeader && rowCount > 1 && columnCount > 0)
                {
                    writer.WriteElement(new AutoFilter
                    {
                        Reference = $"A1:{ColumnName(columnCount - 1)}{rowCount}"
                    });
                }

                writer.WriteEndElement(); // worksheet
            }

            workbookPart.Workbook = new Workbook(
                new Sheets(
                    new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1U,
                        Name = SanitizeSheetName(_conversion.SheetName)
                    }));

            workbookPart.Workbook.Save();
        }

        progress?.Report(rowCount);
        return new ExcelWriteResult(xlsxPath, rowCount, columnCount);
    }

    private IEnumerable<CsvRecord> ReadSourceRecords(CsvAnalysis analysis, CancellationToken cancellationToken)
    {
        using var stream = new FileStream(
            analysis.FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 64 * 1024, FileOptions.SequentialScan);
        using var streamReader = new StreamReader(stream, analysis.Encoding, detectEncodingFromByteOrderMarks: true);
        using var reader = new CsvRecordReader(
            streamReader, analysis.Separator, _detection.QuoteChar, _detection.TrimWhitespace, leaveOpen: true);

        var line = 0;

        while (reader.ReadRecord() is { } record)
        {
            cancellationToken.ThrowIfCancellationRequested();
            line++;

            // Honour the preamble the analyser skipped, so the workbook starts
            // where the data starts.
            if (record.StartLine < analysis.FirstContentLine)
                continue;

            if (_detection.SkipEmptyLines && record.IsBlank)
                continue;

            yield return record;
        }
    }

    private Row CreateRow(
        CsvRecord record,
        uint rowIndex,
        bool isHeaderRow,
        FieldDataType[] columnTypes,
        FieldRule?[] columnRules)
    {
        var row = new Row { RowIndex = rowIndex };

        for (var i = 0; i < record.FieldCount; i++)
        {
            var reference = ColumnName(i) + rowIndex.ToString(CultureInfo.InvariantCulture);
            var field = record.Fields[i];

            if (isHeaderRow)
            {
                row.Append(TextCell(field.Value, reference, _conversion.BoldHeaderRow
                    ? ExcelStyles.Header
                    : ExcelStyles.Text));
                continue;
            }

            var declaredType = i < columnTypes.Length ? columnTypes[i] : FieldDataType.Auto;
            var rule = i < columnRules.Length ? columnRules[i] : null;

            row.Append(CreateCell(field, reference, declaredType, rule));
        }

        return row;
    }

    /// <summary>
    /// Builds one cell. A value that was quoted in the source stays text even when
    /// a rule asks for something else - the quotes are the author's explicit
    /// statement that the characters matter.
    /// </summary>
    private Cell CreateCell(CsvField field, string reference, FieldDataType declaredType, FieldRule? rule)
    {
        if (_detection.QuotedFieldsAreText && field.WasQuoted)
            return TextCell(field.Value, reference, ExcelStyles.Text);

        if (field.Value.Length == 0)
            return new Cell { CellReference = reference, StyleIndex = ExcelStyles.Text };

        var type = declaredType == FieldDataType.Auto
            ? _typeDetector.Detect(field).Type
            : declaredType;

        switch (type)
        {
            case FieldDataType.Integer:
            case FieldDataType.Decimal:
                if (_typeDetector.TryParseNumber(field.Value, out var number, out var isInteger))
                {
                    return new Cell
                    {
                        CellReference = reference,
                        DataType = CellValues.Number,
                        CellValue = new CellValue(number.ToString("R", CultureInfo.InvariantCulture)),
                        StyleIndex = type == FieldDataType.Integer && isInteger
                            ? ExcelStyles.Integer
                            : ExcelStyles.Decimal
                    };
                }
                break;

            case FieldDataType.Date:
                if (_typeDetector.TryParseDate(field.Value, rule?.DateFormat, out var date))
                {
                    return new Cell
                    {
                        CellReference = reference,
                        DataType = CellValues.Number,
                        CellValue = new CellValue(date.ToOADate().ToString("R", CultureInfo.InvariantCulture)),
                        StyleIndex = ExcelStyles.Date
                    };
                }
                break;

            case FieldDataType.Boolean:
                if (_typeDetector.TryParseBoolean(field.Value, out var boolean))
                {
                    return new Cell
                    {
                        CellReference = reference,
                        DataType = CellValues.Boolean,
                        CellValue = new CellValue(boolean ? "1" : "0"),
                        StyleIndex = ExcelStyles.Text
                    };
                }
                break;
        }

        // A value that does not fit its declared type is written verbatim rather
        // than discarded, so nothing is silently lost.
        return TextCell(field.Value, reference, ExcelStyles.Text);
    }

    /// <summary>
    /// Writes a text cell as an inline string.
    /// </summary>
    /// <remarks>
    /// Inline strings are the correct representation for literal text. The
    /// previous implementation used <c>t="str"</c>, which the format reserves for
    /// cached formula results; Excel tolerates it but stricter readers do not.
    /// </remarks>
    private static Cell TextCell(string value, string reference, uint styleIndex) => new()
    {
        CellReference = reference,
        DataType = CellValues.InlineString,
        StyleIndex = styleIndex,
        InlineString = new InlineString(new Text(SanitizeXmlText(value))
        {
            Space = SpaceProcessingModeValues.Preserve
        })
    };

    private static SheetViews CreateFrozenHeaderView() => new(
        new SheetView
        {
            WorkbookViewId = 0U,
            Pane = new Pane
            {
                VerticalSplit = 1D,
                TopLeftCell = "A2",
                ActivePane = PaneValues.BottomLeft,
                State = PaneStateValues.Frozen
            }
        });

    private static Columns CreateColumns(double[] widths)
    {
        var columns = new Columns();

        for (var i = 0; i < widths.Length; i++)
        {
            columns.Append(new Column
            {
                Min = (uint)(i + 1),
                Max = (uint)(i + 1),
                Width = widths[i],
                CustomWidth = true
            });
        }

        return columns;
    }

    /// <summary>
    /// Estimates column widths from the analysis sample. Measuring the whole file
    /// would mean reading it twice; the sample is representative enough for a
    /// sensible default width.
    /// </summary>
    private static double[] EstimateColumnWidths(CsvAnalysis analysis)
    {
        var count = analysis.Columns.Count;
        if (count == 0)
            return [];

        var widths = new double[count];

        for (var i = 0; i < count; i++)
            widths[i] = analysis.Columns[i].Name.Length;

        foreach (var record in analysis.SampleRecords)
        {
            for (var i = 0; i < record.FieldCount && i < count; i++)
                widths[i] = Math.Max(widths[i], record.Fields[i].Value.Length);
        }

        for (var i = 0; i < count; i++)
            widths[i] = Math.Clamp(widths[i] + 2, 8d, 60d);

        return widths;
    }

    /// <summary>
    /// Converts a zero-based index to an Excel column name (A, B, ... Z, AA, AB).
    /// </summary>
    public static string ColumnName(int columnIndex)
    {
        if (columnIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(columnIndex));

        var name = string.Empty;
        var index = columnIndex;

        while (index >= 0)
        {
            name = (char)('A' + index % 26) + name;
            index = index / 26 - 1;
        }

        return name;
    }

    /// <summary>
    /// Trims a worksheet name to what Excel accepts: at most 31 characters and
    /// none of <c>: \ / ? * [ ]</c>.
    /// </summary>
    public static string SanitizeSheetName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "Sheet1";

        var builder = new StringBuilder(name.Length);

        foreach (var c in name)
            builder.Append(c is ':' or '\\' or '/' or '?' or '*' or '[' or ']' ? '_' : c);

        var sanitized = builder.ToString().Trim('\'');
        if (sanitized.Length > 31)
            sanitized = sanitized[..31];

        return sanitized.Length == 0 ? "Sheet1" : sanitized;
    }

    /// <summary>
    /// Removes characters that XML 1.0 cannot represent. Control bytes do turn up
    /// in exports, and left in place they produce a workbook Excel refuses to open.
    /// </summary>
    public static string SanitizeXmlText(string value)
    {
        var needsWork = false;

        foreach (var c in value)
        {
            if (IsInvalidXmlChar(c))
            {
                needsWork = true;
                break;
            }
        }

        if (!needsWork)
            return value;

        var builder = new StringBuilder(value.Length);

        foreach (var c in value)
        {
            if (!IsInvalidXmlChar(c))
                builder.Append(c);
        }

        return builder.ToString();
    }

    private static bool IsInvalidXmlChar(char c) =>
        c switch
        {
            '\t' or '\n' or '\r' => false,
            < ' ' => true,
            >= '\uD800' and <= '\uDFFF' => false, // surrogate halves are handled as pairs
            '￾' or '￿' => true,
            _ => false
        };
}
