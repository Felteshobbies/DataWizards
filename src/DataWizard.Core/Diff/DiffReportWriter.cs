using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DataWizard.Core.Configuration;
using DataWizard.Core.Excel;

namespace DataWizard.Core.Diff;

/// <summary>
/// Writes a <see cref="DiffReport"/> to an XLSX workbook with two worksheets:
/// a summary of what was compared and how much differed, and one row per
/// difference.
/// </summary>
public static class DiffReportWriter
{
    private const string SummarySheet = "Summary";
    private const string DifferencesSheet = "Differences";

    /// <summary>
    /// Writes the report to <paramref name="xlsxPath"/>, replacing an existing
    /// file.
    /// </summary>
    public static void Write(DiffReport report, string xlsxPath)
    {
        ArgumentNullException.ThrowIfNull(report);
        ArgumentException.ThrowIfNullOrEmpty(xlsxPath);

        if (File.Exists(xlsxPath))
            File.Delete(xlsxPath);

        var directory = Path.GetDirectoryName(xlsxPath);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        using var document = SpreadsheetDocument.Create(xlsxPath, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = ExcelStyles.Create(new ConversionSettings());
        stylesPart.Stylesheet.Save();

        var summaryPart = workbookPart.AddNewPart<WorksheetPart>();
        WriteSummary(summaryPart, report);

        var differencesPart = workbookPart.AddNewPart<WorksheetPart>();
        WriteDifferences(differencesPart, report);

        workbookPart.Workbook = new Workbook(
            new Sheets(
                new Sheet
                {
                    Id = workbookPart.GetIdOfPart(summaryPart),
                    SheetId = 1U,
                    Name = SummarySheet
                },
                new Sheet
                {
                    Id = workbookPart.GetIdOfPart(differencesPart),
                    SheetId = 2U,
                    Name = DifferencesSheet
                }));

        workbookPart.Workbook.Save();
    }

    private static void WriteSummary(WorksheetPart part, DiffReport report)
    {
        using var writer = OpenXmlWriter.Create(part);

        writer.WriteStartElement(new Worksheet());
        writer.WriteElement(new Columns(
            new Column { Min = 1U, Max = 1U, Width = 26d, CustomWidth = true },
            new Column { Min = 2U, Max = 2U, Width = 90d, CustomWidth = true }));

        writer.WriteStartElement(new SheetData());

        var rows = new (string Label, string Value)[]
        {
            ("Reference", DescribeTable(report.ReferencePath, report.ReferenceSheet, report.ReferenceRowCount, report.ReferenceColumnCount)),
            ("Candidate", DescribeTable(report.CandidatePath, report.CandidateSheet, report.CandidateRowCount, report.CandidateColumnCount)),
            ("Match mode", report.Options.MatchMode.ToString()),
            ("Key columns", Join(report.Options.KeyColumns)),
            ("Absolute tolerance", report.Options.AbsoluteTolerance.ToString("R", CultureInfo.InvariantCulture)),
            ("Relative tolerance", report.Options.RelativeTolerance.ToString("R", CultureInfo.InvariantCulture)),
            ("Date tolerance", report.Options.DateTolerance.ToString()),
            ("Ignore case", report.Options.IgnoreCase ? "Yes" : "No"),
            ("Trim values", report.Options.TrimValues ? "Yes" : "No"),
            ("Ignored columns", Join(report.Options.IgnoredColumns)),
            ("Compared at", report.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)),
            ("Value changes", report.ValueChangedCount.ToString(CultureInfo.InvariantCulture)),
            ("Rows only in reference", report.RowOnlyInReferenceCount.ToString(CultureInfo.InvariantCulture)),
            ("Rows only in candidate", report.RowOnlyInCandidateCount.ToString(CultureInfo.InvariantCulture)),
            ("Columns only in reference", report.ColumnOnlyInReferenceCount.ToString(CultureInfo.InvariantCulture)),
            ("Columns only in candidate", report.ColumnOnlyInCandidateCount.ToString(CultureInfo.InvariantCulture)),
            ("Duplicate keys", report.DuplicateKeyCount.ToString(CultureInfo.InvariantCulture)),
            ("Total differences", report.DifferenceCount.ToString(CultureInfo.InvariantCulture))
        };

        for (var i = 0; i < rows.Length; i++)
        {
            var (label, value) = rows[i];
            writer.WriteElement(CreateRow(new[] { label, value }, (uint)(i + 1), isHeader: false, boldFirst: true));
        }

        for (var i = 0; i < report.Warnings.Count; i++)
        {
            writer.WriteElement(CreateRow(new[] { "Warning", report.Warnings[i] }, (uint)(rows.Length + i + 1), isHeader: false, boldFirst: false));
        }

        writer.WriteEndElement(); // sheetData
        writer.WriteEndElement(); // worksheet
    }

    private static void WriteDifferences(WorksheetPart part, DiffReport report)
    {
        using var writer = OpenXmlWriter.Create(part);

        writer.WriteStartElement(new Worksheet());
        writer.WriteElement(CreateFrozenHeaderView());
        writer.WriteElement(new Columns(
            new Column { Min = 1U, Max = 1U, Width = 24d, CustomWidth = true },
            new Column { Min = 2U, Max = 2U, Width = 14d, CustomWidth = true },
            new Column { Min = 3U, Max = 3U, Width = 14d, CustomWidth = true },
            new Column { Min = 4U, Max = 4U, Width = 28d, CustomWidth = true },
            new Column { Min = 5U, Max = 5U, Width = 40d, CustomWidth = true },
            new Column { Min = 6U, Max = 6U, Width = 40d, CustomWidth = true },
            new Column { Min = 7U, Max = 7U, Width = 40d, CustomWidth = true }));

        writer.WriteStartElement(new SheetData());

        writer.WriteElement(CreateRow(
            ["Kind", "Reference Row", "Candidate Row", "Column", "Reference Value", "Candidate Value", "Detail"],
            1U, isHeader: true, boldFirst: false));

        for (var i = 0; i < report.Entries.Count; i++)
        {
            var entry = report.Entries[i];

            writer.WriteElement(CreateRow(
                [
                    entry.Kind.ToString(),
                    entry.ReferenceRow.ToString(CultureInfo.InvariantCulture),
                    entry.CandidateRow.ToString(CultureInfo.InvariantCulture),
                    entry.Column,
                    entry.ReferenceValue,
                    entry.CandidateValue,
                    entry.Detail
                ],
                (uint)(i + 2), isHeader: false, boldFirst: false));
        }

        writer.WriteEndElement(); // sheetData
        writer.WriteEndElement(); // worksheet
    }

    private static Row CreateRow(string[] values, uint rowIndex, bool isHeader, bool boldFirst)
    {
        var row = new Row { RowIndex = rowIndex };

        for (var i = 0; i < values.Length; i++)
        {
            var reference = ExcelWriter.ColumnName(i) + rowIndex.ToString(CultureInfo.InvariantCulture);

            if (i == 0 && (isHeader || boldFirst))
            {
                row.Append(TextCell(values[i], reference, ExcelStyles.Header));
                continue;
            }

            if (isHeader)
            {
                row.Append(TextCell(values[i], reference, ExcelStyles.Header));
                continue;
            }

            row.Append(TextCell(values[i], reference, ExcelStyles.Text));
        }

        return row;
    }

    private static Cell TextCell(string value, string reference, uint styleIndex) => new()
    {
        CellReference = reference,
        DataType = CellValues.InlineString,
        StyleIndex = styleIndex,
        InlineString = new InlineString(new Text(ExcelWriter.SanitizeXmlText(value))
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

    private static string DescribeTable(string path, string? sheetName, int rowCount, int columnCount)
    {
        var location = string.IsNullOrEmpty(sheetName) ? path : $"{path} ({sheetName})";
        return $"{location} - {rowCount} rows, {columnCount} columns";
    }

    private static string Join(string[] values) =>
        values.Length == 0 ? string.Empty : string.Join(", ", values);
}
