using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DataWizard.Core.Diff;
using DataWizard.Core.Excel;

namespace DataWizard.Core.Tests;

public class DiffReportWriterTests
{
    private static DiffReport Report(params DiffEntry[] entries) => new()
    {
        CreatedAt = new DateTime(2025, 6, 1, 12, 0, 0),
        ReferencePath = @"C:\data\reference.csv",
        CandidatePath = @"C:\data\candidate.xlsx",
        ReferenceSheet = null,
        CandidateSheet = "Data",
        Options = new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["id"],
            AbsoluteTolerance = 0.5d,
            IgnoreCase = true,
            IgnoredColumns = ["updatedAt"]
        },
        ReferenceRowCount = 10,
        CandidateRowCount = 12,
        ReferenceColumnCount = 4,
        CandidateColumnCount = 5,
        Entries = entries,
        ValueChangedCount = entries.Count(e => e.Kind == DiffKind.ValueChanged),
        RowOnlyInReferenceCount = entries.Count(e => e.Kind == DiffKind.RowOnlyInReference),
        RowOnlyInCandidateCount = entries.Count(e => e.Kind == DiffKind.RowOnlyInCandidate),
        ColumnOnlyInReferenceCount = entries.Count(e => e.Kind == DiffKind.ColumnOnlyInReference),
        ColumnOnlyInCandidateCount = entries.Count(e => e.Kind == DiffKind.ColumnOnlyInCandidate),
        DuplicateKeyCount = 1,
        Warnings = ["Key 'id: 3' occurs 2 times in the reference; rows are compared in order of appearance."]
    };

    [Fact]
    public void TheWorkbookHasASummarySheetAndADifferencesSheet()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.PathTo("report.xlsx");

        DiffReportWriter.Write(Report(
            new DiffEntry { Kind = DiffKind.ValueChanged, ReferenceRow = 2, CandidateRow = 5, Column = "price", ReferenceValue = "10", CandidateValue = "11" },
            new DiffEntry { Kind = DiffKind.RowOnlyInCandidate, CandidateRow = 7, Detail = "id: 9" }), path);

        using var document = SpreadsheetDocument.Open(path, false);
        var names = document.WorkbookPart!.Workbook!.Descendants<Sheet>()
            .Select(s => s.Name?.Value ?? "Sheet")
            .ToArray();

        Assert.Equal(["Summary", "Differences"], names);
    }

    [Fact]
    public void TheDifferencesSheetHasAHeaderAndOneRowPerEntry()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.PathTo("report.xlsx");

        DiffReportWriter.Write(Report(
            new DiffEntry { Kind = DiffKind.ValueChanged, ReferenceRow = 2, CandidateRow = 5, Column = "price", ReferenceValue = "10", CandidateValue = "11" },
            new DiffEntry { Kind = DiffKind.RowOnlyInCandidate, CandidateRow = 7, Detail = "id: 9" }), path);

        var grid = ExcelReader.ReadSheet(path, sheetIndex: 1);

        Assert.Equal("Differences", grid.SheetName);
        Assert.Equal(3, grid.Rows.Length);
        Assert.Equal(
            ["Kind", "Reference Row", "Candidate Row", "Column", "Reference Value", "Candidate Value", "Detail"],
            grid.Rows[0]);
        Assert.Equal("ValueChanged", grid.Rows[1][0]);
        Assert.Equal("2", grid.Rows[1][1]);
        Assert.Equal("5", grid.Rows[1][2]);
        Assert.Equal("price", grid.Rows[1][3]);
        Assert.Equal("10", grid.Rows[1][4]);
        Assert.Equal("11", grid.Rows[1][5]);
        Assert.Equal("RowOnlyInCandidate", grid.Rows[2][0]);
        Assert.Equal("0", grid.Rows[2][1]);
        Assert.Equal("7", grid.Rows[2][2]);
        Assert.Equal("id: 9", grid.Rows[2][6]);
    }

    [Fact]
    public void TheSummarySheetListsTheOptionsAndTheCounts()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.PathTo("report.xlsx");

        DiffReportWriter.Write(Report(
            new DiffEntry { Kind = DiffKind.ValueChanged, ReferenceRow = 1, CandidateRow = 1, Column = "a" }), path);

        var grid = ExcelReader.ReadSheet(path, sheetIndex: 0);
        var summary = grid.Rows.ToDictionary(r => r[0], r => r[1]);

        Assert.Equal("Keyed", summary["Match mode"]);
        Assert.Equal("id", summary["Key columns"]);
        Assert.Equal("0.5", summary["Absolute tolerance"]);
        Assert.Equal("Yes", summary["Ignore case"]);
        Assert.Equal("1", summary["Value changes"]);
        Assert.Equal("1", summary["Total differences"]);
        Assert.Equal("1", summary["Duplicate keys"]);
        Assert.Equal("C:\\data\\reference.csv - 10 rows, 4 columns", summary["Reference"]);
        Assert.Equal("C:\\data\\candidate.xlsx (Data) - 12 rows, 5 columns", summary["Candidate"]);
        Assert.Contains(grid.Rows, r => r[0] == "Warning" && r[1].Contains("occurs 2 times"));
    }

    [Fact]
    public void AnExistingReportIsReplaced()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.PathTo("report.xlsx");

        DiffReportWriter.Write(Report(), path);
        DiffReportWriter.Write(Report(
            new DiffEntry { Kind = DiffKind.ValueChanged, ReferenceRow = 1, CandidateRow = 1, Column = "a" }), path);

        var grid = ExcelReader.ReadSheet(path, sheetIndex: 1);

        Assert.Equal(2, grid.Rows.Length);
    }
}
