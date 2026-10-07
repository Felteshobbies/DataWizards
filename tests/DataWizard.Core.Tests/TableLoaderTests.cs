using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;
using DataWizard.Core.Diff;

namespace DataWizard.Core.Tests;

public class TableLoaderTests
{
    private readonly TableLoader _loader = new(new DetectionSettings());

    private static string MakeXlsx(TempWorkspace workspace, string name, string[] lines, Action<DataWizardSettings>? configure = null)
    {
        var csv = workspace.WriteLines($"{name}.csv", lines);

        var settings = DataWizardSettings.CreateDefault();
        settings.Conversion.Overwrite = true;
        settings.Conversion.WriteByteOrderMark = false;
        settings.Conversion.NumberOutputCulture = string.Empty;
        configure?.Invoke(settings);
        settings.Normalize();

        var result = new ConversionService(settings).Convert(csv);
        Assert.True(result.Success, result.ErrorMessage);
        return result.OutputPaths[0];
    }

    [Fact]
    public void ACsvWithAHeaderLoadsTheHeaderAndTheDataRows()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("input.csv",
            "id;name;price",
            "1;Widget;19.99",
            "2;Gadget;4.50");

        var table = _loader.Load(path);

        Assert.True(table.HasHeader);
        Assert.Equal(["id", "name", "price"], table.Header);
        Assert.Equal(2, table.RowCount);
        Assert.Equal("19.99", table.Rows[0][2]);
        Assert.Equal("Gadget", table.Rows[1][1]);
        Assert.Null(table.SheetName);
    }

    [Fact]
    public void AHeaderlessCsvGetsGeneratedColumnNames()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("input.csv",
            "1;2",
            "3;4");

        var table = _loader.Load(path);

        Assert.False(table.HasHeader);
        Assert.Equal(["Column 1", "Column 2"], table.Header);
        Assert.Equal(2, table.RowCount);
    }

    [Fact]
    public void RaggedRowsArePaddedWithEmptyStrings()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("input.csv",
            "id;name;note",
            "1;Widget",
            "2;Gadget;extra");

        var table = _loader.Load(path);

        Assert.Equal(3, table.ColumnCount);
        Assert.Equal(string.Empty, table.Rows[0][2]);
        Assert.Equal("extra", table.Rows[1][2]);
    }

    [Fact]
    public void AWorkbookLoadsNumbersInInvariantCulture()
    {
        using var workspace = new TempWorkspace();
        var xlsx = MakeXlsx(workspace, "numbers",
            ["id;price;quantity", "1;19.99;5"]);

        var table = _loader.Load(xlsx);

        Assert.True(table.HasHeader);
        Assert.Equal("19.99", table.Rows[0][1]);
        Assert.Equal("5", table.Rows[0][2]);
    }

    [Fact]
    public void AWorkbookLoadsDateStyledCellsAsIsoTimestamps()
    {
        using var workspace = new TempWorkspace();
        var xlsx = MakeXlsx(workspace, "dates",
            ["label;date;amount", "start;31.01.2025;100", "end;01.02.2025;200"]);

        var table = _loader.Load(xlsx);

        Assert.True(table.HasHeader);
        Assert.Equal("2025-01-31T00:00:00.0000000", table.Rows[0][1]);
        Assert.Equal("2025-02-01T00:00:00.0000000", table.Rows[1][1]);
    }

    [Fact]
    public void AWorkbookLoadsBooleanCellsAsTrueAndFalse()
    {
        using var workspace = new TempWorkspace();
        var xlsx = MakeXlsx(workspace, "flags",
            ["name;flag", "on;true", "off;false"],
            s => s.Detection.DetectBooleans = true);

        // The loader must be told to read booleans, just like the converter was.
        var loader = new TableLoader(new DetectionSettings { DetectBooleans = true });
        var table = loader.Load(xlsx);

        Assert.True(table.HasHeader);
        Assert.Equal("TRUE", table.Rows[0][1]);
        Assert.Equal("FALSE", table.Rows[1][1]);
    }

    [Fact]
    public void AWorkbookLoadsTheSheetName()
    {
        using var workspace = new TempWorkspace();
        var xlsx = MakeXlsx(workspace, "named",
            ["id", "1"],
            s => s.Conversion.SheetName = "Products");

        var table = _loader.Load(xlsx);

        Assert.Equal("Products", table.SheetName);
    }

    [Fact]
    public void ASecondWorksheetCanBeSelectedByIndex()
    {
        using var workspace = new TempWorkspace();
        var xlsx = workspace.PathTo("two.xlsx");

        using (var document = SpreadsheetDocument.Create(xlsx, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();

            var first = workbookPart.AddNewPart<WorksheetPart>();
            first.Worksheet = new Worksheet(new SheetData(new Row(
                new Cell
                {
                    CellReference = "A1",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text("1"))
                })));
            first.Worksheet.Save();

            // Two rows: a lone value would score as a header and be skipped.
            var second = workbookPart.AddNewPart<WorksheetPart>();
            second.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell
                {
                    CellReference = "A1",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text("2"))
                }),
                new Row(new Cell
                {
                    CellReference = "A2",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text("3"))
                })));
            second.Worksheet.Save();

            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = workbookPart.GetIdOfPart(first), SheetId = 1U, Name = "First" },
                new Sheet { Id = workbookPart.GetIdOfPart(second), SheetId = 2U, Name = "Second" }));
            workbookPart.Workbook.Save();
        }

        var table = _loader.Load(xlsx, sheetIndex: 1);

        Assert.Equal("Second", table.SheetName);
        Assert.Equal("2", table.Rows[0][0]);
    }

    [Fact]
    public void AMissingFileThrows()
    {
        using var workspace = new TempWorkspace();

        Assert.Throws<FileNotFoundException>(() => _loader.Load(workspace.PathTo("nope.csv")));
    }

    [Fact]
    public void AWorksheetIndexOutOfRangeThrows()
    {
        using var workspace = new TempWorkspace();
        var xlsx = MakeXlsx(workspace, "one", ["id", "1"]);

        Assert.Throws<ArgumentOutOfRangeException>(() => _loader.Load(xlsx, sheetIndex: 5));
    }
}
