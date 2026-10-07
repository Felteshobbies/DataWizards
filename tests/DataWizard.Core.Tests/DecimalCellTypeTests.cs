using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataWizard.Core.Tests;

/// <summary>
/// Regression tests for the "decimals end up partially as text" bug. Values
/// that contain both a dot and a comma (for example <c>1.234,56</c>) used to
/// fail numeric parsing and be written as text cells, even inside a column
/// that was otherwise all numbers.
/// </summary>
public class DecimalCellTypeTests
{
    private static DataWizardSettings Settings()
    {
        var settings = DataWizardSettings.CreateDefault();
        settings.Conversion.Overwrite = true;
        settings.Conversion.WriteByteOrderMark = false;
        settings.Normalize();
        return settings;
    }

    private static (string[] Values, string[] Types) ReadSheet(string xlsxPath)
    {
        using var package = SpreadsheetDocument.Open(xlsxPath, false);
        var sheetData = package.WorkbookPart!
            .WorksheetParts.First()
            .Worksheet!
            .GetFirstChild<SheetData>()!;

        var values = new List<string>();
        var types = new List<string>();

        foreach (var cell in sheetData.Elements<Row>().SelectMany(row => row.Elements<Cell>()))
        {
            var isText = cell.DataType?.Value == CellValues.InlineString;
            types.Add(isText ? "text" : "number");
            values.Add(isText
                ? cell.InlineString?.Text?.Text ?? string.Empty
                : cell.CellValue?.Text ?? string.Empty);
        }

        return ([.. values], [.. types]);
    }

    [Fact]
    public void AThousandsSeparatorValueInSampleBecomesANumberCell()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("prices.csv",
            "label;price",
            "a;19,99",
            "b;4,50",
            "c;1.234,56");

        var result = new ConversionService(Settings()).Convert(source);
        Assert.True(result.Success, result.ErrorMessage);

        Assert.Equal(FieldDataType.Decimal, result.Analysis!.Columns[1].EffectiveType);

        var (values, types) = ReadSheet(result.OutputPaths[0]);

        Assert.Equal(["label", "price", "a", "19.99", "b", "4.5", "c", "1234.56"], values);
        Assert.Equal(["text", "text", "text", "number", "text", "number", "text", "number"], types);
    }

    [Fact]
    public void AThousandsSeparatorValueOutsideTheAnalysisSampleBecomesANumberCell()
    {
        using var workspace = new TempWorkspace();

        var lines = new List<string> { "label;price" };
        for (var i = 0; i < 205; i++)
            lines.Add($"row{i};19,99");

        // Beyond the 200-row analysis sample, so the column type is decided
        // without ever seeing this value.
        lines.Add("big;1.234,56");

        var source = workspace.WriteLines("prices.csv", lines.ToArray());

        var result = new ConversionService(Settings()).Convert(source);
        Assert.True(result.Success, result.ErrorMessage);

        Assert.Equal(FieldDataType.Decimal, result.Analysis!.Columns[1].EffectiveType);

        var (values, types) = ReadSheet(result.OutputPaths[0]);

        Assert.Equal(206, types.Count(t => t == "number"));
        Assert.Equal("1234.56", values[^1]);
        Assert.Equal("number", types[^1]);
    }

    /// <summary>
    /// A column mixing whole numbers and decimals (12 next to 2,75) used to be
    /// classified as text because neither kind reached the dominance threshold
    /// on its own. Every value must still end up as a number cell.
    /// </summary>
    [Fact]
    public void AColumnMixingIntegersAndDecimalsBecomesNumberCells()
    {
        using var workspace = new TempWorkspace();
        // The column names are deliberately neutral: "menge" and "preis" would
        // match the built-in field rules and settle the type without detection.
        var source = workspace.WriteLines("prices.csv",
            "id;c2;c3",
            "E-EURO;12;2,75",
            "Kabel;0,45;125,5",
            "Schrauben;0,89;500");

        var result = new ConversionService(Settings()).Convert(source);
        Assert.True(result.Success, result.ErrorMessage);

        Assert.Equal(FieldDataType.Decimal, result.Analysis!.Columns[1].EffectiveType);
        Assert.Equal(FieldDataType.Decimal, result.Analysis.Columns[2].EffectiveType);

        var (values, types) = ReadSheet(result.OutputPaths[0]);

        Assert.Equal(["id", "c2", "c3", "E-EURO", "12", "2.75",
                      "Kabel", "0.45", "125.5", "Schrauben", "0.89", "500"], values);
        Assert.Equal(["text", "text", "text", "text", "number", "number",
                      "text", "number", "number", "text", "number", "number"], types);
    }

    [Fact]
    public void PlainGermanDecimalsStayNumberCells()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("prices.csv",
            "label;price",
            "a;19,99",
            "b;4,50",
            "c;123,00");

        var result = new ConversionService(Settings()).Convert(source);
        Assert.True(result.Success, result.ErrorMessage);

        var (values, types) = ReadSheet(result.OutputPaths[0]);

        Assert.Equal(["label", "price", "a", "19.99", "b", "4.5", "c", "123"], values);
        Assert.Equal(["text", "text", "text", "number", "text", "number", "text", "number"], types);
    }
}
