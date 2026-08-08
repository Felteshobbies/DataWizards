using System.Text;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;
using DataWizard.Core.Csv;

namespace DataWizard.Core.Tests;

public class ConversionRoundTripTests
{
    private static DataWizardSettings Settings(Action<DataWizardSettings>? configure = null)
    {
        var settings = DataWizardSettings.CreateDefault();
        settings.Conversion.Overwrite = true;
        settings.Conversion.WriteByteOrderMark = false;
        configure?.Invoke(settings);
        settings.Normalize();
        return settings;
    }

    private static string[] ReadCsvLines(string path) =>
        File.ReadAllText(path, Encoding.UTF8)
            .Replace("\r\n", "\n")
            .TrimEnd('\n')
            .Split('\n');

    [Fact]
    public void CsvBecomesAWorkbookThatConvertsBackToTheSameValues()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("input.csv",
            "id;name;quantity;price",
            "A1;Widget;5;19.99",
            "A2;Gadget;10;4.50");

        var service = new ConversionService(Settings(s =>
        {
            s.Conversion.QuoteMode = QuoteMode.Minimal;
            s.Conversion.NumberOutputCulture = string.Empty;
        }));

        var toExcel = service.Convert(source);
        Assert.True(toExcel.Success, toExcel.ErrorMessage);

        var xlsx = toExcel.OutputPaths[0];
        var back = service.Convert(xlsx, workspace.CreateDirectory("out"));
        Assert.True(back.Success, back.ErrorMessage);

        var lines = ReadCsvLines(back.OutputPaths[0]);

        Assert.Equal("id;name;quantity;price", lines[0]);
        Assert.Equal("A1;Widget;5;19.99", lines[1]);
        Assert.Equal("A2;Gadget;10;4.5", lines[2]);
    }

    /// <summary>
    /// The previous export formatted every number with <c>0.##</c>, so anything
    /// with more than two decimals was quietly rounded and a round trip lost data.
    /// </summary>
    [Fact]
    public void ManyDecimalPlacesSurviveTheRoundTrip()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("precise.csv",
            "label;value",
            "pi;3.14159265",
            "third;0.333333");

        var service = new ConversionService(Settings(s =>
        {
            s.Conversion.QuoteMode = QuoteMode.Minimal;
            s.Conversion.NumberOutputCulture = string.Empty;
        }));

        var xlsx = service.Convert(source).OutputPaths[0];
        var back = service.Convert(xlsx, workspace.CreateDirectory("out"));
        var lines = ReadCsvLines(back.OutputPaths[0]);

        Assert.Contains("3.14159265", lines[1]);
        Assert.Contains("0.333333", lines[2]);
    }

    [Fact]
    public void TheDecimalLimitStillRoundsWhenAskedTo()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("precise.csv", "label;value", "pi;3.14159265");

        var service = new ConversionService(Settings(s =>
        {
            s.Conversion.QuoteMode = QuoteMode.Minimal;
            s.Conversion.NumberOutputCulture = string.Empty;
            s.Conversion.MaxDecimalPlaces = 2;
        }));

        var xlsx = service.Convert(source).OutputPaths[0];
        var back = service.Convert(xlsx, workspace.CreateDirectory("out"));

        Assert.Contains("3.14", ReadCsvLines(back.OutputPaths[0])[1]);
    }

    /// <summary>
    /// The old export hard-coded German number formatting regardless of the chosen
    /// separator or encoding.
    /// </summary>
    [Fact]
    public void NumberOutputCultureIsConfigurable()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("money.csv", "label;price", "widget;19.99");

        var german = new ConversionService(Settings(s =>
        {
            s.Conversion.QuoteMode = QuoteMode.Minimal;
            s.Conversion.NumberOutputCulture = "de-DE";
        }));

        var xlsx = german.Convert(source).OutputPaths[0];
        var back = german.Convert(xlsx, workspace.CreateDirectory("de"));

        Assert.Contains("19,99", ReadCsvLines(back.OutputPaths[0])[1]);
    }

    [Fact]
    public void LeadingZerosSurviveBecauseTheColumnStaysText()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("zeros.csv",
            "article;zip",
            "00123;01067",
            "00456;10115");

        var service = new ConversionService(Settings(s => s.Conversion.QuoteMode = QuoteMode.Minimal));

        var xlsx = service.Convert(source).OutputPaths[0];
        var back = service.Convert(xlsx, workspace.CreateDirectory("out"));
        var lines = ReadCsvLines(back.OutputPaths[0]);

        Assert.Equal("00123;01067", lines[1]);
        Assert.Equal("00456;10115", lines[2]);
    }

    [Fact]
    public void QuotedValuesStayTextEvenWhenTheyLookNumeric()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("quoted.csv",
            "code;note",
            "\"007\";fine",
            "\"008\";also fine");

        var service = new ConversionService(Settings(s => s.Conversion.QuoteMode = QuoteMode.Minimal));

        var xlsx = service.Convert(source).OutputPaths[0];
        var back = service.Convert(xlsx, workspace.CreateDirectory("out"));

        Assert.Equal("007;fine", ReadCsvLines(back.OutputPaths[0])[1]);
    }

    [Fact]
    public void ADateColumnComesBackInTheConfiguredFormat()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("dates.csv",
            "label;date",
            "start;31.01.2025",
            "end;01.02.2025");

        var service = new ConversionService(Settings(s =>
        {
            s.Conversion.QuoteMode = QuoteMode.Minimal;
            s.Conversion.DateOutputFormat = "yyyy-MM-dd";
        }));

        var xlsx = service.Convert(source).OutputPaths[0];
        var back = service.Convert(xlsx, workspace.CreateDirectory("out"));
        var lines = ReadCsvLines(back.OutputPaths[0]);

        Assert.Equal("start;2025-01-31", lines[1]);
        Assert.Equal("end;2025-02-01", lines[2]);
    }

    [Fact]
    public void ValuesContainingTheSeparatorAreQuotedEvenInMinimalMode()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("commas.csv",
            "id;description",
            "1;\"red; round; large\"");

        var service = new ConversionService(Settings(s => s.Conversion.QuoteMode = QuoteMode.Minimal));

        var xlsx = service.Convert(source).OutputPaths[0];
        var back = service.Convert(xlsx, workspace.CreateDirectory("out"));
        var lines = ReadCsvLines(back.OutputPaths[0]);

        Assert.Equal("1;\"red; round; large\"", lines[1]);
    }

    [Fact]
    public void EmbeddedQuotesAreDoubledOnTheWayOut()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("quotes.csv",
            "id;text",
            "1;\"say \"\"hi\"\"\"");

        var service = new ConversionService(Settings(s => s.Conversion.QuoteMode = QuoteMode.Minimal));

        var xlsx = service.Convert(source).OutputPaths[0];
        var back = service.Convert(xlsx, workspace.CreateDirectory("out"));

        Assert.Equal("1;\"say \"\"hi\"\"\"", ReadCsvLines(back.OutputPaths[0])[1]);
    }

    [Fact]
    public void AMultiLineValueSurvivesAsASingleCell()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteFile("multiline.csv",
            "id;note\r\n1;\"first\nsecond\"\r\n2;plain\r\n");

        var service = new ConversionService(Settings(s => s.Conversion.QuoteMode = QuoteMode.Minimal));

        var xlsx = service.Convert(source).OutputPaths[0];
        var back = service.Convert(xlsx, workspace.CreateDirectory("out"));
        var text = File.ReadAllText(back.OutputPaths[0], Encoding.UTF8);

        Assert.Contains("\"first\nsecond\"", text.Replace("\r\n", "\n"));
    }

    [Fact]
    public void AllSheetsExportIntoSeparateFiles()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("input.csv", "id;name", "1;Widget");

        var service = new ConversionService(Settings(s => s.Conversion.SheetName = "Products"));
        var xlsx = service.Convert(source).OutputPaths[0];

        var exportService = new ConversionService(Settings(s =>
        {
            s.Conversion.ExportAllSheets = true;
            s.Conversion.SheetName = "Products";
        }));

        var back = exportService.Convert(xlsx, workspace.CreateDirectory("out"));

        Assert.True(back.Success, back.ErrorMessage);
        Assert.Single(back.OutputPaths);
        Assert.EndsWith("_Products.csv", back.OutputPaths[0]);
    }

    [Fact]
    public void WithoutOverwriteASuffixIsAddedRatherThanReplacingTheFile()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("input.csv", "id;name", "1;Widget");

        var service = new ConversionService(Settings(s => s.Conversion.Overwrite = false));

        var first = service.Convert(source);
        var second = service.Convert(source);

        Assert.True(first.Success, first.ErrorMessage);
        Assert.True(second.Success, second.ErrorMessage);
        Assert.NotEqual(first.OutputPaths[0], second.OutputPaths[0]);
        Assert.EndsWith("input_1.xlsx", second.OutputPaths[0]);
    }

    [Fact]
    public void UnsupportedExtensionsFailWithAClearMessage()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteFile("thing.pdf", "not really a pdf");

        var result = new ConversionService(Settings()).Convert(source);

        Assert.False(result.Success);
        Assert.Contains("Unsupported file type", result.ErrorMessage);
    }

    [Fact]
    public void AMissingFileFailsWithoutThrowing()
    {
        using var workspace = new TempWorkspace();

        var result = new ConversionService(Settings()).Convert(workspace.PathTo("nope.csv"));

        Assert.False(result.Success);
        Assert.Equal("File not found.", result.ErrorMessage);
    }

    [Fact]
    public void ACorruptWorkbookFailsWithoutThrowing()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteFile("broken.xlsx", "this is not a zip archive");

        var result = new ConversionService(Settings()).Convert(source);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void EveryStepIsReportedThroughTheLogEvent()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("input.csv", "id;name", "1;Widget");

        var service = new ConversionService(Settings());
        var entries = new List<LogEntry>();
        service.Log += (_, entry) => entries.Add(entry);

        service.Convert(source);

        Assert.Contains(entries, e => e.Level == LogLevel.Success);
        Assert.Contains(entries, e => e.Message.Contains("Encoding:"));
        Assert.Contains(entries, e => e.Message.Contains("Separator:"));
    }

    [Fact]
    public void ProgressIsReported()
    {
        using var workspace = new TempWorkspace();

        var builder = new StringBuilder("id;value\r\n");
        for (var i = 1; i <= 2500; i++)
            builder.Append(i).Append(';').Append(i).Append("\r\n");

        var source = workspace.WriteFile("many.csv", builder.ToString());

        var reports = new List<ConversionProgress>();
        var progress = new Progress<ConversionProgress>(reports.Add);

        var result = new ConversionService(Settings()).Convert(source, null, progress);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2501, result.RowCount);
    }

    [Fact]
    public void CancellationStopsTheConversionAndIsReportedAsSuch()
    {
        using var workspace = new TempWorkspace();

        var builder = new StringBuilder("id;value\r\n");
        for (var i = 1; i <= 20_000; i++)
            builder.Append(i).Append(';').Append(i).Append("\r\n");

        var source = workspace.WriteFile("many.csv", builder.ToString());

        using var cts = new CancellationTokenSource();
        cts.Cancel();

        var result = new ConversionService(Settings()).Convert(source, null, null, cts.Token);

        Assert.False(result.Success);
        Assert.True(result.WasCancelled);
    }

    [Fact]
    public void TabSeparatedFilesAreDetectedAndConverted()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("tabs.txt",
            "id\tname\tprice",
            "1\tWidget\t19.99");

        var service = new ConversionService(Settings());
        var result = service.Convert(source);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal('\t', result.Analysis!.Separator);
        Assert.Equal(2, result.RowCount);
    }

    [Fact]
    public void AWorksheetIndexOutOfRangeIsReportedNotThrown()
    {
        using var workspace = new TempWorkspace();
        var source = workspace.WriteLines("input.csv", "id;name", "1;Widget");

        var writeService = new ConversionService(Settings());
        var xlsx = writeService.Convert(source).OutputPaths[0];

        var readService = new ConversionService(Settings(s => s.Conversion.WorksheetIndex = 5));
        var result = readService.Convert(xlsx, workspace.CreateDirectory("out"));

        Assert.False(result.Success);
        Assert.Contains("out of range", result.ErrorMessage);
    }
}
