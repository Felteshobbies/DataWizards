using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;

namespace DataWizard.Core.Tests;

public class HeaderDetectorTests
{
    private static DetectionSettings Settings(Action<DetectionSettings>? configure = null)
    {
        var settings = DataWizardSettings.CreateDefault().Detection;
        configure?.Invoke(settings);
        settings.Normalize();
        return settings;
    }

    private static HeaderDetectionResult Detect(string sample, DetectionSettings? settings = null, char separator = ';')
    {
        settings ??= Settings();

        using var reader = new CsvRecordReader(new StringReader(sample), separator, settings.QuoteChar);
        var records = reader.ReadRecords().Where(r => !r.IsBlank).ToList();

        return new HeaderDetector(settings).Detect(records);
    }

    [Fact]
    public void RecognisesAHeaderOfKnownColumnNames()
    {
        var result = Detect("id;name;price\n1;Widget;9.99\n2;Gadget;4.50");

        Assert.True(result.HasHeader);
        Assert.Equal(["id", "name", "price"], result.FieldNames);
    }

    [Fact]
    public void RecognisesAHeaderFromTypeDivergenceAloneWhenNamesAreUnknown()
    {
        var result = Detect("alpha;beta;gamma\n1;2.5;2025-01-31\n4;5.5;2025-02-01\n7;8.5;2025-02-02");

        Assert.True(result.HasHeader);
    }

    [Fact]
    public void DataOnlyFilesAreNotTreatedAsHavingAHeader()
    {
        var result = Detect("1;2;3\n4;5;6\n7;8;9");

        Assert.False(result.HasHeader);
        Assert.Equal(["Column 1", "Column 2", "Column 3"], result.FieldNames);
    }

    /// <summary>
    /// The previous implementation indexed the second row using the first row's
    /// field count, threw <see cref="IndexOutOfRangeException"/> on ragged files and
    /// silently swallowed it - which meant header detection quietly failed.
    /// </summary>
    [Fact]
    public void ARaggedSecondRowDoesNotBreakDetection()
    {
        var result = Detect("id;name;price;note\n1;Widget\n2;Gadget;4.50");

        Assert.True(result.HasHeader);
        Assert.Equal(4, result.FieldNames.Length);
    }

    [Fact]
    public void QuotedHeaderCellsContainingTheSeparatorStayAligned()
    {
        var result = Detect("\"id;code\";name;price\n1;Widget;9.99\n2;Gadget;4.50");

        Assert.True(result.HasHeader);
        Assert.Equal(3, result.FieldNames.Length);
        Assert.Equal("id;code", result.FieldNames[0]);
    }

    [Fact]
    public void ForcedModesSkipScoringEntirely()
    {
        const string dataOnly = "1;2;3\n4;5;6";

        Assert.True(Detect(dataOnly, Settings(s => s.HeaderMode = HeaderMode.Always)).HasHeader);
        Assert.False(Detect("id;name;price\n1;a;2", Settings(s => s.HeaderMode = HeaderMode.Never)).HasHeader);
    }

    [Fact]
    public void ForcedNeverStillGeneratesColumnNames()
    {
        var result = Detect("id;name\n1;a", Settings(s => s.HeaderMode = HeaderMode.Never));

        Assert.Equal(["Column 1", "Column 2"], result.FieldNames);
    }

    [Fact]
    public void RaisingTheThresholdMakesDetectionStricter()
    {
        const string sample = "alpha;beta\n1;x\n2;y";

        var lenient = Detect(sample, Settings(s => s.HeaderScoreThreshold = 0.1));
        var strict = Detect(sample, Settings(s => s.HeaderScoreThreshold = 0.99));

        Assert.True(lenient.HasHeader);
        Assert.False(strict.HasHeader);
    }

    [Fact]
    public void EverySignalIsReportedWithItsReasoning()
    {
        var result = Detect("id;name;price\n1;Widget;9.99\n2;Gadget;4.50");

        Assert.Equal(4, result.Signals.Count);
        Assert.All(result.Signals, s => Assert.False(string.IsNullOrWhiteSpace(s.Detail)));
        Assert.Contains(result.Signals, s => s.Name == "Type divergence" && s.Applied);
    }

    [Fact]
    public void SignalsCanBeSwitchedOffIndividually()
    {
        var settings = Settings(s =>
        {
            s.UseKnownFieldNames = false;
            s.UseUniqueness = false;
            s.UseNonEmptyRule = false;
        });

        var result = Detect("id;name;price\n1;Widget;9.99", settings);

        Assert.Single(result.Signals);
        Assert.Equal("Type divergence", result.Signals[0].Name);
    }

    [Fact]
    public void TypeDivergenceIsNotAppliedWhenThereAreNoDataRows()
    {
        var result = Detect("id;name;price");

        var divergence = result.Signals.Single(s => s.Name == "Type divergence");
        Assert.False(divergence.Applied);
    }

    [Fact]
    public void EmptyHeaderCellsBecomeGeneratedNames()
    {
        var result = Detect("id;;price\n1;x;9.99", Settings(s => s.HeaderMode = HeaderMode.Always));

        Assert.Equal(["id", "Column 2", "price"], result.FieldNames);
    }

    [Fact]
    public void DuplicateHeaderNamesAreMadeUnique()
    {
        var result = Detect("name;name;name\n1;2;3", Settings(s => s.HeaderMode = HeaderMode.Always));

        Assert.Equal(["name", "name (2)", "name (3)"], result.FieldNames);
    }

    [Fact]
    public void CustomKnownFieldNamesAreHonoured()
    {
        var settings = Settings(s =>
        {
            s.KnownFieldNames = [new NamePattern { Pattern = "widget_", Match = MatchMode.StartsWith }];
            s.UseTypeDivergence = false;
            s.UseUniqueness = false;
            s.UseNonEmptyRule = false;
            s.HeaderScoreThreshold = 0.9;
        });

        Assert.True(Detect("widget_a;widget_b\n1;2", settings).HasHeader);
        Assert.False(Detect("gadget_a;gadget_b\n1;2", settings).HasHeader);
    }

    [Fact]
    public void AnInvalidRegexInAPatternIsIgnoredRatherThanThrowing()
    {
        var settings = Settings(s =>
            s.KnownFieldNames = [new NamePattern { Pattern = "([unclosed", Match = MatchMode.Regex }]);

        var result = Detect("id;name\n1;a", settings);

        Assert.NotNull(result);
    }
}
