using System.Diagnostics;
using System.Text;
using DataWizard.App.Services;
using DataWizard.App.ViewModels;
using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;

namespace DataWizard.App.Tests;

/// <summary>
/// The analysis panel is closed by default and opens on its own only when the
/// detection was genuinely ambiguous. These tests pin down what counts as
/// ambiguous, because an over-eager rule makes the panel as intrusive as it was
/// when it was always open.
/// </summary>
public class AnalysisPanelTests : IDisposable
{
    private readonly string _root = Path.Combine(
        Path.GetTempPath(), "DataWizardPanelTests", Guid.NewGuid().ToString("N"));

    public AnalysisPanelTests() => Directory.CreateDirectory(_root);

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_root))
                Directory.Delete(_root, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    private string WriteFile(string name, string content, Encoding? encoding = null)
    {
        var path = Path.Combine(_root, name);
        File.WriteAllText(path, content, encoding ?? new UTF8Encoding(false));
        return path;
    }

    private static AnalysisPreview Analyze(string path, DataWizardSettings? settings = null)
    {
        settings ??= DataWizardSettings.CreateDefault();
        return new AnalysisPreview(new CsvAnalyzer(settings.Detection).Analyze(path));
    }

    [Fact]
    public void ACleanFileDoesNotAskForAttention()
    {
        var path = WriteFile("clean.csv",
            "id;name;price;order_date\r\n" +
            "1;Widget;19.99;2025-01-31\r\n" +
            "2;Gadget;4.50;2025-02-01\r\n" +
            "3;Sprocket;7.25;2025-02-02\r\n");

        var preview = Analyze(path);

        Assert.False(preview.NeedsAttention);
        Assert.Empty(preview.AttentionReasons);
        Assert.Equal("Detection looks unambiguous.", preview.AttentionHeadline);
    }

    [Fact]
    public void RaggedRowsAskForAttention()
    {
        var path = WriteFile("ragged.csv",
            "id;name;price\r\n1;Widget\r\n2;Gadget;4.50;extra\r\n");

        var preview = Analyze(path);

        Assert.True(preview.NeedsAttention);
        Assert.NotEmpty(preview.AttentionReasons);
    }

    /// <summary>
    /// Nothing is wrong here, but the decision could have gone either way - the
    /// case most worth surfacing, and the one a warnings-only rule would miss.
    /// </summary>
    [Fact]
    public void ABorderlineHeaderScoreAsksForAttention()
    {
        var path = WriteFile("borderline.csv", "alpha;beta\r\n1;x\r\n2;y\r\n");

        var settings = DataWizardSettings.CreateDefault();

        // Put the threshold exactly on the score this file produces.
        var probe = new CsvAnalyzer(settings.Detection).Analyze(path);
        settings.Detection.HeaderScoreThreshold = probe.HeaderResult.Score;

        var preview = Analyze(path, settings);

        Assert.True(preview.NeedsAttention);
        Assert.Contains(preview.AttentionReasons, r => r.Contains("header decision was close"));
    }

    [Fact]
    public void AConfidentHeaderDecisionDoesNotAskForAttention()
    {
        var path = WriteFile("confident.csv",
            "id;name;price\r\n1;Widget;19.99\r\n2;Gadget;4.50\r\n3;Sprocket;7.25\r\n");

        var settings = DataWizardSettings.CreateDefault();
        settings.Detection.HeaderScoreThreshold = 0.2;

        var preview = Analyze(path, settings);

        Assert.DoesNotContain(preview.AttentionReasons, r => r.Contains("header decision was close"));
    }

    /// <summary>
    /// Whole numbers among decimals are simply what a numeric column looks like.
    /// Treating that as a disagreement fired on nearly every file and made the
    /// panel open constantly.
    /// </summary>
    [Fact]
    public void IntegersAmongDecimalsAreNotAnInconsistency()
    {
        var path = WriteFile("numeric.csv",
            "label;amount\r\nа;434\r\nb;434\r\nc;545.4\r\nd;12\r\n".Replace("а", "a"));

        var preview = Analyze(path);

        Assert.DoesNotContain(preview.AttentionReasons, r => r.Contains("more than one kind of value"));
    }

    /// <summary>
    /// A postal code column holding both <c>01067</c> and <c>80331</c> types as
    /// text and integer, but a rule settles it - so there is nothing to review.
    /// </summary>
    [Fact]
    public void AColumnSettledByARuleIsNotAnInconsistency()
    {
        var path = WriteFile("zips.csv",
            "name;zip\r\nAlpha;01067\r\nBeta;80331\r\nGamma;20095\r\n");

        var preview = Analyze(path);

        var zip = preview.Columns.Single(c => c.Name == "zip");
        Assert.Equal("Text", zip.EffectiveType);
        Assert.False(string.IsNullOrEmpty(zip.Rule));
        Assert.DoesNotContain(preview.AttentionReasons, r => r.Contains("more than one kind of value"));
    }

    [Fact]
    public void AColumnOfMixedValuesAsksForAttention()
    {
        var path = WriteFile("mixed.csv",
            "code;value\r\nA1;1\r\nB2;2\r\nC3;not a number\r\n");

        var preview = Analyze(path);

        Assert.True(preview.NeedsAttention);
        Assert.Contains(preview.AttentionReasons, r => r.Contains("more than one kind of value"));
    }

    [Fact]
    public void ForcedHeaderModesAreNeverCalledBorderline()
    {
        var path = WriteFile("forced.csv", "alpha;beta\r\n1;x\r\n2;y\r\n");

        var settings = DataWizardSettings.CreateDefault();
        settings.Detection.HeaderMode = HeaderMode.Always;

        var preview = Analyze(path, settings);

        Assert.DoesNotContain(preview.AttentionReasons, r => r.Contains("header decision was close"));
    }

    [Fact]
    public void TheCompactSummaryNamesTheEssentials()
    {
        var path = WriteFile("summary.csv", "id;name\r\n1;Widget\r\n2;Gadget\r\n");

        var preview = Analyze(path);

        Assert.Contains("Semicolon", preview.Compact);
        Assert.Contains("2 columns", preview.Compact);
        Assert.Contains("header", preview.Compact);
    }

    // ── View model behaviour ────────────────────────────────────────────────

    private ConvertViewModel CreateViewModel(DataWizardSettings settings)
    {
        var session = new AppSession(settings, Path.Combine(_root, "settings.json"));
        return new ConvertViewModel(session, new DialogService(() => null));
    }

    private static async Task<bool> WaitForPreviewAsync(ConvertViewModel model, int timeoutMs = 10_000)
    {
        var stopwatch = Stopwatch.StartNew();

        while (stopwatch.ElapsedMilliseconds < timeoutMs)
        {
            if (model.Preview is not null || model.PreviewError is not null)
                return true;

            await Task.Delay(25);
        }

        return false;
    }

    [Fact]
    public void ThePanelStartsClosedByDefault()
    {
        var model = CreateViewModel(DataWizardSettings.CreateDefault());

        Assert.False(model.IsAnalysisExpanded);
    }

    [Fact]
    public void ThePanelStartsOpenWhenTheSettingSaysSo()
    {
        var settings = DataWizardSettings.CreateDefault();
        settings.ShowAnalysisPanel = true;

        Assert.True(CreateViewModel(settings).IsAnalysisExpanded);
    }

    [Fact]
    public void TogglingRemembersTheChoice()
    {
        var settings = DataWizardSettings.CreateDefault();
        var model = CreateViewModel(settings);

        model.ToggleAnalysisCommand.Execute(null);

        Assert.True(model.IsAnalysisExpanded);
        Assert.True(settings.ShowAnalysisPanel);

        model.ToggleAnalysisCommand.Execute(null);

        Assert.False(model.IsAnalysisExpanded);
        Assert.False(settings.ShowAnalysisPanel);
    }

    [Fact]
    public async Task SelectingACleanFileLeavesThePanelClosed()
    {
        var path = WriteFile("clean.csv",
            "id;name;price;order_date\r\n" +
            "1;Widget;19.99;2025-01-31\r\n" +
            "2;Gadget;4.50;2025-02-01\r\n" +
            "3;Sprocket;7.25;2025-02-02\r\n");

        var model = CreateViewModel(DataWizardSettings.CreateDefault());
        model.AddFiles([path]);
        model.SelectedFile = model.Files[0];

        Assert.True(await WaitForPreviewAsync(model), "analysis did not complete");
        Assert.False(model.AnalysisNeedsAttention);
        Assert.False(model.IsAnalysisExpanded);
    }

    [Fact]
    public async Task SelectingAnAmbiguousFileOpensThePanel()
    {
        var path = WriteFile("ragged.csv",
            "id;name;price\r\n1;Widget\r\n2;Gadget;4.50;extra\r\n");

        var model = CreateViewModel(DataWizardSettings.CreateDefault());
        model.AddFiles([path]);
        model.SelectedFile = model.Files[0];

        Assert.True(await WaitForPreviewAsync(model), "analysis did not complete");
        Assert.True(model.AnalysisNeedsAttention);
        Assert.True(model.IsAnalysisExpanded);
    }

    [Fact]
    public void TheWidthIsClampedToSomethingUsable()
    {
        var settings = DataWizardSettings.CreateDefault();
        var model = CreateViewModel(settings);

        model.AnalysisPanelWidth = 10;
        Assert.Equal(320d, model.AnalysisPanelWidth);

        model.AnalysisPanelWidth = 5000;
        Assert.Equal(900d, model.AnalysisPanelWidth);

        model.AnalysisPanelWidth = 500;
        Assert.Equal(500d, model.AnalysisPanelWidth);
        Assert.Equal(500d, settings.AnalysisPanelWidth);
    }
}
