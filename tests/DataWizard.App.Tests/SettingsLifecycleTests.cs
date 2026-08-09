using DataWizard.App.Services;
using DataWizard.App.ViewModels;
using DataWizard.Core.Configuration;

namespace DataWizard.App.Tests;

/// <summary>
/// The settings file is written on every shutdown from whatever the tabs are
/// holding. That makes an empty tab collection indistinguishable from "the user
/// deleted everything", so these tests pin down that a normal start-and-stop
/// cycle cannot quietly erase the pattern lists.
/// </summary>
public class SettingsLifecycleTests : IDisposable
{
    private readonly string _root = Path.Combine(
        Path.GetTempPath(), "DataWizardSettingsTests", Guid.NewGuid().ToString("N"));

    public SettingsLifecycleTests() => Directory.CreateDirectory(_root);

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

    private string SettingsPath => Path.Combine(_root, "settings.json");

    private MainWindowViewModel CreateShell(DataWizardSettings settings) =>
        new(new AppSession(settings, SettingsPath), () => null);

    [Fact]
    public async Task ShutdownPreservesTheDefaultPatternLists()
    {
        var settings = DataWizardSettings.CreateDefault();
        var expectedRules = settings.Detection.FieldRules.Count;
        var expectedNames = settings.Detection.KnownFieldNames.Count;

        Assert.True(expectedRules > 0);
        Assert.True(expectedNames > 0);

        var shell = CreateShell(settings);
        await shell.ShutdownAsync();

        var reloaded = DataWizardSettings.Load(SettingsPath);

        Assert.Equal(expectedRules, reloaded.Detection.FieldRules.Count);
        Assert.Equal(expectedNames, reloaded.Detection.KnownFieldNames.Count);
    }

    [Fact]
    public async Task ShutdownPreservesUserEditedPatterns()
    {
        var settings = DataWizardSettings.CreateDefault();
        settings.Detection.FieldRules.Insert(0, new FieldRule
        {
            Pattern = "^my_article_no$",
            Match = MatchMode.Regex,
            DataType = FieldDataType.Text
        });

        var shell = CreateShell(settings);
        await shell.ShutdownAsync();

        var reloaded = DataWizardSettings.Load(SettingsPath);

        Assert.Contains(reloaded.Detection.FieldRules, r => r.Pattern == "^my_article_no$");
    }

    [Fact]
    public async Task SavingTwiceInARowIsStable()
    {
        var settings = DataWizardSettings.CreateDefault();
        var shell = CreateShell(settings);

        shell.SaveSettingsCommand.Execute(null);
        var afterFirst = DataWizardSettings.Load(SettingsPath).Detection.FieldRules.Count;

        shell.SaveSettingsCommand.Execute(null);
        await shell.ShutdownAsync();

        var afterThird = DataWizardSettings.Load(SettingsPath).Detection.FieldRules.Count;

        Assert.Equal(afterFirst, afterThird);
        Assert.True(afterThird > 0);
    }

    /// <summary>
    /// Deliberately clearing the lists must still be honoured - the protection is
    /// against losing them by accident, not against an intentional edit.
    /// </summary>
    [Fact]
    public async Task ClearingThePatternListsOnPurposeIsHonoured()
    {
        var settings = DataWizardSettings.CreateDefault();
        var shell = CreateShell(settings);

        shell.Detection.HeaderPatterns.Clear();
        shell.Detection.FieldRules.Clear();

        await shell.ShutdownAsync();

        var reloaded = DataWizardSettings.Load(SettingsPath);

        Assert.Empty(reloaded.Detection.FieldRules);
        Assert.Empty(reloaded.Detection.KnownFieldNames);
    }

    [Fact]
    public async Task ResettingToDefaultsRestoresThePatternsInTheTabs()
    {
        var settings = DataWizardSettings.CreateDefault();
        var shell = CreateShell(settings);

        shell.Detection.HeaderPatterns.Clear();
        shell.Detection.FieldRules.Clear();

        shell.ConfirmResetSettingsCommand.Execute(null);

        Assert.NotEmpty(shell.Detection.FieldRules);
        Assert.NotEmpty(shell.Detection.HeaderPatterns);

        await shell.ShutdownAsync();

        Assert.NotEmpty(DataWizardSettings.Load(SettingsPath).Detection.FieldRules);
    }

    [Fact]
    public async Task TheAnalysisPanelStateSurvivesARestart()
    {
        var settings = DataWizardSettings.CreateDefault();
        var shell = CreateShell(settings);

        shell.Convert.ToggleAnalysisCommand.Execute(null);
        shell.Convert.AnalysisPanelWidth = 555;

        await shell.ShutdownAsync();

        var reloaded = DataWizardSettings.Load(SettingsPath);

        Assert.True(reloaded.ShowAnalysisPanel);
        Assert.Equal(555d, reloaded.AnalysisPanelWidth);
    }
}
