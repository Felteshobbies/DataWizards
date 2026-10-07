using DataWizard.Core.Configuration;
using DataWizard.Core.Diff;

namespace DataWizard.Core.Tests;

public class DiffProfileTests
{
    private static DiffProfile Profile() => new()
    {
        Name = "Prices Q1",
        ReferencePath = @"C:\data\reference.csv",
        CandidatePath = @"C:\data\candidate.xlsx",
        Options = new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["id"],
            AbsoluteTolerance = 0.5d,
            RelativeTolerance = 0.001d,
            DateTolerance = TimeSpan.FromMinutes(5),
            IgnoreCase = true,
            TrimValues = true,
            IgnoredColumns = ["updatedAt"]
        }
    };

    [Fact]
    public void AProfileRoundTripsThroughDisk()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.PathTo("profile.json");

        Profile().Save(path);
        var loaded = DiffProfile.Load(path);

        Assert.Equal("Prices Q1", loaded.Name);
        Assert.Equal(@"C:\data\reference.csv", loaded.ReferencePath);
        Assert.Equal(@"C:\data\candidate.xlsx", loaded.CandidatePath);
        Assert.Equal(DiffMatchMode.Keyed, loaded.Options.MatchMode);
        Assert.Equal(["id"], loaded.Options.KeyColumns);
        Assert.Equal(0.5d, loaded.Options.AbsoluteTolerance);
        Assert.Equal(0.001d, loaded.Options.RelativeTolerance);
        Assert.Equal(TimeSpan.FromMinutes(5), loaded.Options.DateTolerance);
        Assert.True(loaded.Options.IgnoreCase);
        Assert.Equal(["updatedAt"], loaded.Options.IgnoredColumns);
    }

    [Fact]
    public void TheEnumIsStoredAsStringSoTheFileStaysReadable()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.PathTo("profile.json");

        Profile().Save(path);
        var json = File.ReadAllText(path);

        Assert.Contains("\"keyed\"", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("\"matchMode\": 1", json);
    }

    [Fact]
    public void LoadingAMissingProfileThrows()
    {
        using var workspace = new TempWorkspace();

        Assert.Throws<FileNotFoundException>(() => DiffProfile.Load(workspace.PathTo("nope.json")));
    }

    [Fact]
    public void LoadingCorruptJsonThrowsWithThePathInTheMessage()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteFile("broken.json", "{ not json");

        var ex = Assert.Throws<InvalidDataException>(() => DiffProfile.Load(path));

        Assert.Contains("broken.json", ex.Message);
    }

    [Fact]
    public void AnEmptyFileIsRejected()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteFile("empty.json", "");

        Assert.Throws<InvalidDataException>(() => DiffProfile.Load(path));
    }

    [Fact]
    public void AProfileWithoutOptionsFallsBackToTheDefaults()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteFile("minimal.json", """
            {
                "version": 1,
                "name": "Minimal"
            }
            """);

        var loaded = DiffProfile.Load(path);

        Assert.Equal("Minimal", loaded.Name);
        Assert.NotNull(loaded.Options);
        Assert.Equal(DiffMatchMode.Positional, loaded.Options.MatchMode);
    }

    [Fact]
    public void ACloneIsIndependentOfTheOriginal()
    {
        var profile = Profile();
        var clone = profile.Clone();

        clone.Options.KeyColumns = ["other"];
        clone.Name = "Changed";

        Assert.Equal(["id"], profile.Options.KeyColumns);
        Assert.Equal("Prices Q1", profile.Name);
    }

    [Fact]
    public void SavingCreatesTheDirectoryWhenItIsMissing()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.PathTo("deep/nested/profile.json");

        Profile().Save(path);

        Assert.True(File.Exists(path));
    }

    [Fact]
    public void DiffSettingsBookkeepingIsPersistedInTheMainSettings()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.PathTo("settings.json");

        var settings = DataWizardSettings.CreateDefault();
        settings.Diff.RememberProfile(@"C:\profiles\one.json");
        settings.Diff.RememberProfile(@"C:\profiles\two.json");
        settings.Save(path);

        var loaded = DataWizardSettings.Load(path);

        Assert.Equal(@"C:\profiles\two.json", loaded.Diff.LastProfilePath);
        Assert.Equal(
            [@"C:\profiles\two.json", @"C:\profiles\one.json"],
            loaded.Diff.RecentProfilePaths);
    }
}

public class DiffSettingsTests
{
    [Fact]
    public void RememberProfileMovesThePathToTheFront()
    {
        var settings = new DiffSettings
        {
            RecentProfilePaths = [@"C:\a.json", @"C:\b.json"]
        };

        settings.RememberProfile(@"C:\a.json");

        Assert.Equal(@"C:\a.json", settings.LastProfilePath);
        Assert.Equal([@"C:\a.json", @"C:\b.json"], settings.RecentProfilePaths);
    }

    [Fact]
    public void TheRecentListIsCappedAtTenEntries()
    {
        var settings = new DiffSettings();

        for (var i = 1; i <= 12; i++)
            settings.RememberProfile($@"C:\{i}.json");

        Assert.Equal(10, settings.RecentProfilePaths.Count);
        Assert.Equal(@"C:\12.json", settings.RecentProfilePaths[0]);
        Assert.Equal(@"C:\3.json", settings.RecentProfilePaths[^1]);
    }

    [Fact]
    public void BlankPathsAreNotRemembered()
    {
        var settings = new DiffSettings();

        settings.RememberProfile("  ");
        settings.RememberProfile(null);

        Assert.Null(settings.LastProfilePath);
        Assert.Empty(settings.RecentProfilePaths);
    }

    [Fact]
    public void NormalizeDropsBlankEntriesAndCapsTheList()
    {
        var settings = new DiffSettings
        {
            RecentProfilePaths = [@"C:\a.json", "", .. Enumerable.Repeat(@"C:\x.json", 15)]
        };

        settings.Normalize();

        Assert.Equal(DiffSettings.MaxRecentProfiles, settings.RecentProfilePaths.Count);
        Assert.DoesNotContain(string.Empty, settings.RecentProfilePaths);
    }

    [Fact]
    public void ACloneIsIndependentOfTheOriginal()
    {
        var settings = new DiffSettings
        {
            LastProfilePath = @"C:\a.json",
            RecentProfilePaths = [@"C:\a.json"]
        };
        var clone = settings.Clone();

        clone.RememberProfile(@"C:\b.json");

        Assert.Equal(@"C:\a.json", settings.LastProfilePath);
        Assert.Single(settings.RecentProfilePaths);
    }
}
