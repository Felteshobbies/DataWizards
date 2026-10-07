using System.Text;
using DataWizard.App.Services;
using DataWizard.App.ViewModels;
using DataWizard.Core.Configuration;
using DataWizard.Core.Diff;

namespace DataWizard.App.Tests;

/// <summary>
/// The Diff tab end to end: real CSV files on disk, a real comparison, and the
/// profile the tab saves for the next start.
/// </summary>
public class DiffTabTests : IDisposable
{
    private readonly string _root = Path.Combine(
        Path.GetTempPath(), "DataWizardDiffTests", Guid.NewGuid().ToString("N"));

    public DiffTabTests() => Directory.CreateDirectory(_root);

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

    private DiffViewModel CreateModel(DataWizardSettings? settings = null)
    {
        settings ??= DataWizardSettings.CreateDefault();
        return new DiffViewModel(
            new AppSession(settings, Path.Combine(_root, "settings.json")),
            new DialogService(() => null));
    }

    private string WriteCsv(string name, string content)
    {
        var path = Path.Combine(_root, name);
        File.WriteAllText(path, content.Replace("\n", "\r\n"));
        return path;
    }

    [Fact]
    public async Task PositionalComparisonReportsTheChangedCell()
    {
        var reference = WriteCsv("ref.csv", "id;name;amount\n1;Alpha;10\n2;Beta;20\n");
        var candidate = WriteCsv("cand.csv", "id;name;amount\n1;Alpha;10\n2;Gamma;20\n");
        var model = CreateModel();

        model.ReferencePath = reference;
        model.CandidatePath = candidate;
        await model.CompareCommand.ExecuteAsync(null);

        Assert.Null(model.ErrorText);
        Assert.NotNull(model.Report);
        Assert.Equal(1, model.Report!.DifferenceCount);

        var entry = model.Report.Entries.Single();
        Assert.Equal(DiffKind.ValueChanged, entry.Kind);
        Assert.Equal(2, entry.ReferenceRow);
        Assert.Equal(2, entry.CandidateRow);
        Assert.Equal("name", entry.Column);
        Assert.Equal("Beta", entry.ReferenceValue);
        Assert.Equal("Gamma", entry.CandidateValue);
    }

    [Fact]
    public async Task KeyedComparisonMatchesRowsOutOfOrder()
    {
        var reference = WriteCsv("ref.csv", "id;name\n1;Alpha\n2;Beta\n3;Gamma\n");
        var candidate = WriteCsv("cand.csv", "id;name\n3;Gamma\n1;Alpha2\n2;Beta\n");
        var model = CreateModel();

        model.ReferencePath = reference;
        model.CandidatePath = candidate;
        model.MatchMode = DiffMatchMode.Keyed;
        model.KeyColumns = "id";
        await model.CompareCommand.ExecuteAsync(null);

        Assert.Null(model.ErrorText);
        Assert.Equal(1, model.Report!.DifferenceCount);

        var entry = model.Report.Entries.Single();
        Assert.Equal(DiffKind.ValueChanged, entry.Kind);
        Assert.Equal(1, entry.ReferenceRow);
        Assert.Equal(2, entry.CandidateRow);
        Assert.Equal("name", entry.Column);
    }

    [Fact]
    public async Task KeyedComparisonWithoutKeyColumnsIsRefused()
    {
        var reference = WriteCsv("ref.csv", "id;name\n1;Alpha\n");
        var candidate = WriteCsv("cand.csv", "id;name\n1;Alpha\n");
        var model = CreateModel();

        model.ReferencePath = reference;
        model.CandidatePath = candidate;
        model.MatchMode = DiffMatchMode.Keyed;
        await model.CompareCommand.ExecuteAsync(null);

        Assert.NotNull(model.ErrorText);
        Assert.Null(model.Report);
        Assert.Contains("key column", model.ErrorText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task NumericToleranceHidesSmallDifferences()
    {
        var reference = WriteCsv("ref.csv", "code;value\n1;10.00\n2;20.00\n");
        var candidate = WriteCsv("cand.csv", "code;value\n1;10.005\n2;20.50\n");
        var model = CreateModel();

        model.ReferencePath = reference;
        model.CandidatePath = candidate;
        model.AbsoluteTolerance = "0.01";
        await model.CompareCommand.ExecuteAsync(null);

        Assert.Null(model.ErrorText);
        Assert.Equal(1, model.Report!.DifferenceCount);
        Assert.Equal(2, model.Report.Entries.Single().ReferenceRow);
    }

    [Fact]
    public async Task TheGridIsCappedAtFiveThousandEntries()
    {
        var reference = WriteCsv("ref.csv", BuildBigCsv(start: 1));
        var candidate = WriteCsv("cand.csv", BuildBigCsv(start: 2));
        var model = CreateModel();

        model.ReferencePath = reference;
        model.CandidatePath = candidate;
        await model.CompareCommand.ExecuteAsync(null);

        Assert.Null(model.ErrorText);
        Assert.Equal(6000, model.Report!.DifferenceCount);
        Assert.Equal(5000, model.Entries.Count);
        Assert.True(model.IsTruncated);
    }

    [Fact]
    public void SavingAndLoadingARoundTripRestoresTheFields()
    {
        var reference = WriteCsv("ref.csv", "id;name\n1;Alpha\n");
        var candidate = WriteCsv("cand.csv", "id;name\n1;Alpha\n");
        var profilePath = Path.Combine(_root, "profile.json");

        var model = CreateModel();
        model.ReferencePath = reference;
        model.CandidatePath = candidate;
        model.ProfileName = "Stock check";
        model.MatchMode = DiffMatchMode.Keyed;
        model.KeyColumns = "id";
        model.AbsoluteTolerance = "0.01";
        model.IgnoreCase = true;
        model.TrimValues = false;
        model.IgnoredColumns = "note";
        model.CurrentProfile().Save(profilePath);

        var loaded = DiffProfile.Load(profilePath);
        Assert.Equal("Stock check", loaded.Name);
        Assert.Equal(reference, loaded.ReferencePath);
        Assert.Equal(candidate, loaded.CandidatePath);
        Assert.Equal(DiffMatchMode.Keyed, loaded.Options.MatchMode);
        Assert.Equal(["id"], loaded.Options.KeyColumns);
        Assert.Equal(0.01, loaded.Options.AbsoluteTolerance);
        Assert.True(loaded.Options.IgnoreCase);
        Assert.False(loaded.Options.TrimValues);
        Assert.Equal(["note"], loaded.Options.IgnoredColumns);

        // A fresh tab that remembers the profile opens with the same settings.
        var settings = DataWizardSettings.CreateDefault();
        settings.Diff.LastProfilePath = profilePath;
        var restored = CreateModel(settings);

        Assert.Equal(profilePath, restored.ProfilePath);
        Assert.Equal("Stock check", restored.ProfileName);
        Assert.Equal(reference, restored.ReferencePath);
        Assert.Equal(candidate, restored.CandidatePath);
        Assert.Equal(DiffMatchMode.Keyed, restored.MatchMode);
        Assert.Equal("id", restored.KeyColumns);
        Assert.Equal("0.01", restored.AbsoluteTolerance);
        Assert.True(restored.IgnoreCase);
        Assert.False(restored.TrimValues);
        Assert.Equal("note", restored.IgnoredColumns);
    }

    [Fact]
    public async Task ComparingWithAnActiveProfileResavesItAndRemembersIt()
    {
        var reference = WriteCsv("ref.csv", "id;name\n1;Alpha\n");
        var candidate = WriteCsv("cand.csv", "id;name\n1;Alpha\n");
        var profilePath = Path.Combine(_root, "profile.json");
        var settings = DataWizardSettings.CreateDefault();
        var model = CreateModel(settings);

        model.ReferencePath = reference;
        model.CandidatePath = candidate;
        model.ProfilePath = profilePath;
        model.MatchMode = DiffMatchMode.Keyed;
        model.KeyColumns = "id";

        await model.CompareCommand.ExecuteAsync(null);

        Assert.Null(model.ErrorText);
        Assert.Equal(profilePath, settings.Diff.LastProfilePath);
        Assert.Contains(profilePath, settings.Diff.RecentProfilePaths);

        var saved = DiffProfile.Load(profilePath);
        Assert.Equal(DiffMatchMode.Keyed, saved.Options.MatchMode);
        Assert.Equal(reference, saved.ReferencePath);
        Assert.Equal(candidate, saved.CandidatePath);
    }

    private static string BuildBigCsv(int start)
    {
        var builder = new StringBuilder("i\n");

        for (var i = start; i < start + 6000; i++)
            builder.Append(i).Append('\n');

        return builder.ToString();
    }
}
