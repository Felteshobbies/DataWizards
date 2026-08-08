using System.Diagnostics;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;
using DataWizard.Core.Watching;

namespace DataWizard.Core.Tests;

public class FolderWatchServiceTests
{
    private static DataWizardSettings Settings()
    {
        var settings = DataWizardSettings.CreateDefault();
        settings.Conversion.Overwrite = true;
        settings.Normalize();
        return settings;
    }

    /// <summary>
    /// Polls until <paramref name="condition"/> holds. The watcher is driven by
    /// filesystem events and a stabilization delay, so tests have to wait for
    /// something to happen rather than assert immediately.
    /// </summary>
    private static async Task<bool> WaitUntilAsync(Func<bool> condition, int timeoutMs = 15_000)
    {
        var stopwatch = Stopwatch.StartNew();

        while (stopwatch.ElapsedMilliseconds < timeoutMs)
        {
            if (condition())
                return true;

            await Task.Delay(50);
        }

        return condition();
    }

    private static WatchRule Rule(string input, string output) => new()
    {
        Name = "test",
        InputFolder = input,
        OutputFolder = output,
        FilePatterns = "*.csv",
        StabilizationDelayMs = 100,
        MaxRetries = 1,
        RetryDelayMs = 50
    };

    [Fact]
    public async Task ConvertsAFileThatAppearsInTheWatchedFolder()
    {
        using var workspace = new TempWorkspace();
        var input = workspace.CreateDirectory("in");
        var output = workspace.CreateDirectory("out");

        var service = new ConversionService(Settings());
        await using var watcher = new FolderWatchService(service);

        var results = new List<ConversionResult>();
        watcher.FileConverted += (_, r) => { lock (results) results.Add(r); };

        watcher.Start(new WatcherSettings { Rules = [Rule(input, output)] });

        await File.WriteAllTextAsync(Path.Combine(input, "data.csv"), "id;name\r\n1;Widget\r\n");

        var converted = await WaitUntilAsync(() =>
        {
            lock (results) return results.Count > 0;
        });

        Assert.True(converted, "the watcher did not convert the file");

        lock (results)
        {
            Assert.True(results[0].Success, results[0].ErrorMessage);
        }

        Assert.True(File.Exists(Path.Combine(output, "data.xlsx")));
    }

    [Fact]
    public async Task PicksUpFilesThatWereAlreadyThereWhenAsked()
    {
        using var workspace = new TempWorkspace();
        var input = workspace.CreateDirectory("in");
        var output = workspace.CreateDirectory("out");

        await File.WriteAllTextAsync(Path.Combine(input, "existing.csv"), "id;name\r\n1;Widget\r\n");

        var rule = Rule(input, output);
        rule.ProcessExistingFiles = true;

        await using var watcher = new FolderWatchService(new ConversionService(Settings()));
        watcher.Start(new WatcherSettings { Rules = [rule] });

        Assert.True(await WaitUntilAsync(() => File.Exists(Path.Combine(output, "existing.xlsx"))));
    }

    [Fact]
    public async Task IgnoresFilesThatDoNotMatchThePattern()
    {
        using var workspace = new TempWorkspace();
        var input = workspace.CreateDirectory("in");
        var output = workspace.CreateDirectory("out");

        await using var watcher = new FolderWatchService(new ConversionService(Settings()));
        watcher.Start(new WatcherSettings { Rules = [Rule(input, output)] });

        await File.WriteAllTextAsync(Path.Combine(input, "notes.txt"), "id;name\r\n1;Widget\r\n");
        await File.WriteAllTextAsync(Path.Combine(input, "data.csv"), "id;name\r\n1;Widget\r\n");

        Assert.True(await WaitUntilAsync(() => File.Exists(Path.Combine(output, "data.xlsx"))));
        Assert.False(File.Exists(Path.Combine(output, "notes.xlsx")));
    }

    [Fact]
    public async Task DeletesTheSourceWhenTheRuleSaysSo()
    {
        using var workspace = new TempWorkspace();
        var input = workspace.CreateDirectory("in");
        var output = workspace.CreateDirectory("out");

        var rule = Rule(input, output);
        rule.SourceAction = SourceFileAction.Delete;

        await using var watcher = new FolderWatchService(new ConversionService(Settings()));
        watcher.Start(new WatcherSettings { Rules = [rule] });

        var source = Path.Combine(input, "data.csv");
        await File.WriteAllTextAsync(source, "id;name\r\n1;Widget\r\n");

        Assert.True(await WaitUntilAsync(() => !File.Exists(source) && File.Exists(Path.Combine(output, "data.xlsx"))));
    }

    [Fact]
    public async Task MovesTheSourceToTheProcessedFolder()
    {
        using var workspace = new TempWorkspace();
        var input = workspace.CreateDirectory("in");
        var output = workspace.CreateDirectory("out");
        var processed = workspace.PathTo("done");

        var rule = Rule(input, output);
        rule.SourceAction = SourceFileAction.Move;
        rule.ProcessedFolder = processed;

        await using var watcher = new FolderWatchService(new ConversionService(Settings()));
        watcher.Start(new WatcherSettings { Rules = [rule] });

        await File.WriteAllTextAsync(Path.Combine(input, "data.csv"), "id;name\r\n1;Widget\r\n");

        Assert.True(await WaitUntilAsync(() => File.Exists(Path.Combine(processed, "data.csv"))));
        Assert.False(File.Exists(Path.Combine(input, "data.csv")));
    }

    /// <summary>
    /// A rule watching both directions with its output in the watched folder would
    /// otherwise convert its own results indefinitely.
    /// </summary>
    [Fact]
    public async Task DoesNotConvertItsOwnOutput()
    {
        using var workspace = new TempWorkspace();
        var folder = workspace.CreateDirectory("both");

        var rule = new WatchRule
        {
            Name = "loop risk",
            InputFolder = folder,
            OutputFolder = folder,
            FilePatterns = "*.csv;*.xlsx",
            StabilizationDelayMs = 100,
            MaxRetries = 1
        };

        await using var watcher = new FolderWatchService(new ConversionService(Settings()));

        var conversions = 0;
        watcher.FileConverted += (_, _) => Interlocked.Increment(ref conversions);

        var warnings = new List<string>();
        watcher.Log += (_, e) =>
        {
            if (e.Level == LogLevel.Warning)
                lock (warnings) warnings.Add(e.Message);
        };

        watcher.Start(new WatcherSettings { Rules = [rule] });

        await File.WriteAllTextAsync(Path.Combine(folder, "data.csv"), "id;name\r\n1;Widget\r\n");

        await WaitUntilAsync(() => Volatile.Read(ref conversions) >= 1);
        await Task.Delay(1500);

        // One conversion, not a cascade.
        Assert.Equal(1, Volatile.Read(ref conversions));

        lock (warnings)
        {
            Assert.Contains(warnings, w => w.Contains("same folder"));
        }
    }

    [Fact]
    public async Task MovesAFailingFileToTheErrorFolder()
    {
        using var workspace = new TempWorkspace();
        var input = workspace.CreateDirectory("in");
        var output = workspace.CreateDirectory("out");
        var errors = workspace.PathTo("errors");

        var rule = new WatchRule
        {
            Name = "errors",
            InputFolder = input,
            OutputFolder = output,
            FilePatterns = "*.xlsx",
            ErrorFolder = errors,
            StabilizationDelayMs = 100,
            MaxRetries = 1,
            RetryDelayMs = 50
        };

        await using var watcher = new FolderWatchService(new ConversionService(Settings()));
        watcher.Start(new WatcherSettings { Rules = [rule] });

        // Not a real workbook, so the conversion fails.
        await File.WriteAllTextAsync(Path.Combine(input, "broken.xlsx"), "definitely not a zip archive");

        Assert.True(await WaitUntilAsync(() => File.Exists(Path.Combine(errors, "broken.xlsx"))));
    }

    [Fact]
    public void AnInvalidRuleIsReportedWithoutStoppingTheOthers()
    {
        using var workspace = new TempWorkspace();
        var valid = workspace.CreateDirectory("in");

        var broken = new WatchRule { Name = "broken", InputFolder = workspace.PathTo("does-not-exist") };
        var working = Rule(valid, workspace.CreateDirectory("out"));

        var watcher = new FolderWatchService(new ConversionService(Settings()));

        try
        {
            var errors = new List<string>();
            watcher.Log += (_, e) =>
            {
                if (e.Level == LogLevel.Error)
                    errors.Add(e.Message);
            };

            watcher.Start(new WatcherSettings { Rules = [broken, working] });

            Assert.True(watcher.IsRunning);
            Assert.Contains(errors, e => e.Contains("broken"));

            var statuses = watcher.RuleStatuses;
            Assert.Contains(statuses, s => s.Rule == broken && !s.IsActive && s.Error is not null);
            Assert.Contains(statuses, s => s.Rule == working && s.IsActive);
        }
        finally
        {
            watcher.Stop();
        }
    }

    [Fact]
    public void ADisabledRuleIsNotStarted()
    {
        using var workspace = new TempWorkspace();
        var rule = Rule(workspace.CreateDirectory("in"), workspace.CreateDirectory("out"));
        rule.Enabled = false;

        var watcher = new FolderWatchService(new ConversionService(Settings()));

        try
        {
            watcher.Start(new WatcherSettings { Rules = [rule] });

            Assert.False(watcher.IsRunning);
            Assert.Contains(watcher.RuleStatuses, s => !s.IsActive && s.Error == "Disabled.");
        }
        finally
        {
            watcher.Stop();
        }
    }

    [Fact]
    public async Task StopIsIdempotentAndSafeToCallTwice()
    {
        using var workspace = new TempWorkspace();
        var watcher = new FolderWatchService(new ConversionService(Settings()));

        watcher.Start(new WatcherSettings
        {
            Rules = [Rule(workspace.CreateDirectory("in"), workspace.CreateDirectory("out"))]
        });

        Assert.True(watcher.IsRunning);

        await watcher.StopAsync();
        await watcher.StopAsync();

        Assert.False(watcher.IsRunning);
    }
}
