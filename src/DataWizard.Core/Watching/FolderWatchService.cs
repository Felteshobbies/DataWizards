using System.Collections.Concurrent;
using System.IO.Enumeration;
using System.Threading.Channels;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;

namespace DataWizard.Core.Watching;

/// <summary>Live state of one watch rule.</summary>
public sealed class WatchRuleStatus
{
    /// <summary>The rule this status belongs to.</summary>
    public required WatchRule Rule { get; init; }

    /// <summary>Whether the rule is currently watching.</summary>
    public bool IsActive { get; set; }

    /// <summary>Why the rule is not watching, when it is not.</summary>
    public string? Error { get; set; }

    /// <summary>Files converted successfully since the watcher started.</summary>
    public int SucceededCount { get; set; }

    /// <summary>Files that could not be converted.</summary>
    public int FailedCount { get; set; }

    /// <summary>When the rule last converted something.</summary>
    public DateTimeOffset? LastActivity { get; set; }
}

/// <summary>
/// Watches folders and converts the files that appear in them.
/// </summary>
/// <remarks>
/// <para>
/// A file is not converted the moment it shows up. Large files arrive in pieces,
/// and reading one mid-copy yields truncated data, so each candidate must stop
/// growing for <see cref="WatchRule.StabilizationDelayMs"/> before it is queued.
/// </para>
/// <para>
/// Conversions run one at a time on a single worker. Files land in bursts, and
/// converting a folder's worth of them in parallel only makes the disk thrash.
/// </para>
/// <para>
/// Output written by the watcher is remembered for a short while and ignored on
/// the way back in. Without that, a rule watching both CSV and XLSX with its
/// output in the watched folder would convert its own results forever.
/// </para>
/// </remarks>
public sealed class FolderWatchService : IAsyncDisposable
{
    private readonly ConversionService _conversionService;
    private readonly List<FileSystemWatcher> _watchers = [];
    private readonly Dictionary<WatchRule, WatchRuleStatus> _statuses = [];
    private readonly ConcurrentDictionary<string, PendingFile> _pending = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, DateTimeOffset> _ownOutput = new(StringComparer.OrdinalIgnoreCase);
    private readonly Lock _gate = new();

    private static readonly TimeSpan OwnOutputMemory = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan StabilizationPollInterval = TimeSpan.FromMilliseconds(400);

    private Channel<QueuedFile>? _queue;
    private CancellationTokenSource? _cancellation;
    private Task? _worker;
    private Timer? _stabilizationTimer;

    public FolderWatchService(ConversionService conversionService)
    {
        _conversionService = conversionService ?? throw new ArgumentNullException(nameof(conversionService));
    }

    /// <summary>Raised for every log line the watcher produces.</summary>
    public event EventHandler<LogEntry>? Log;

    /// <summary>Raised after each conversion attempt, successful or not.</summary>
    public event EventHandler<ConversionResult>? FileConverted;

    /// <summary>Raised when a rule becomes active, stops, or its counters change.</summary>
    public event EventHandler? StatusChanged;

    /// <summary>Whether the watcher is running.</summary>
    public bool IsRunning { get; private set; }

    /// <summary>Files seen but not yet stable enough to convert.</summary>
    public int PendingCount => _pending.Count;

    /// <summary>Per-rule state, for the watcher panel.</summary>
    public IReadOnlyList<WatchRuleStatus> RuleStatuses
    {
        get
        {
            lock (_gate)
                return _statuses.Values.ToArray();
        }
    }

    /// <summary>
    /// Starts watching. Rules that fail validation are reported and skipped rather
    /// than preventing the others from running.
    /// </summary>
    public void Start(WatcherSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        if (IsRunning)
            Stop();

        lock (_gate)
        {
            _statuses.Clear();

            _cancellation = new CancellationTokenSource();
            _queue = Channel.CreateUnbounded<QueuedFile>(new UnboundedChannelOptions
            {
                SingleReader = true,
                SingleWriter = false
            });

            var started = new List<WatchRule>();

            foreach (var rule in settings.Rules)
            {
                var status = new WatchRuleStatus { Rule = rule };
                _statuses[rule] = status;

                if (!rule.Enabled)
                {
                    status.Error = "Disabled.";
                    continue;
                }

                if (!rule.Validate(out var error))
                {
                    status.Error = error;
                    WriteLog(LogLevel.Error, $"Rule '{rule.Name}' cannot start: {error}");
                    continue;
                }

                WarnAboutSelfFeedingRule(rule);

                try
                {
                    var watcher = CreateWatcher(rule);
                    _watchers.Add(watcher);
                    status.IsActive = true;
                    started.Add(rule);

                    WriteLog(LogLevel.Info,
                        $"Watching '{rule.InputFolder}' for {rule.FilePatterns}" +
                        (rule.IncludeSubfolders ? " including subfolders" : string.Empty));
                }
                catch (Exception ex) when (ex is IOException or ArgumentException or UnauthorizedAccessException)
                {
                    status.IsActive = false;
                    status.Error = ex.Message;
                    WriteLog(LogLevel.Error, $"Rule '{rule.Name}' could not start: {ex.Message}");
                }
            }

            if (started.Count == 0)
            {
                WriteLog(LogLevel.Warning, "No watch rule could be started.");
                _cancellation.Dispose();
                _cancellation = null;
                _queue = null;
                RaiseStatusChanged();
                return;
            }

            IsRunning = true;

            // Capture the reader and token as locals. The lambda would otherwise
            // close over the fields, which Stop clears - and a Start immediately
            // followed by Stop can null them before the thread pool has even run
            // the worker.
            var reader = _queue.Reader;
            var token = _cancellation.Token;

            _worker = Task.Run(() => ProcessQueueAsync(reader, token), token);
            _stabilizationTimer = new Timer(
                _ => PromoteStableFiles(),
                null,
                StabilizationPollInterval,
                StabilizationPollInterval);

            // Only now, with IsRunning set, will Observe accept anything - so the
            // existing-file sweep has to come after the watcher is live.
            foreach (var rule in started.Where(r => r.ProcessExistingFiles))
                QueueExistingFiles(rule);
        }

        RaiseStatusChanged();
    }

    /// <summary>Stops watching and waits for the file in flight to finish.</summary>
    public void Stop() => StopAsync().GetAwaiter().GetResult();

    /// <summary>Stops watching and waits for the file in flight to finish.</summary>
    public async Task StopAsync()
    {
        Timer? timer;
        Task? worker;
        CancellationTokenSource? cancellation;

        lock (_gate)
        {
            if (!IsRunning)
                return;

            IsRunning = false;

            foreach (var watcher in _watchers)
            {
                watcher.EnableRaisingEvents = false;
                watcher.Dispose();
            }

            _watchers.Clear();
            _pending.Clear();

            _queue?.Writer.TryComplete();

            timer = _stabilizationTimer;
            worker = _worker;
            cancellation = _cancellation;

            _stabilizationTimer = null;
            _worker = null;
            _cancellation = null;

            foreach (var status in _statuses.Values)
                status.IsActive = false;
        }

        if (timer is not null)
            await timer.DisposeAsync().ConfigureAwait(false);

        if (worker is not null)
        {
            try
            {
                await worker.ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                // Expected when the worker is cancelled mid-file.
            }
        }

        cancellation?.Dispose();

        WriteLog(LogLevel.Info, "Watcher stopped.");
        RaiseStatusChanged();
    }

    private FileSystemWatcher CreateWatcher(WatchRule rule)
    {
        var watcher = new FileSystemWatcher(rule.InputFolder)
        {
            IncludeSubdirectories = rule.IncludeSubfolders,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size,
            InternalBufferSize = 64 * 1024
        };

        watcher.Created += (_, e) => Observe(rule, e.FullPath);
        watcher.Changed += (_, e) => Observe(rule, e.FullPath);
        watcher.Renamed += (_, e) => Observe(rule, e.FullPath);
        watcher.Error += (_, e) => WriteLog(
            LogLevel.Error,
            $"Rule '{rule.Name}' lost events: {e.GetException().Message}");

        watcher.EnableRaisingEvents = true;
        return watcher;
    }

    /// <summary>
    /// Points out a configuration that would feed on its own output before it does
    /// any damage.
    /// </summary>
    private void WarnAboutSelfFeedingRule(WatchRule rule)
    {
        var output = string.IsNullOrWhiteSpace(rule.OutputFolder) ? rule.InputFolder : rule.OutputFolder;

        if (!PathsEqual(output, rule.InputFolder))
            return;

        var patterns = rule.ResolveFilePatterns();
        var watchesCsv = patterns.Any(p => ConversionService.CsvExtensions.Any(e => p.EndsWith(e, StringComparison.OrdinalIgnoreCase)));
        var watchesExcel = patterns.Any(p => ConversionService.ExcelExtensions.Any(e => p.EndsWith(e, StringComparison.OrdinalIgnoreCase)));

        if (watchesCsv && watchesExcel)
        {
            WriteLog(LogLevel.Warning,
                $"Rule '{rule.Name}' watches both CSV and Excel files and writes into the same folder. " +
                "Results are ignored for two minutes after being written, but a separate output folder is safer.");
        }
    }

    private void QueueExistingFiles(WatchRule rule)
    {
        var option = rule.IncludeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var count = 0;

        foreach (var pattern in rule.ResolveFilePatterns())
        {
            IEnumerable<string> files;

            try
            {
                files = Directory.EnumerateFiles(rule.InputFolder, pattern, option);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                WriteLog(LogLevel.Error, $"Rule '{rule.Name}' could not list '{rule.InputFolder}': {ex.Message}");
                continue;
            }

            foreach (var file in files)
            {
                Observe(rule, file);
                count++;
            }
        }

        if (count > 0)
            WriteLog(LogLevel.Info, $"Rule '{rule.Name}' picked up {count} existing file(s).");
    }

    /// <summary>
    /// Records that a file was seen. Nothing is converted here - the stabilization
    /// pass decides when the file has finished arriving.
    /// </summary>
    private void Observe(WatchRule rule, string path)
    {
        if (!IsRunning)
            return;

        if (!MatchesRule(rule, path))
            return;

        if (IsOwnOutput(path))
            return;

        long size;
        try
        {
            var info = new FileInfo(path);

            if (!info.Exists)
            {
                _pending.TryRemove(path, out _);
                return;
            }

            size = info.Length;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return;
        }

        _pending.AddOrUpdate(
            path,
            _ => new PendingFile(rule, size, DateTimeOffset.UtcNow),
            (_, existing) => existing.Size == size
                ? existing
                : existing with { Size = size, LastChanged = DateTimeOffset.UtcNow });
    }

    private static bool MatchesRule(WatchRule rule, string path)
    {
        if (!ConversionService.CanConvert(path))
            return false;

        var name = Path.GetFileName(path);

        foreach (var pattern in rule.ResolveFilePatterns())
        {
            if (FileSystemName.MatchesSimpleExpression(pattern, name, ignoreCase: true))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Moves files that have stopped changing into the conversion queue.
    /// </summary>
    private void PromoteStableFiles()
    {
        var writer = _queue?.Writer;
        if (writer is null || !IsRunning)
            return;

        var now = DateTimeOffset.UtcNow;
        PruneOwnOutput(now);

        foreach (var (path, pending) in _pending.ToArray())
        {
            if ((now - pending.LastChanged).TotalMilliseconds < pending.Rule.StabilizationDelayMs)
                continue;

            long size;
            try
            {
                var info = new FileInfo(path);

                if (!info.Exists)
                {
                    _pending.TryRemove(path, out _);
                    continue;
                }

                size = info.Length;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                continue;
            }

            // Grew since the last look, so it is still being written.
            if (size != pending.Size)
            {
                _pending.TryUpdate(path, pending with { Size = size, LastChanged = now }, pending);
                continue;
            }

            if (!_pending.TryRemove(path, out _))
                continue;

            writer.TryWrite(new QueuedFile(path, pending.Rule, 0));
        }
    }

    private async Task ProcessQueueAsync(
        ChannelReader<QueuedFile> reader,
        CancellationToken cancellationToken)
    {
        try
        {
            await foreach (var item in reader.ReadAllAsync(cancellationToken).ConfigureAwait(false))
            {
                await ConvertQueuedFileAsync(item, cancellationToken).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            // Normal shutdown.
        }
    }

    private async Task ConvertQueuedFileAsync(QueuedFile item, CancellationToken cancellationToken)
    {
        var rule = item.Rule;

        if (!File.Exists(item.Path))
            return;

        var outputFolder = string.IsNullOrWhiteSpace(rule.OutputFolder) ? null : rule.OutputFolder;

        // Claim the output paths before writing them. The filesystem raises Created
        // the instant the package is opened, so registering them afterwards is too
        // late - the watcher would already have queued its own result.
        foreach (var predicted in _conversionService.PredictOutputPaths(item.Path, outputFolder))
            _ownOutput[predicted] = DateTimeOffset.UtcNow;

        ConversionResult result;
        try
        {
            result = await _conversionService
                .ConvertAsync(item.Path, outputFolder, null, cancellationToken)
                .ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            result = new ConversionResult
            {
                SourcePath = item.Path,
                Direction = ConversionService.DirectionFor(item.Path) ?? ConversionDirection.CsvToExcel,
                Success = false,
                ErrorMessage = ex.Message,
                Exception = ex
            };
        }

        if (result.Success)
        {
            foreach (var output in result.OutputPaths)
                _ownOutput[output] = DateTimeOffset.UtcNow;

            UpdateStatus(rule, success: true);
            HandleSourceFile(rule, item.Path);
            FileConverted?.Invoke(this, result);
            return;
        }

        if (result.WasCancelled)
            return;

        if (item.Attempt + 1 < Math.Max(1, rule.MaxRetries))
        {
            WriteLog(LogLevel.Warning,
                $"{Path.GetFileName(item.Path)} failed ({result.ErrorMessage}); " +
                $"retrying in {rule.RetryDelayMs} ms " +
                $"({item.Attempt + 2} of {rule.MaxRetries}).");

            await Task.Delay(Math.Max(0, rule.RetryDelayMs), cancellationToken).ConfigureAwait(false);
            _queue?.Writer.TryWrite(item with { Attempt = item.Attempt + 1 });
            return;
        }

        WriteLog(LogLevel.Error,
            $"{Path.GetFileName(item.Path)} failed after {rule.MaxRetries} attempt(s): {result.ErrorMessage}");

        UpdateStatus(rule, success: false);
        MoveToErrorFolder(rule, item.Path);
        FileConverted?.Invoke(this, result);
    }

    /// <summary>Applies the rule's source action once a file has been converted.</summary>
    private void HandleSourceFile(WatchRule rule, string path)
    {
        try
        {
            switch (rule.SourceAction)
            {
                case SourceFileAction.Delete:
                    File.Delete(path);
                    WriteLog(LogLevel.Detail, $"Deleted source {Path.GetFileName(path)}");
                    break;

                case SourceFileAction.Move:
                    var target = MoveFile(path, rule.ProcessedFolder);
                    if (target is not null)
                        WriteLog(LogLevel.Detail, $"Moved source to {target}");
                    break;
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            WriteLog(LogLevel.Warning, $"Could not apply source action to {Path.GetFileName(path)}: {ex.Message}");
        }
    }

    private void MoveToErrorFolder(WatchRule rule, string path)
    {
        if (string.IsNullOrWhiteSpace(rule.ErrorFolder))
            return;

        try
        {
            var target = MoveFile(path, rule.ErrorFolder);
            if (target is not null)
                WriteLog(LogLevel.Info, $"Moved failed file to {target}");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            WriteLog(LogLevel.Warning, $"Could not move {Path.GetFileName(path)} to the error folder: {ex.Message}");
        }
    }

    /// <summary>
    /// Moves a file into a folder, adding a numeric suffix rather than overwriting
    /// a file that is already there.
    /// </summary>
    private static string? MoveFile(string path, string folder)
    {
        if (string.IsNullOrWhiteSpace(folder))
            return null;

        Directory.CreateDirectory(folder);

        var stem = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);
        var target = Path.Combine(folder, stem + extension);

        for (var counter = 1; File.Exists(target) && counter < 10_000; counter++)
            target = Path.Combine(folder, $"{stem}_{counter}{extension}");

        File.Move(path, target, overwrite: false);
        return target;
    }

    private bool IsOwnOutput(string path)
    {
        if (!_ownOutput.TryGetValue(path, out var written))
            return false;

        if (DateTimeOffset.UtcNow - written <= OwnOutputMemory)
            return true;

        _ownOutput.TryRemove(path, out _);
        return false;
    }

    private void PruneOwnOutput(DateTimeOffset now)
    {
        foreach (var (path, written) in _ownOutput.ToArray())
        {
            if (now - written > OwnOutputMemory)
                _ownOutput.TryRemove(path, out _);
        }
    }

    private static bool PathsEqual(string left, string right)
    {
        try
        {
            return string.Equals(
                Path.GetFullPath(left).TrimEnd(Path.DirectorySeparatorChar),
                Path.GetFullPath(right).TrimEnd(Path.DirectorySeparatorChar),
                StringComparison.OrdinalIgnoreCase);
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    private void UpdateStatus(WatchRule rule, bool success)
    {
        lock (_gate)
        {
            if (!_statuses.TryGetValue(rule, out var status))
                return;

            if (success)
                status.SucceededCount++;
            else
                status.FailedCount++;

            status.LastActivity = DateTimeOffset.Now;
        }

        RaiseStatusChanged();
    }

    private void RaiseStatusChanged() => StatusChanged?.Invoke(this, EventArgs.Empty);

    private void WriteLog(LogLevel level, string message) =>
        Log?.Invoke(this, new LogEntry(DateTimeOffset.Now, level, message));

    public async ValueTask DisposeAsync() => await StopAsync().ConfigureAwait(false);

    /// <summary>A file seen by a watcher but not yet known to have finished arriving.</summary>
    private readonly record struct PendingFile(WatchRule Rule, long Size, DateTimeOffset LastChanged);

    /// <summary>A file ready to convert, with its retry count.</summary>
    private readonly record struct QueuedFile(string Path, WatchRule Rule, int Attempt);
}
