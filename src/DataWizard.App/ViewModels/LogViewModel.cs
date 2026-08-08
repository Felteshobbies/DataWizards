using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input.Platform;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DataWizard.App.Services;
using DataWizard.Core.Conversion;

namespace DataWizard.App.ViewModels;

/// <summary>One log line, styled by severity.</summary>
public sealed class LogRow
{
    public required LogEntry Entry { get; init; }

    public string Timestamp => Entry.Timestamp.ToString("HH:mm:ss");

    public string Message => Entry.Message;

    /// <summary>The severity, shown in its own column.</summary>
    public string LevelText => Entry.Level.ToString();

    /// <summary>
    /// Severity flags the view binds style classes to, so the colours stay in the
    /// theme and follow a light or dark switch.
    /// </summary>
    public bool IsSuccess => Entry.Level == LogLevel.Success;

    /// <inheritdoc cref="IsSuccess"/>
    public bool IsWarning => Entry.Level == LogLevel.Warning;

    /// <inheritdoc cref="IsSuccess"/>
    public bool IsError => Entry.Level == LogLevel.Error;

    /// <inheritdoc cref="IsSuccess"/>
    public bool IsDetail => Entry.Level == LogLevel.Detail;
}

/// <summary>
/// The Log tab: everything the application did, filterable by severity.
/// </summary>
public partial class LogViewModel : ViewModelBase
{
    private readonly AppSession _session;
    private readonly Func<TopLevel?> _topLevel;

    [ObservableProperty]
    public partial bool ShowDetails { get; set; }

    [ObservableProperty]
    public partial bool ShowInfo { get; set; }

    [ObservableProperty]
    public partial bool ShowWarnings { get; set; }

    [ObservableProperty]
    public partial bool ShowErrors { get; set; }

    [ObservableProperty]
    public partial string FilterText { get; set; }

    [ObservableProperty]
    public partial bool AutoScroll { get; set; }

    public LogViewModel(AppSession session, Func<TopLevel?> topLevel)
    {
        _session = session;
        _topLevel = topLevel;

        ShowInfo = true;
        ShowWarnings = true;
        ShowErrors = true;
        AutoScroll = true;
        FilterText = string.Empty;

        _session.LogEntries.CollectionChanged += OnSourceChanged;
        Rebuild();
    }

    /// <summary>The visible log lines.</summary>
    public ObservableCollection<LogRow> Rows { get; } = [];

    /// <summary>Summary shown under the log.</summary>
    public string SummaryText =>
        $"{Rows.Count} of {_session.LogEntries.Count} entries";

    [RelayCommand]
    private void ClearLog()
    {
        _session.LogEntries.Clear();
        Rebuild();
    }

    [RelayCommand]
    private async Task CopyAsync()
    {
        var clipboard = _topLevel()?.Clipboard;
        if (clipboard is null)
            return;

        var builder = new StringBuilder();

        foreach (var row in Rows)
            builder.AppendLine($"[{row.Timestamp}] {row.Entry.Level}: {row.Message}");

        await clipboard.SetTextAsync(builder.ToString());
        _session.Log(LogLevel.Info, $"Copied {Rows.Count} log line(s) to the clipboard.");
    }

    private void OnSourceChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        // Appending is the common case, so add just the new entry rather than
        // rebuilding the whole list on every log line.
        if (e.Action == NotifyCollectionChangedAction.Add && e.NewItems is not null)
        {
            foreach (var entry in e.NewItems.OfType<LogEntry>())
            {
                if (Matches(entry))
                    Rows.Add(new LogRow { Entry = entry });
            }

            OnPropertyChanged(nameof(SummaryText));
            return;
        }

        Rebuild();
    }

    private void Rebuild()
    {
        Rows.Clear();

        foreach (var entry in _session.LogEntries.Where(Matches))
            Rows.Add(new LogRow { Entry = entry });

        OnPropertyChanged(nameof(SummaryText));
    }

    private bool Matches(LogEntry entry)
    {
        var levelAllowed = entry.Level switch
        {
            LogLevel.Detail => ShowDetails,
            LogLevel.Info => ShowInfo,
            LogLevel.Success => ShowInfo,
            LogLevel.Warning => ShowWarnings,
            LogLevel.Error => ShowErrors,
            _ => true
        };

        if (!levelAllowed)
            return false;

        return string.IsNullOrWhiteSpace(FilterText)
               || entry.Message.Contains(FilterText, StringComparison.OrdinalIgnoreCase);
    }

    partial void OnShowDetailsChanged(bool value) => Rebuild();
    partial void OnShowInfoChanged(bool value) => Rebuild();
    partial void OnShowWarningsChanged(bool value) => Rebuild();
    partial void OnShowErrorsChanged(bool value) => Rebuild();
    partial void OnFilterTextChanged(string value) => Rebuild();
}
