using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DataWizard.App.Services;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;

namespace DataWizard.App.ViewModels;

/// <summary>
/// The Watcher tab: folder rules and the running state of each one.
/// </summary>
public partial class WatcherViewModel : ViewModelBase
{
    private readonly AppSession _session;
    private readonly DialogService _dialogs;

    [ObservableProperty]
    public partial WatchRuleRow? SelectedRule { get; set; }

    [ObservableProperty]
    public partial bool IsRunning { get; set; }

    [ObservableProperty]
    public partial string StatusText { get; set; }

    [ObservableProperty]
    public partial int ConvertedCount { get; set; }

    [ObservableProperty]
    public partial int FailedCount { get; set; }

    public WatcherViewModel(AppSession session, DialogService dialogs)
    {
        _session = session;
        _dialogs = dialogs;
        StatusText = "Stopped.";

        LoadFromSettings();

        _session.Watcher.StatusChanged += (_, _) => Dispatcher.UIThread.Post(RefreshStatus);
        _session.Watcher.FileConverted += (_, result) => Dispatcher.UIThread.Post(() => OnFileConverted(result));

        Rules.CollectionChanged += (_, _) =>
        {
            OnPropertyChanged(nameof(HasRules));
            StartCommand.NotifyCanExecuteChanged();
        };
    }

    /// <summary>The configured watch rules.</summary>
    public ObservableCollection<WatchRuleRow> Rules { get; } = [];

    /// <summary>Whether any rule exists.</summary>
    public bool HasRules => Rules.Count > 0;

    /// <summary>Whether the watcher starts with the application.</summary>
    public bool AutoStart
    {
        get => _session.Settings.Watcher.AutoStart;
        set
        {
            if (_session.Settings.Watcher.AutoStart == value)
                return;

            _session.Settings.Watcher.AutoStart = value;
            OnPropertyChanged();
        }
    }

    [RelayCommand]
    private void AddRule()
    {
        var row = new WatchRuleRow
        {
            Name = $"Rule {Rules.Count + 1}"
        };

        Rules.Add(row);
        SelectedRule = row;
    }

    [RelayCommand]
    private void RemoveRule()
    {
        if (SelectedRule is null)
            return;

        var index = Rules.IndexOf(SelectedRule);
        Rules.Remove(SelectedRule);
        SelectedRule = Rules.ElementAtOrDefault(Math.Min(index, Rules.Count - 1));
    }

    [RelayCommand]
    private void DuplicateRule()
    {
        if (SelectedRule is null)
            return;

        var copy = new WatchRuleRow(SelectedRule.ToModel())
        {
            Name = SelectedRule.DisplayName + " (copy)"
        };

        Rules.Add(copy);
        SelectedRule = copy;
    }

    [RelayCommand]
    private async Task BrowseInputFolderAsync()
    {
        if (SelectedRule is null)
            return;

        var folder = await _dialogs.PickFolderAsync("Select the folder to watch", SelectedRule.InputFolder);
        if (folder is not null)
            SelectedRule.InputFolder = folder;
    }

    [RelayCommand]
    private async Task BrowseOutputFolderAsync()
    {
        if (SelectedRule is null)
            return;

        var folder = await _dialogs.PickFolderAsync("Select the output folder", SelectedRule.OutputFolder);
        if (folder is not null)
            SelectedRule.OutputFolder = folder;
    }

    [RelayCommand]
    private async Task BrowseProcessedFolderAsync()
    {
        if (SelectedRule is null)
            return;

        var folder = await _dialogs.PickFolderAsync("Select the folder for processed sources", SelectedRule.ProcessedFolder);
        if (folder is not null)
            SelectedRule.ProcessedFolder = folder;
    }

    [RelayCommand]
    private async Task BrowseErrorFolderAsync()
    {
        if (SelectedRule is null)
            return;

        var folder = await _dialogs.PickFolderAsync("Select the folder for failed files", SelectedRule.ErrorFolder);
        if (folder is not null)
            SelectedRule.ErrorFolder = folder;
    }

    private bool CanStart() => HasRules && !IsRunning;

    [RelayCommand(CanExecute = nameof(CanStart))]
    private void Start()
    {
        SyncToSettings();

        ConvertedCount = 0;
        FailedCount = 0;

        _session.Watcher.Start(_session.Settings.Watcher);
        IsRunning = _session.Watcher.IsRunning;

        StatusText = IsRunning
            ? $"Watching {_session.Watcher.RuleStatuses.Count(s => s.IsActive)} folder(s)."
            : "Could not start; check the log.";

        RefreshStatus();
    }

    private bool CanStop() => IsRunning;

    [RelayCommand(CanExecute = nameof(CanStop))]
    private async Task StopAsync()
    {
        StatusText = "Stopping...";
        await _session.Watcher.StopAsync();

        IsRunning = false;
        StatusText = "Stopped.";
        RefreshStatus();
    }

    /// <summary>
    /// Starts the watcher if the settings ask for it. Called once at startup.
    /// </summary>
    public void StartIfConfigured()
    {
        if (AutoStart && HasRules && !IsRunning)
            Start();
    }

    private void OnFileConverted(ConversionResult result)
    {
        if (result.Success)
            ConvertedCount++;
        else
            FailedCount++;

        RefreshStatus();
    }

    private void RefreshStatus()
    {
        IsRunning = _session.Watcher.IsRunning;

        foreach (var status in _session.Watcher.RuleStatuses)
        {
            var row = Rules.FirstOrDefault(r => r.Name == status.Rule.Name && r.InputFolder == status.Rule.InputFolder);

            if (row is null)
                continue;

            row.Status = status.IsActive
                ? $"active - {status.SucceededCount} converted, {status.FailedCount} failed"
                : status.Error ?? "inactive";
        }

        if (IsRunning)
        {
            var pending = _session.Watcher.PendingCount;
            StatusText = $"Watching. {ConvertedCount} converted, {FailedCount} failed" +
                         (pending > 0 ? $", {pending} waiting to settle." : ".");
        }

        StartCommand.NotifyCanExecuteChanged();
        StopCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsRunningChanged(bool value)
    {
        StartCommand.NotifyCanExecuteChanged();
        StopCommand.NotifyCanExecuteChanged();
        OnPropertyChanged(nameof(CanEditRules));
    }

    /// <summary>Rules cannot be edited while the watcher is running.</summary>
    public bool CanEditRules => !IsRunning;

    /// <summary>Copies the rule rows back into the settings.</summary>
    public void SyncToSettings() =>
        _session.Settings.Watcher.Rules = Rules.Select(r => r.ToModel()).ToList();

    /// <summary>Rebuilds the rule rows from the settings.</summary>
    public void LoadFromSettings()
    {
        Rules.Clear();

        foreach (var rule in _session.Settings.Watcher.Rules)
            Rules.Add(new WatchRuleRow(rule));

        SelectedRule = Rules.FirstOrDefault();
        OnPropertyChanged(nameof(AutoStart));
    }
}
