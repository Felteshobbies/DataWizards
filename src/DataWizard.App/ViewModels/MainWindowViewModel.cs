using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using Avalonia.Controls;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DataWizard.App.Services;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;

namespace DataWizard.App.ViewModels;

/// <summary>
/// The window shell: the tabs, the appearance selector, and settings persistence.
/// </summary>
public partial class MainWindowViewModel : ViewModelBase
{
    private readonly AppSession _session;

    [ObservableProperty]
    public partial string StatusText { get; set; }

    [ObservableProperty]
    public partial bool IsConfirmingReset { get; set; }

    public MainWindowViewModel(AppSession session, Func<TopLevel?> topLevel)
    {
        _session = session;

        var dialogs = new DialogService(topLevel);

        Convert = new ConvertViewModel(session, dialogs);
        Detection = new DetectionViewModel(session, dialogs);
        Output = new OutputViewModel(session, dialogs);
        Watcher = new WatcherViewModel(session, dialogs);
        Log = new LogViewModel(session, topLevel);

        StatusText = $"Settings: {session.SettingsPath}";
    }

    /// <summary>The Convert tab.</summary>
    public ConvertViewModel Convert { get; }

    /// <summary>The Detection tab.</summary>
    public DetectionViewModel Detection { get; }

    /// <summary>The Output tab.</summary>
    public OutputViewModel Output { get; }

    /// <summary>The Watcher tab.</summary>
    public WatcherViewModel Watcher { get; }

    /// <summary>The Log tab.</summary>
    public LogViewModel Log { get; }

    /// <summary>Window title including the version.</summary>
    public string Title => $"DataWizards {VersionText} - CSV and Excel converter";

    /// <summary>The assembly version, for the title and the help tab.</summary>
    public string VersionText =>
        Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.2.0";

    /// <summary>Appearance options for the selector.</summary>
    public IReadOnlyList<ThemeMode> ThemeModes { get; } = Enum.GetValues<ThemeMode>();

    /// <summary>The selected appearance.</summary>
    public ThemeMode Theme
    {
        get => _session.Settings.Theme;
        set
        {
            if (_session.Settings.Theme == value)
                return;

            _session.SetTheme(value);
            OnPropertyChanged();
        }
    }

    /// <summary>Where the settings file lives.</summary>
    public string SettingsPath => _session.SettingsPath;

    [RelayCommand]
    private void SaveSettings()
    {
        // The tabs hold grid rows that have not been written back yet.
        Detection.SyncToSettings();
        Watcher.SyncToSettings();

        if (_session.SaveSettings())
        {
            StatusText = $"Settings saved to {_session.SettingsPath}";
            _session.Log(LogLevel.Success, "Settings saved.");
        }
        else
        {
            StatusText = "Settings could not be saved; see the log.";
        }
    }

    [RelayCommand]
    private void BeginResetSettings() => IsConfirmingReset = true;

    [RelayCommand]
    private void CancelResetSettings() => IsConfirmingReset = false;

    [RelayCommand]
    private void ConfirmResetSettings()
    {
        IsConfirmingReset = false;

        _session.ResetToDefaults();
        Detection.LoadFromSettings();
        Output.Refresh();
        Watcher.LoadFromSettings();
        Detection.RunTestCommand.Execute(null);

        OnPropertyChanged(nameof(Theme));
        StatusText = "Settings reset to defaults. Save to keep them.";
    }

    [RelayCommand]
    private void OpenSettingsFolder()
    {
        var folder = Path.GetDirectoryName(_session.SettingsPath);

        if (string.IsNullOrWhiteSpace(folder))
            return;

        try
        {
            Directory.CreateDirectory(folder);
            Process.Start(new ProcessStartInfo { FileName = folder, UseShellExecute = true });
        }
        catch (Exception ex)
        {
            _session.Log(LogLevel.Warning, $"Could not open '{folder}': {ex.Message}");
        }
    }

    /// <summary>
    /// Saves on the way out, so the watcher rules and detection settings a user
    /// just edited are still there next time.
    /// </summary>
    public async Task ShutdownAsync()
    {
        Detection.SyncToSettings();
        Watcher.SyncToSettings();
        _session.SaveSettings();

        await _session.Watcher.StopAsync();
    }
}
