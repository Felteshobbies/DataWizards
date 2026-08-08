using System;
using System.Collections.ObjectModel;
using System.Linq;
using Avalonia.Threading;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;
using DataWizard.Core.Watching;

namespace DataWizard.App.Services;

/// <summary>
/// The state every tab shares: the settings, the services that act on them, and
/// the activity log.
/// </summary>
/// <remarks>
/// Log entries arrive from background threads - the conversion worker and the
/// folder watcher - so appending to the collection is marshalled onto the UI
/// thread here rather than at each call site.
/// </remarks>
public sealed class AppSession
{
    private const int MaxLogEntries = 5000;

    public AppSession(DataWizardSettings settings, string settingsPath)
    {
        Settings = settings;
        SettingsPath = settingsPath;

        ConversionService = new ConversionService(settings);
        ConversionService.Log += (_, entry) => AppendLog(entry);

        Watcher = new FolderWatchService(ConversionService);
        Watcher.Log += (_, entry) => AppendLog(entry);
    }

    /// <summary>The live settings instance. Editing it affects later conversions.</summary>
    public DataWizardSettings Settings { get; private set; }

    /// <summary>Where the settings are persisted.</summary>
    public string SettingsPath { get; }

    /// <summary>Performs conversions.</summary>
    public ConversionService ConversionService { get; }

    /// <summary>Watches folders.</summary>
    public FolderWatchService Watcher { get; }

    /// <summary>The activity log, newest last.</summary>
    public ObservableCollection<LogEntry> LogEntries { get; } = [];

    /// <summary>Raised when the theme setting changes.</summary>
    public event EventHandler<ThemeMode>? ThemeChanged;

    /// <summary>Raised when a log entry is appended, so the log view can scroll.</summary>
    public event EventHandler<LogEntry>? LogAppended;

    /// <summary>Adds a line to the activity log from any thread.</summary>
    public void AppendLog(LogEntry entry)
    {
        if (Dispatcher.UIThread.CheckAccess())
        {
            AppendLogCore(entry);
            return;
        }

        Dispatcher.UIThread.Post(() => AppendLogCore(entry));
    }

    /// <summary>Adds a line to the activity log.</summary>
    public void Log(LogLevel level, string message, string? filePath = null) =>
        AppendLog(new LogEntry(DateTimeOffset.Now, level, message, filePath));

    private void AppendLogCore(LogEntry entry)
    {
        LogEntries.Add(entry);

        // Bounded, so a long watcher session cannot grow the log without limit.
        while (LogEntries.Count > MaxLogEntries)
            LogEntries.RemoveAt(0);

        LogAppended?.Invoke(this, entry);
    }

    /// <summary>Applies a new theme and tells the shell to repaint.</summary>
    public void SetTheme(ThemeMode theme)
    {
        if (Settings.Theme == theme)
            return;

        Settings.Theme = theme;
        ThemeChanged?.Invoke(this, theme);
    }

    /// <summary>
    /// Writes the settings to disk, reporting failure to the log rather than
    /// throwing at the caller.
    /// </summary>
    public bool SaveSettings()
    {
        try
        {
            Settings.Save(SettingsPath);
            ConversionService.Settings = Settings;
            return true;
        }
        catch (Exception ex)
        {
            Log(LogLevel.Error, $"Could not save settings to '{SettingsPath}': {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Replaces every setting with the shipped defaults. The caller is expected to
    /// have confirmed this with the user first.
    /// </summary>
    public void ResetToDefaults()
    {
        Settings = DataWizardSettings.CreateDefault();
        ConversionService.Settings = Settings;
        Log(LogLevel.Info, "Settings reset to defaults.");
        ThemeChanged?.Invoke(this, Settings.Theme);
    }

    /// <summary>Replaces the settings wholesale, for example after an import.</summary>
    public void ReplaceSettings(DataWizardSettings settings)
    {
        Settings = settings;
        ConversionService.Settings = settings;
        ThemeChanged?.Invoke(this, settings.Theme);
    }
}
