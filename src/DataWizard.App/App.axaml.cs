using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Styling;
using DataWizard.App.Services;
using DataWizard.App.ViewModels;
using DataWizard.App.Views;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;

namespace DataWizard.App;

public partial class App : Application
{
    private AppSession? _session;
    private MainWindowViewModel? _viewModel;

    public override void Initialize() => AvaloniaXamlLoader.Load(this);

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var settingsPath = DataWizardSettings.DefaultPath;
            var settings = DataWizardSettings.Load(settingsPath, out var warning);

            _session = new AppSession(settings, settingsPath);
            _session.ThemeChanged += (_, theme) => ApplyTheme(theme);

            if (warning is not null)
                _session.Log(LogLevel.Warning, warning);

            _session.Log(LogLevel.Info, $"DataWizards started. Settings: {settingsPath}");

            ApplyTheme(settings.Theme);

            var window = new MainWindow();
            _viewModel = new MainWindowViewModel(_session, () => window);
            window.DataContext = _viewModel;

            // Saving on close is what makes the settings actually stick; the
            // previous version had an empty saveSettings() stub.
            desktop.ShutdownRequested += OnShutdownRequested;

            desktop.MainWindow = window;

            window.Opened += (_, _) => _viewModel.Watcher.StartIfConfigured();
        }

        base.OnFrameworkInitializationCompleted();
    }

    private void OnShutdownRequested(object? sender, ShutdownRequestedEventArgs e)
    {
        if (_viewModel is null)
            return;

        try
        {
            _viewModel.ShutdownAsync().GetAwaiter().GetResult();
        }
        catch (Exception)
        {
            // Never block shutdown on a failed save.
        }
    }

    /// <summary>
    /// Applies the chosen appearance. <see cref="ThemeMode.System"/> maps to
    /// Avalonia's Default variant, which follows the operating system.
    /// </summary>
    private void ApplyTheme(ThemeMode theme) =>
        RequestedThemeVariant = theme switch
        {
            ThemeMode.Light => ThemeVariant.Light,
            ThemeMode.Dark => ThemeVariant.Dark,
            _ => ThemeVariant.Default
        };
}
