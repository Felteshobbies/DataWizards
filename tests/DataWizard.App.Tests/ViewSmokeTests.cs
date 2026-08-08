using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using DataWizard.App;
using DataWizard.App.Services;
using DataWizard.App.ViewModels;
using DataWizard.App.Views;
using DataWizard.Core.Configuration;

namespace DataWizard.App.Tests;

/// <summary>
/// Loads every view against a real view model on a headless Avalonia session.
/// </summary>
/// <remarks>
/// XAML is only parsed when a control is first realised, and a tab that has never
/// been opened is never realised. A typo in a binding path or a missing resource
/// therefore compiles cleanly and only fails the moment a user clicks the tab.
/// These tests open all of them.
/// </remarks>
[Collection(AvaloniaCollection.Name)]
public class ViewSmokeTests(AvaloniaFixture fixture)
{
    private readonly AvaloniaFixture _fixture = fixture;

    private static AppSession CreateSession()
    {
        var settings = DataWizardSettings.CreateDefault();

        settings.Watcher.Rules.Add(new WatchRule
        {
            Name = "Sample rule",
            InputFolder = Path.GetTempPath()
        });

        return new AppSession(settings, Path.Combine(Path.GetTempPath(), "dw-test-settings.json"));
    }

    private MainWindowViewModel CreateViewModel() =>
        _fixture.Invoke(() => new MainWindowViewModel(CreateSession(), () => null));

    [Fact]
    public void MainWindowLoadsWithAllTabs()
    {
        _fixture.Invoke(() =>
        {
            var window = new MainWindow { DataContext = CreateViewModel() };
            window.Show();

            var tabs = window.GetVisualDescendants().OfType<TabControl>().Single();

            Assert.Equal(6, tabs.Items.Count);

            window.Close();
        });
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void EveryTabRealisesWithoutError(int tabIndex)
    {
        _fixture.Invoke(() =>
        {
            var window = new MainWindow { DataContext = CreateViewModel() };
            window.Show();

            var tabs = window.GetVisualDescendants().OfType<TabControl>().Single();
            tabs.SelectedIndex = tabIndex;

            // Force a layout pass so the newly selected tab's content is built.
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(tabIndex, tabs.SelectedIndex);

            window.Close();
        });
    }

    [Fact]
    public void ConvertViewBindsToItsViewModel()
    {
        _fixture.Invoke(() =>
        {
            var model = CreateViewModel();
            var view = new ConvertView { DataContext = model.Convert };

            var window = new Window { Content = view };
            window.Show();
            window.UpdateLayout();

            Assert.NotNull(view.FindControl<Border>("DropZone"));

            window.Close();
        });
    }

    [Fact]
    public void DetectionViewShowsTheLiveSampleResult()
    {
        _fixture.Invoke(() =>
        {
            var model = CreateViewModel();

            // The view model analyses its built-in sample on construction.
            Assert.NotNull(model.Detection.TestResult);
            Assert.Null(model.Detection.TestError);

            var view = new DetectionView { DataContext = model.Detection };
            var window = new Window { Content = view };

            window.Show();
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            window.Close();
        });
    }

    [Fact]
    public void ThemeVariantSwitchesBetweenLightAndDark()
    {
        _fixture.Invoke(() =>
        {
            var session = CreateSession();
            var model = new MainWindowViewModel(session, () => null);

            model.Theme = ThemeMode.Dark;
            Assert.Equal(ThemeMode.Dark, session.Settings.Theme);

            model.Theme = ThemeMode.Light;
            Assert.Equal(ThemeMode.Light, session.Settings.Theme);
        });
    }
}
