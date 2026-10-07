using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using DataWizard.App;
using DataWizard.App.Services;
using DataWizard.App.ViewModels;
using DataWizard.App.Views;
using DataWizard.Core.Configuration;

namespace DataWizard.App.Tests;

/// <summary>
/// The Diff tab packs two option cards and a result grid into one window. A Grid
/// without row definitions silently stacks every child into row 0, so these
/// tests pin down that the option rows really do stack and nothing spills out.
/// </summary>
[Collection(AvaloniaCollection.Name)]
public class DiffLayoutTests(AvaloniaFixture fixture)
{
    private readonly AvaloniaFixture _fixture = fixture;

    private static MainWindowViewModel CreateModel()
    {
        var settings = DataWizardSettings.CreateDefault();
        return new MainWindowViewModel(
            new AppSession(settings, Path.Combine(Path.GetTempPath(), "dw-layout-settings.json")),
            () => null);
    }

    [Theory]
    [InlineData(1180, 760)]
    [InlineData(900, 600)]
    public void TheDiffTabFitsWithoutOverlapping(int width, int height)
    {
        _fixture.Invoke(() =>
        {
            var window = new MainWindow
            {
                DataContext = CreateModel(),
                Width = width,
                Height = height
            };
            window.Show();

            var tabs = window.GetVisualDescendants().OfType<TabControl>().Single();
            tabs.SelectedIndex = 1;
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            var view = window.GetVisualDescendants().OfType<DiffView>().Single();
            var grid = view.GetVisualDescendants().OfType<Grid>().First();
            var input = grid.Children.OfType<StackPanel>().First();
            var cards = input.Children.OfType<Border>().ToList();
            var dataGrid = grid.Children.OfType<DataGrid>().Single();

            // Two and three stacked option rows, not overlapping ones.
            Assert.True(cards[0].Bounds.Height > 90,
                $"the table card is only {cards[0].Bounds.Height:F0}px tall; its rows overlap");
            Assert.True(cards[1].Bounds.Height > cards[0].Bounds.Height,
                $"the options card is only {cards[1].Bounds.Height:F0}px tall; its rows overlap");

            // Nothing may spill out of the tab, even in the smallest window.
            Assert.True(
                input.Bounds.Height + dataGrid.Bounds.Height + 20 <= grid.Bounds.Height + 1,
                $"input {input.Bounds.Height:F0} + grid {dataGrid.Bounds.Height:F0} " +
                $"does not fit into {grid.Bounds.Height:F0}");

            window.Close();
        });
    }
}
