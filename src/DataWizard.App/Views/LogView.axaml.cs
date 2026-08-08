using System.Collections.Specialized;
using Avalonia.Controls;
using Avalonia.Threading;
using DataWizard.App.ViewModels;

namespace DataWizard.App.Views;

/// <summary>
/// The Log tab. Following the newest entry is a scroll concern, so it is handled
/// here rather than in the view model.
/// </summary>
public partial class LogView : UserControl
{
    private INotifyCollectionChanged? _observed;

    public LogView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => Attach();
    }

    private void Attach()
    {
        if (_observed is not null)
            _observed.CollectionChanged -= OnRowsChanged;

        _observed = (DataContext as LogViewModel)?.Rows;

        if (_observed is not null)
            _observed.CollectionChanged += OnRowsChanged;
    }

    private void OnRowsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (DataContext is not LogViewModel { AutoScroll: true })
            return;

        if (e.Action != NotifyCollectionChangedAction.Add)
            return;

        // Let the panel measure the new row before scrolling to it.
        Dispatcher.UIThread.Post(
            () => this.FindControl<ScrollViewer>("LogScroller")?.ScrollToEnd(),
            DispatcherPriority.Background);
    }
}
