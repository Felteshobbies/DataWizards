using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using DataWizard.App.ViewModels;

namespace DataWizard.App.Views;

/// <summary>
/// The Convert tab. Drag and drop is wired here because it is a view concern; the
/// view model only ever receives a list of paths.
/// </summary>
public partial class ConvertView : UserControl
{
    public ConvertView()
    {
        InitializeComponent();

        AddHandler(DragDrop.DragEnterEvent, OnDragEnter);
        AddHandler(DragDrop.DragOverEvent, OnDragOver);
        AddHandler(DragDrop.DragLeaveEvent, OnDragLeave);
        AddHandler(DragDrop.DropEvent, OnDrop);
    }

    private ConvertViewModel? Model => DataContext as ConvertViewModel;

    private void OnDragEnter(object? sender, DragEventArgs e)
    {
        if (!e.DataTransfer.Contains(DataFormat.File))
        {
            e.DragEffects = DragDropEffects.None;
            return;
        }

        e.DragEffects = DragDropEffects.Copy;
        SetDropZoneActive(true);
    }

    private void OnDragOver(object? sender, DragEventArgs e) =>
        e.DragEffects = e.DataTransfer.Contains(DataFormat.File)
            ? DragDropEffects.Copy
            : DragDropEffects.None;

    private void OnDragLeave(object? sender, RoutedEventArgs e) => SetDropZoneActive(false);

    private void OnDrop(object? sender, DragEventArgs e)
    {
        SetDropZoneActive(false);

        var items = e.DataTransfer.TryGetFiles();
        if (items is null)
            return;

        var paths = items
            .Select(item => item.TryGetLocalPath())
            .Where(path => !string.IsNullOrEmpty(path))
            .Cast<string>()
            .ToList();

        if (paths.Count > 0)
            Model?.AddFiles(paths);
    }

    private async void OnDropZoneClicked(object? sender, PointerReleasedEventArgs e)
    {
        if (Model is { } model)
            await model.BrowseFilesCommand.ExecuteAsync(null);
    }

    /// <summary>Highlights the drop zone while a drag is over the tab.</summary>
    private void SetDropZoneActive(bool active)
    {
        if (this.FindControl<Border>("DropZone") is not { } zone)
            return;

        if (this.TryFindResource(active ? "DwDropZoneActive" : "DwDropZone", ActualThemeVariant, out var background))
            zone.Background = background as IBrush;

        if (this.TryFindResource(active ? "DwAccent" : "DwBorderStrong", ActualThemeVariant, out var border))
            zone.BorderBrush = border as IBrush;
    }
}
