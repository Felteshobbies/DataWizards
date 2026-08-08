using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DataWizard.App.Services;
using DataWizard.Core.Conversion;
using DataWizard.Core.Csv;

namespace DataWizard.App.ViewModels;

/// <summary>
/// The Convert tab: collect files, inspect what analysis made of them, convert.
/// </summary>
public partial class ConvertViewModel : ViewModelBase
{
    private readonly AppSession _session;
    private readonly DialogService _dialogs;
    private CancellationTokenSource? _cancellation;

    [ObservableProperty]
    public partial bool IsBusy { get; set; }

    [ObservableProperty]
    public partial bool IsDragOver { get; set; }

    [ObservableProperty]
    public partial double ProgressValue { get; set; }

    [ObservableProperty]
    public partial string StatusText { get; set; }

    [ObservableProperty]
    public partial FileRow? SelectedFile { get; set; }

    [ObservableProperty]
    public partial AnalysisPreview? Preview { get; set; }

    [ObservableProperty]
    public partial string? PreviewError { get; set; }

    public ConvertViewModel(AppSession session, DialogService dialogs)
    {
        _session = session;
        _dialogs = dialogs;
        StatusText = "Drop files to get started.";

        Files.CollectionChanged += (_, _) =>
        {
            OnPropertyChanged(nameof(HasFiles));
            OnPropertyChanged(nameof(FileCountText));
            ConvertCommand.NotifyCanExecuteChanged();
            ClearCommand.NotifyCanExecuteChanged();
        };
    }

    /// <summary>The files waiting to be converted.</summary>
    public ObservableCollection<FileRow> Files { get; } = [];

    /// <summary>Whether anything is queued.</summary>
    public bool HasFiles => Files.Count > 0;

    /// <summary>A summary of the queue for the status line.</summary>
    public string FileCountText => Files.Count switch
    {
        0 => "No files selected",
        1 => "1 file selected",
        _ => $"{Files.Count} files selected"
    };

    /// <summary>The output folder shown on the tab, empty meaning "next to the source".</summary>
    public string OutputFolder
    {
        get => _session.Settings.Conversion.OutputFolder;
        set
        {
            if (_session.Settings.Conversion.OutputFolder == value)
                return;

            _session.Settings.Conversion.OutputFolder = value ?? string.Empty;
            OnPropertyChanged();
            OnPropertyChanged(nameof(OutputFolderDisplay));
        }
    }

    /// <summary>What the output folder box shows when nothing is set.</summary>
    public string OutputFolderDisplay =>
        string.IsNullOrWhiteSpace(OutputFolder) ? "(same folder as the source file)" : OutputFolder;

    /// <summary>Whether existing files may be replaced.</summary>
    public bool Overwrite
    {
        get => _session.Settings.Conversion.Overwrite;
        set
        {
            if (_session.Settings.Conversion.Overwrite == value)
                return;

            _session.Settings.Conversion.Overwrite = value;
            OnPropertyChanged();
        }
    }

    /// <summary>Adds files, ignoring duplicates and unsupported types.</summary>
    public void AddFiles(IEnumerable<string> paths)
    {
        var added = 0;
        var skipped = 0;

        foreach (var path in paths)
        {
            if (Directory.Exists(path))
            {
                // A dropped folder contributes the convertible files it contains.
                foreach (var file in EnumerateConvertibleFiles(path))
                {
                    if (TryAdd(file))
                        added++;
                }

                continue;
            }

            if (!File.Exists(path))
                continue;

            if (!ConversionService.CanConvert(path))
            {
                skipped++;
                continue;
            }

            if (TryAdd(path))
                added++;
        }

        if (added > 0)
            _session.Log(LogLevel.Info, $"Queued {added} file(s).");

        if (skipped > 0)
        {
            StatusText = $"Ignored {skipped} unsupported file(s). Supported: " +
                         string.Join(", ", ConversionService.SupportedExtensions);
        }
        else if (added > 0)
        {
            StatusText = FileCountText;
        }

        SelectedFile ??= Files.FirstOrDefault();
    }

    private static IEnumerable<string> EnumerateConvertibleFiles(string folder)
    {
        IEnumerable<string> files;

        try
        {
            files = Directory.EnumerateFiles(folder, "*", SearchOption.TopDirectoryOnly);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return [];
        }

        return files.Where(ConversionService.CanConvert);
    }

    private bool TryAdd(string path)
    {
        if (Files.Any(f => string.Equals(f.Path, path, StringComparison.OrdinalIgnoreCase)))
            return false;

        Files.Add(new FileRow(path));
        return true;
    }

    [RelayCommand]
    private async Task BrowseFilesAsync()
    {
        var files = await _dialogs.PickFilesAsync();
        if (files.Count > 0)
            AddFiles(files);
    }

    [RelayCommand]
    private async Task BrowseOutputFolderAsync()
    {
        var folder = await _dialogs.PickFolderAsync("Select the output folder", OutputFolder);
        if (folder is not null)
            OutputFolder = folder;
    }

    [RelayCommand]
    private void ClearOutputFolder() => OutputFolder = string.Empty;

    [RelayCommand(CanExecute = nameof(HasFiles))]
    private void Clear()
    {
        Files.Clear();
        SelectedFile = null;
        Preview = null;
        PreviewError = null;
        ProgressValue = 0;
        StatusText = "Drop files to get started.";
    }

    [RelayCommand]
    private void Remove(FileRow? row)
    {
        if (row is not null)
            Files.Remove(row);
    }

    [RelayCommand]
    private void OpenContainingFolder(FileRow? row)
    {
        var target = row?.Folder;

        if (string.IsNullOrWhiteSpace(target) || !Directory.Exists(target))
            return;

        try
        {
            Process.Start(new ProcessStartInfo { FileName = target, UseShellExecute = true });
        }
        catch (Exception ex)
        {
            _session.Log(LogLevel.Warning, $"Could not open '{target}': {ex.Message}");
        }
    }

    /// <summary>Whether a conversion can start.</summary>
    private bool CanConvert() => HasFiles && !IsBusy;

    [RelayCommand(CanExecute = nameof(CanConvert))]
    private async Task ConvertAsync()
    {
        IsBusy = true;
        ProgressValue = 0;
        _cancellation = new CancellationTokenSource();

        var queue = Files.ToList();
        var succeeded = 0;
        var failed = 0;

        try
        {
            for (var i = 0; i < queue.Count; i++)
            {
                if (_cancellation.IsCancellationRequested)
                    break;

                var row = queue[i];
                row.State = FileState.Running;
                row.Detail = null;
                StatusText = $"Converting {row.FileName} ({i + 1} of {queue.Count})";

                var result = await _session.ConversionService
                    .ConvertAsync(row.Path, null, null, _cancellation.Token)
                    .ConfigureAwait(true);

                if (result.Success)
                {
                    row.State = FileState.Done;
                    row.Detail = string.Join(", ", result.OutputPaths.Select(Path.GetFileName));
                    succeeded++;
                }
                else if (result.WasCancelled)
                {
                    row.State = FileState.Pending;
                    row.Detail = "Cancelled";
                }
                else
                {
                    row.State = FileState.Failed;
                    row.Detail = result.ErrorMessage;
                    failed++;
                }

                ProgressValue = (i + 1) / (double)queue.Count * 100d;
            }

            StatusText = _cancellation.IsCancellationRequested
                ? $"Cancelled. {succeeded} converted, {failed} failed."
                : $"Finished. {succeeded} converted, {failed} failed.";
        }
        finally
        {
            _cancellation.Dispose();
            _cancellation = null;

            // Re-enabling the button is what the previous version forgot, which
            // meant the application had to be restarted after one conversion.
            IsBusy = false;
        }
    }

    private bool CanCancel() => IsBusy;

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel()
    {
        _cancellation?.Cancel();
        StatusText = "Cancelling...";
    }

    partial void OnIsBusyChanged(bool value)
    {
        ConvertCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedFileChanged(FileRow? value) => _ = RefreshPreviewAsync(value);

    /// <summary>
    /// Analyses the selected file in the background so the preview never blocks the
    /// interface, which matters for files of any size.
    /// </summary>
    private async Task RefreshPreviewAsync(FileRow? row)
    {
        Preview = null;
        PreviewError = null;

        if (row is null)
            return;

        if (row.Direction == ConversionDirection.ExcelToCsv)
        {
            PreviewError = "Preview is available for delimited text files. " +
                           "Excel workbooks are read sheet by sheet during conversion.";
            return;
        }

        var path = row.Path;

        try
        {
            var analysis = await Task.Run(() => _session.ConversionService.Analyze(path));
            Preview = new AnalysisPreview(analysis);
        }
        catch (Exception ex)
        {
            PreviewError = ex.Message;
        }
    }
}
