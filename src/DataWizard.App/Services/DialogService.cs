using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using DataWizard.Core.Conversion;

namespace DataWizard.App.Services;

/// <summary>
/// File and folder pickers, wrapped so view models never touch a window.
/// </summary>
public sealed class DialogService
{
    private readonly Func<TopLevel?> _topLevel;

    public DialogService(Func<TopLevel?> topLevel)
    {
        _topLevel = topLevel;
    }

    /// <summary>Asks for one or more files to convert.</summary>
    public async Task<IReadOnlyList<string>> PickFilesAsync()
    {
        var topLevel = _topLevel();
        if (topLevel is null)
            return [];

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select files to convert",
            AllowMultiple = true,
            FileTypeFilter =
            [
                new FilePickerFileType("All supported")
                {
                    Patterns = ConversionService.SupportedExtensions.Select(e => "*" + e).ToArray()
                },
                new FilePickerFileType("Delimited text")
                {
                    Patterns = ConversionService.CsvExtensions.Select(e => "*" + e).ToArray()
                },
                new FilePickerFileType("Excel workbooks")
                {
                    Patterns = ConversionService.ExcelExtensions.Select(e => "*" + e).ToArray()
                },
                FilePickerFileTypes.All
            ]
        });

        return files.Select(f => f.TryGetLocalPath()).Where(p => p is not null).Cast<string>().ToArray();
    }

    /// <summary>Asks for a folder.</summary>
    public async Task<string?> PickFolderAsync(string title, string? startAt = null)
    {
        var topLevel = _topLevel();
        if (topLevel is null)
            return null;

        var options = new FolderPickerOpenOptions
        {
            Title = title,
            AllowMultiple = false
        };

        if (!string.IsNullOrWhiteSpace(startAt))
        {
            try
            {
                options.SuggestedStartLocation =
                    await topLevel.StorageProvider.TryGetFolderFromPathAsync(startAt);
            }
            catch (Exception ex) when (ex is ArgumentException or UnauthorizedAccessException)
            {
                // An unreachable start folder is not worth failing the picker over.
            }
        }

        var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(options);
        return folders.Count > 0 ? folders[0].TryGetLocalPath() : null;
    }

    /// <summary>Asks for a file to read settings or a legacy configuration from.</summary>
    public async Task<string?> PickOpenFileAsync(string title, string extension, string description)
    {
        var topLevel = _topLevel();
        if (topLevel is null)
            return null;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = title,
            AllowMultiple = false,
            FileTypeFilter =
            [
                new FilePickerFileType(description) { Patterns = ["*" + extension] },
                FilePickerFileTypes.All
            ]
        });

        return files.Count > 0 ? files[0].TryGetLocalPath() : null;
    }

    /// <summary>Asks where to write a file.</summary>
    public async Task<string?> PickSaveFileAsync(string title, string suggestedName, string extension, string description)
    {
        var topLevel = _topLevel();
        if (topLevel is null)
            return null;

        var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = title,
            SuggestedFileName = suggestedName,
            DefaultExtension = extension.TrimStart('.'),
            FileTypeChoices =
            [
                new FilePickerFileType(description) { Patterns = ["*" + extension] }
            ]
        });

        return file?.TryGetLocalPath();
    }
}
