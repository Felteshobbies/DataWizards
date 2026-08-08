using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.Input;
using DataWizard.App.Services;
using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;

namespace DataWizard.App.ViewModels;

/// <summary>
/// The Output tab: how results are written in each direction.
/// </summary>
public partial class OutputViewModel : ViewModelBase
{
    private readonly AppSession _session;
    private readonly DialogService _dialogs;

    public OutputViewModel(AppSession session, DialogService dialogs)
    {
        _session = session;
        _dialogs = dialogs;
    }

    private ConversionSettings Conversion => _session.Settings.Conversion;

    // ── Option lists ────────────────────────────────────────────────────────

    public IReadOnlyList<string> SeparatorOptions { get; } = SeparatorToken.KnownTokens;

    public IReadOnlyList<string> EncodingOptions { get; } = EncodingDetector.CommonEncodings();

    public IReadOnlyList<QuoteMode> QuoteModes { get; } = Enum.GetValues<QuoteMode>();

    public IReadOnlyList<LineEndingStyle> LineEndings { get; } = Enum.GetValues<LineEndingStyle>();

    public IReadOnlyList<string> NumberCultureOptions { get; } =
        ["", "de-DE", "en-US", "en-GB", "fr-FR", "nl-NL"];

    // ── Shared ──────────────────────────────────────────────────────────────

    public string OutputFolder
    {
        get => Conversion.OutputFolder;
        set => Set(v => Conversion.OutputFolder = v ?? string.Empty, value, Conversion.OutputFolder);
    }

    public bool Overwrite
    {
        get => Conversion.Overwrite;
        set => Set(v => Conversion.Overwrite = v, value, Conversion.Overwrite);
    }

    // ── Excel to CSV ────────────────────────────────────────────────────────

    public string CsvSeparator
    {
        get => Conversion.CsvSeparator;
        set => Set(v => Conversion.CsvSeparator = v ?? "Semicolon", value, Conversion.CsvSeparator);
    }

    public string CsvEncoding
    {
        get => Conversion.CsvEncoding;
        set => Set(v => Conversion.CsvEncoding = v ?? "utf-8", value, Conversion.CsvEncoding);
    }

    public bool WriteByteOrderMark
    {
        get => Conversion.WriteByteOrderMark;
        set => Set(v => Conversion.WriteByteOrderMark = v, value, Conversion.WriteByteOrderMark);
    }

    public QuoteMode QuoteMode
    {
        get => Conversion.QuoteMode;
        set => Set(v => Conversion.QuoteMode = v, value, Conversion.QuoteMode);
    }

    public LineEndingStyle LineEnding
    {
        get => Conversion.LineEnding;
        set => Set(v => Conversion.LineEnding = v, value, Conversion.LineEnding);
    }

    public bool ExportAllSheets
    {
        get => Conversion.ExportAllSheets;
        set => Set(v => Conversion.ExportAllSheets = v, value, Conversion.ExportAllSheets);
    }

    public int WorksheetIndex
    {
        get => Conversion.WorksheetIndex;
        set => Set(v => Conversion.WorksheetIndex = Math.Max(0, v), value, Conversion.WorksheetIndex);
    }

    public string NumberOutputCulture
    {
        get => Conversion.NumberOutputCulture;
        set => Set(v => Conversion.NumberOutputCulture = v ?? string.Empty, value, Conversion.NumberOutputCulture);
    }

    public string DateOutputFormat
    {
        get => Conversion.DateOutputFormat;
        set => Set(v => Conversion.DateOutputFormat = v ?? "yyyy-MM-dd", value, Conversion.DateOutputFormat);
    }

    public int MaxDecimalPlaces
    {
        get => Conversion.MaxDecimalPlaces;
        set => Set(v => Conversion.MaxDecimalPlaces = Math.Clamp(v, 0, 15), value, Conversion.MaxDecimalPlaces);
    }

    public bool KeepEmptyRows
    {
        get => Conversion.KeepEmptyRows;
        set => Set(v => Conversion.KeepEmptyRows = v, value, Conversion.KeepEmptyRows);
    }

    // ── CSV to Excel ────────────────────────────────────────────────────────

    public string SheetName
    {
        get => Conversion.SheetName;
        set => Set(v => Conversion.SheetName = string.IsNullOrWhiteSpace(v) ? "Sheet1" : v, value, Conversion.SheetName);
    }

    public bool BoldHeaderRow
    {
        get => Conversion.BoldHeaderRow;
        set => Set(v => Conversion.BoldHeaderRow = v, value, Conversion.BoldHeaderRow);
    }

    public bool FreezeHeaderRow
    {
        get => Conversion.FreezeHeaderRow;
        set => Set(v => Conversion.FreezeHeaderRow = v, value, Conversion.FreezeHeaderRow);
    }

    public bool AutoFilterHeaderRow
    {
        get => Conversion.AutoFilterHeaderRow;
        set => Set(v => Conversion.AutoFilterHeaderRow = v, value, Conversion.AutoFilterHeaderRow);
    }

    public bool AutoFitColumns
    {
        get => Conversion.AutoFitColumns;
        set => Set(v => Conversion.AutoFitColumns = v, value, Conversion.AutoFitColumns);
    }

    public string DateNumberFormat
    {
        get => Conversion.DateNumberFormat;
        set => Set(v => Conversion.DateNumberFormat = v ?? "yyyy\\-mm\\-dd", value, Conversion.DateNumberFormat);
    }

    public string DecimalNumberFormat
    {
        get => Conversion.DecimalNumberFormat;
        set => Set(v => Conversion.DecimalNumberFormat = v ?? "#,##0.00", value, Conversion.DecimalNumberFormat);
    }

    public string IntegerNumberFormat
    {
        get => Conversion.IntegerNumberFormat;
        set => Set(v => Conversion.IntegerNumberFormat = v ?? "0", value, Conversion.IntegerNumberFormat);
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

    /// <summary>Re-reads every property, after a reset or import.</summary>
    public void Refresh() => OnPropertyChanged(string.Empty);

    private void Set<T>(Action<T> apply, T value, T current, [System.Runtime.CompilerServices.CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(value, current))
            return;

        apply(value);
        OnPropertyChanged(propertyName);
    }
}
