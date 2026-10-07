using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DataWizard.App.Services;
using DataWizard.Core.Conversion;
using DataWizard.Core.Diff;
using DataWizard.Core.Excel;

namespace DataWizard.App.ViewModels;

/// <summary>
/// The Diff tab: compare a reference table with a candidate table, show the
/// differences, and keep the comparison as a small JSON profile.
/// </summary>
public partial class DiffViewModel : ViewModelBase
{
    /// <summary>How many difference rows the grid shows at most.</summary>
    private const int MaxDisplayedEntries = 5000;

    private readonly AppSession _session;
    private readonly DialogService _dialogs;

    [ObservableProperty]
    public partial string ReferencePath { get; set; } = string.Empty;

    [ObservableProperty]
    public partial string CandidatePath { get; set; } = string.Empty;

    [ObservableProperty]
    public partial string? ProfilePath { get; set; }

    [ObservableProperty]
    public partial string ProfileName { get; set; } = string.Empty;

    [ObservableProperty]
    public partial string? ReferenceSheet { get; set; }

    [ObservableProperty]
    public partial string? CandidateSheet { get; set; }

    [ObservableProperty]
    public partial DiffMatchMode MatchMode { get; set; } = DiffMatchMode.Positional;

    [ObservableProperty]
    public partial string KeyColumns { get; set; } = string.Empty;

    [ObservableProperty]
    public partial string AbsoluteTolerance { get; set; } = string.Empty;

    [ObservableProperty]
    public partial string RelativeTolerance { get; set; } = string.Empty;

    [ObservableProperty]
    public partial string DateToleranceMinutes { get; set; } = string.Empty;

    [ObservableProperty]
    public partial bool IgnoreCase { get; set; }

    [ObservableProperty]
    public partial bool TrimValues { get; set; } = true;

    [ObservableProperty]
    public partial string IgnoredColumns { get; set; } = string.Empty;

    [ObservableProperty]
    public partial bool IsBusy { get; set; }

    [ObservableProperty]
    public partial string StatusText { get; set; }

    [ObservableProperty]
    public partial string? ErrorText { get; set; }

    /// <summary>Whether the error line shows.</summary>
    public bool HasError => ErrorText is not null;

    /// <summary>Whether a comparison report is shown.</summary>
    public bool HasReport => Report is not null;

    [ObservableProperty]
    public partial DiffReport? Report { get; set; }

    [ObservableProperty]
    public partial string SummaryText { get; set; } = string.Empty;

    [ObservableProperty]
    public partial bool IsTruncated { get; set; }

    public DiffViewModel(AppSession session, DialogService dialogs)
    {
        _session = session;
        _dialogs = dialogs;
        StatusText = "Pick a reference and a candidate table, then compare.";

        // The last profile is restored so the tab opens where the user left it.
        var last = session.Settings.Diff.LastProfilePath;

        if (!string.IsNullOrWhiteSpace(last) && File.Exists(last))
        {
            try
            {
                ApplyProfile(DiffProfile.Load(last), last);
                StatusText = $"Profile '{Path.GetFileName(last)}' restored.";
            }
            catch (Exception)
            {
                // An unreadable last profile must not break the tab.
            }
        }
    }

    /// <summary>Worksheets of the reference workbook, for the selector.</summary>
    public ObservableCollection<string> ReferenceSheets { get; } = [];

    /// <summary>Worksheets of the candidate workbook, for the selector.</summary>
    public ObservableCollection<string> CandidateSheets { get; } = [];

    /// <summary>The differences shown in the grid, capped at a usable size.</summary>
    public ObservableCollection<DiffEntry> Entries { get; } = [];

    /// <summary>The match modes for the selector.</summary>
    public IReadOnlyList<DiffMatchMode> MatchModes { get; } = Enum.GetValues<DiffMatchMode>();

    /// <summary>Whether the reference is a workbook, so the sheet selector shows.</summary>
    public bool HasReferenceSheets => ReferenceSheets.Count > 0;

    /// <summary>Whether the candidate is a workbook, so the sheet selector shows.</summary>
    public bool HasCandidateSheets => CandidateSheets.Count > 0;

    /// <summary>Whether the key columns box applies.</summary>
    public bool IsKeyed => MatchMode == DiffMatchMode.Keyed;

    /// <summary>What a comparison costs in words, for the status area.</summary>
    public string ProfileDisplay =>
        string.IsNullOrWhiteSpace(ProfilePath)
            ? "No profile. Save one to remember this comparison."
            : ProfilePath;

    partial void OnReferencePathChanged(string value)
    {
        RefreshReferenceSheets();
        CompareCommand.NotifyCanExecuteChanged();
    }

    partial void OnCandidatePathChanged(string value)
    {
        RefreshCandidateSheets();
        CompareCommand.NotifyCanExecuteChanged();
    }

    partial void OnMatchModeChanged(DiffMatchMode value) =>
        OnPropertyChanged(nameof(IsKeyed));

    partial void OnErrorTextChanged(string? value) =>
        OnPropertyChanged(nameof(HasError));

    partial void OnReportChanged(DiffReport? value) =>
        OnPropertyChanged(nameof(HasReport));

    partial void OnIsBusyChanged(bool value) =>
        CompareCommand.NotifyCanExecuteChanged();

    partial void OnProfilePathChanged(string? value) =>
        OnPropertyChanged(nameof(ProfileDisplay));

    private void RefreshReferenceSheets()
    {
        ReferenceSheets.Clear();
        ReferenceSheet = LoadSheets(ReferenceSheets, ReferencePath);
        OnPropertyChanged(nameof(HasReferenceSheets));
    }

    private void RefreshCandidateSheets()
    {
        CandidateSheets.Clear();
        CandidateSheet = LoadSheets(CandidateSheets, CandidatePath);
        OnPropertyChanged(nameof(HasCandidateSheets));
    }

    private static string? LoadSheets(ObservableCollection<string> target, string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)
            || !IsWorkbook(path))
        {
            return null;
        }

        try
        {
            foreach (var name in ExcelReader.ListSheetNames(path))
                target.Add(name);

            return target.FirstOrDefault();
        }
        catch (Exception)
        {
            // A workbook that cannot be listed is reported when the comparison runs.
            return null;
        }
    }

    private static bool IsWorkbook(string path) =>
        ConversionService.ExcelExtensions.Contains(
            Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    private bool CanCompare() =>
        !IsBusy && File.Exists(ReferencePath) && File.Exists(CandidatePath);

    [RelayCommand(CanExecute = nameof(CanCompare))]
    private async Task CompareAsync()
    {
        var options = BuildOptions();
        if (options is null)
            return;

        IsBusy = true;
        ErrorText = null;
        StatusText = "Comparing...";

        var referencePath = ReferencePath;
        var candidatePath = CandidatePath;
        var referenceSheet = ReferenceSheetIndex();
        var candidateSheet = CandidateSheetIndex();

        try
        {
            // The loader and engine are created per run so a settings reset
            // mid-session is picked up without restarting the app.
            var loader = new TableLoader(_session.Settings.Detection);
            var engine = new DiffEngine(_session.Settings.Detection);

            var report = await Task.Run(() =>
            {
                var reference = loader.Load(referencePath, referenceSheet);
                var candidate = loader.Load(candidatePath, candidateSheet);
                return engine.Compare(reference, candidate, options);
            }).ConfigureAwait(true);

            Report = report;
            FillEntries(report.Entries);
            SummaryText = BuildSummary(report);
            StatusText = $"Compared {report.ReferenceRowCount} and {report.CandidateRowCount} rows: " +
                         $"{report.DifferenceCount} difference(s).";

            _session.Log(LogLevel.Info,
                $"Diff: {report.DifferenceCount} difference(s) between " +
                $"{Path.GetFileName(referencePath)} and {Path.GetFileName(candidatePath)}.");

            AutoSaveProfile();
        }
        catch (Exception ex)
        {
            ErrorText = ex.Message;
            _session.Log(LogLevel.Error, $"Diff failed: {ex.Message}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    private int ReferenceSheetIndex() => SheetIndex(ReferenceSheets, ReferenceSheet);

    private int CandidateSheetIndex() => SheetIndex(CandidateSheets, CandidateSheet);

    private static int SheetIndex(ObservableCollection<string> sheets, string? selected)
    {
        var index = selected is null ? -1 : sheets.IndexOf(selected);
        return index < 0 ? 0 : index;
    }

    private void FillEntries(IReadOnlyList<DiffEntry> entries)
    {
        Entries.Clear();
        IsTruncated = entries.Count > MaxDisplayedEntries;

        foreach (var entry in entries.Take(MaxDisplayedEntries))
            Entries.Add(entry);
    }

    private static string BuildSummary(DiffReport report)
    {
        if (report.DifferenceCount == 0)
            return "No differences found.";

        var parts = new List<string>(5);

        if (report.ValueChangedCount > 0)
            parts.Add($"{report.ValueChangedCount} value change(s)");
        if (report.RowOnlyInReferenceCount > 0)
            parts.Add($"{report.RowOnlyInReferenceCount} only in reference");
        if (report.RowOnlyInCandidateCount > 0)
            parts.Add($"{report.RowOnlyInCandidateCount} only in candidate");
        if (report.ColumnOnlyInReferenceCount > 0)
            parts.Add($"{report.ColumnOnlyInReferenceCount} column(s) only in reference");
        if (report.ColumnOnlyInCandidateCount > 0)
            parts.Add($"{report.ColumnOnlyInCandidateCount} column(s) only in candidate");

        return $"{report.DifferenceCount} difference(s): " + string.Join(", ", parts);
    }

    [RelayCommand]
    private async Task BrowseReferenceAsync()
    {
        var path = await _dialogs.PickTableFileAsync("Select the reference table");
        if (path is not null)
            ReferencePath = path;
    }

    [RelayCommand]
    private async Task BrowseCandidateAsync()
    {
        var path = await _dialogs.PickTableFileAsync("Select the candidate table");
        if (path is not null)
            CandidatePath = path;
    }

    private bool CanExport() => Report is not null && !IsBusy;

    [RelayCommand(CanExecute = nameof(CanExport))]
    private async Task ExportReportAsync()
    {
        if (Report is null)
            return;

        var path = await _dialogs.PickSaveFileAsync("Export diff report", "diff-report", ".xlsx", "Excel workbook");
        if (path is null)
            return;

        try
        {
            DiffReportWriter.Write(Report, path);
            _session.Log(LogLevel.Info, $"Diff report exported to '{path}'.");
            StatusText = $"Report exported to {Path.GetFileName(path)}.";
        }
        catch (Exception ex)
        {
            ErrorText = $"Could not export the report: {ex.Message}";
        }
    }

    [RelayCommand]
    private async Task SaveProfileAsync()
    {
        var path = ProfilePath;

        if (string.IsNullOrWhiteSpace(path))
        {
            var name = string.IsNullOrWhiteSpace(ProfileName)
                ? "diff-profile.json"
                : SanitizeFileName(ProfileName.Trim()) + ".json";

            path = await _dialogs.PickSaveFileAsync("Save diff profile", name, ".json", "Diff profile");
            if (path is null)
                return;
        }

        try
        {
            CurrentProfile().Save(path);
            ProfilePath = path;
            _session.Settings.Diff.RememberProfile(path);
            _session.Log(LogLevel.Info, $"Diff profile saved to '{path}'.");
            StatusText = $"Profile saved to {Path.GetFileName(path)}.";
        }
        catch (Exception ex)
        {
            ErrorText = $"Could not save the profile: {ex.Message}";
        }
    }

    [RelayCommand]
    private async Task LoadProfileAsync()
    {
        var path = await _dialogs.PickOpenFileAsync("Open diff profile", ".json", "Diff profile");
        if (path is null)
            return;

        try
        {
            ApplyProfile(DiffProfile.Load(path), path);
            _session.Settings.Diff.RememberProfile(path);
            StatusText = $"Profile loaded from {Path.GetFileName(path)}.";
        }
        catch (Exception ex)
        {
            ErrorText = $"Could not load the profile: {ex.Message}";
        }
    }

    /// <summary>
    /// Copies a profile into the tab, for loading from disk and for restoring
    /// the last profile on start.
    /// </summary>
    public void ApplyProfile(DiffProfile profile, string? path = null)
    {
        ProfilePath = path;
        ProfileName = profile.Name ?? string.Empty;
        ReferencePath = profile.ReferencePath ?? string.Empty;
        CandidatePath = profile.CandidatePath ?? string.Empty;
        ApplyOptions(profile.Options);
    }

    private void ApplyOptions(DiffOptions options)
    {
        MatchMode = options.MatchMode;
        KeyColumns = string.Join(", ", options.KeyColumns);
        AbsoluteTolerance = FormatNumber(options.AbsoluteTolerance);
        RelativeTolerance = FormatNumber(options.RelativeTolerance);
        DateToleranceMinutes = options.DateTolerance == TimeSpan.Zero
            ? string.Empty
            : FormatNumber(options.DateTolerance.TotalMinutes);
        IgnoreCase = options.IgnoreCase;
        TrimValues = options.TrimValues;
        IgnoredColumns = string.Join(", ", options.IgnoredColumns);
    }

    /// <summary>The profile the tab currently holds, built from the visible fields.</summary>
    public DiffProfile CurrentProfile()
    {
        var options = BuildOptions() ?? new DiffOptions
        {
            MatchMode = MatchMode,
            KeyColumns = SplitList(KeyColumns),
            IgnoreCase = IgnoreCase,
            TrimValues = TrimValues,
            IgnoredColumns = SplitList(IgnoredColumns)
        };

        return new DiffProfile
        {
            Name = string.IsNullOrWhiteSpace(ProfileName) ? null : ProfileName.Trim(),
            ReferencePath = string.IsNullOrWhiteSpace(ReferencePath) ? null : ReferencePath,
            CandidatePath = string.IsNullOrWhiteSpace(CandidatePath) ? null : CandidatePath,
            Options = options
        };
    }

    /// <summary>
    /// Re-saves the active profile after a comparison, so the file always holds
    /// what the tab last did.
    /// </summary>
    private void AutoSaveProfile()
    {
        if (string.IsNullOrWhiteSpace(ProfilePath))
            return;

        try
        {
            CurrentProfile().Save(ProfilePath);
            _session.Settings.Diff.RememberProfile(ProfilePath);
        }
        catch (Exception ex)
        {
            _session.Log(LogLevel.Warning, $"Could not save diff profile '{ProfilePath}': {ex.Message}");
        }
    }

    /// <summary>
    /// Turns the option fields into a <see cref="DiffOptions"/>, reporting
    /// invalid input in <see cref="ErrorText"/> and returning null.
    /// </summary>
    private DiffOptions? BuildOptions()
    {
        if (!TryParseTolerance(AbsoluteTolerance, out var absolute))
            return null;

        if (!TryParseTolerance(RelativeTolerance, out var relative))
            return null;

        if (!TryParseTolerance(DateToleranceMinutes, out var dateMinutes))
            return null;

        var options = new DiffOptions
        {
            MatchMode = MatchMode,
            KeyColumns = SplitList(KeyColumns),
            AbsoluteTolerance = absolute,
            RelativeTolerance = relative,
            DateTolerance = TimeSpan.FromMinutes(dateMinutes),
            IgnoreCase = IgnoreCase,
            TrimValues = TrimValues,
            IgnoredColumns = SplitList(IgnoredColumns)
        };

        if (options.MatchMode == DiffMatchMode.Keyed && options.KeyColumns.Length == 0)
        {
            ErrorText = "Keyed comparison needs at least one key column.";
            return null;
        }

        return options;
    }

    private bool TryParseTolerance(string text, out double value)
    {
        value = 0d;

        if (string.IsNullOrWhiteSpace(text))
            return true;

        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
            && !double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
        {
            ErrorText = $"'{text}' is not a number.";
            return false;
        }

        if (value < 0)
        {
            ErrorText = "Tolerances must be zero or positive.";
            return false;
        }

        return true;
    }

    private static string FormatNumber(double value) =>
        value == 0d ? string.Empty : value.ToString("R", CultureInfo.InvariantCulture);

    private static string[] SplitList(string text) =>
        text.Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(part => part.Length > 0)
            .ToArray();

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
        return cleaned.Length > 0 ? cleaned : "diff-profile";
    }
}
