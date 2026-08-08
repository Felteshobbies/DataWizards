using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DataWizard.App.Services;
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;
using DataWizard.Core.Csv;

namespace DataWizard.App.ViewModels;

/// <summary>An editable single-value row, for the date format list.</summary>
public partial class TextRow : ViewModelBase
{
    [ObservableProperty]
    public partial string Value { get; set; }

    public TextRow(string value) => Value = value;
}

/// <summary>
/// The Detection tab: every knob that decides how a CSV file is interpreted.
/// </summary>
/// <remarks>
/// Properties write straight through to the live
/// <see cref="DetectionSettings"/> instance so the effect is immediate, and the
/// test panel re-runs the analyser against a pasted sample after every change.
/// Being able to see the effect of a setting is the point: the previous version
/// hid all of this in an XML file that the interface never even loaded.
/// </remarks>
public partial class DetectionViewModel : ViewModelBase
{
    private readonly AppSession _session;
    private readonly DialogService _dialogs;
    private bool _suppressTest;

    [ObservableProperty]
    public partial string SampleText { get; set; }

    [ObservableProperty]
    public partial AnalysisPreview? TestResult { get; set; }

    [ObservableProperty]
    public partial string? TestError { get; set; }

    [ObservableProperty]
    public partial NamePatternRow? SelectedHeaderPattern { get; set; }

    [ObservableProperty]
    public partial FieldRuleRow? SelectedFieldRule { get; set; }

    [ObservableProperty]
    public partial TextRow? SelectedDateFormat { get; set; }

    public DetectionViewModel(AppSession session, DialogService dialogs)
    {
        _session = session;
        _dialogs = dialogs;

        SampleText = string.Join(Environment.NewLine,
            "customer_id;name;zip;price;order_date",
            "00123;Widget Ltd;01067;19,99;31.01.2025",
            "00456;Gadget GmbH;80331;4,50;01.02.2025",
            "00789;Sprocket AG;20095;123,00;15.02.2025");

        LoadFromSettings();
        RunTest();

        HeaderPatterns.CollectionChanged += OnRuleCollectionChanged;
        FieldRules.CollectionChanged += OnRuleCollectionChanged;
        DateFormats.CollectionChanged += OnRuleCollectionChanged;
    }

    private DetectionSettings Detection => _session.Settings.Detection;

    /// <summary>Header field name patterns.</summary>
    public ObservableCollection<NamePatternRow> HeaderPatterns { get; } = [];

    /// <summary>Data type rules.</summary>
    public ObservableCollection<FieldRuleRow> FieldRules { get; } = [];

    /// <summary>Accepted date formats.</summary>
    public ObservableCollection<TextRow> DateFormats { get; } = [];

    // ── Option lists for the drop-downs ─────────────────────────────────────

    public IReadOnlyList<string> EncodingOptions { get; } =
        new[] { string.Empty }.Concat(EncodingDetector.CommonEncodings()).ToArray();

    public IReadOnlyList<string> FallbackEncodingOptions { get; } = EncodingDetector.CommonEncodings();

    public IReadOnlyList<string> SeparatorOptions { get; } =
        new[] { SeparatorToken.Auto }.Concat(SeparatorToken.KnownTokens).ToArray();

    public IReadOnlyList<HeaderMode> HeaderModes { get; } = Enum.GetValues<HeaderMode>();

    public IReadOnlyList<NumberFormatMode> NumberFormatModes { get; } = Enum.GetValues<NumberFormatMode>();

    public IReadOnlyList<string> CultureOptions { get; } =
    [
        "de-DE", "en-US", "en-GB", "fr-FR", "it-IT", "es-ES", "nl-NL", "pl-PL", "tr-TR", "sv-SE"
    ];

    // ── Encoding ────────────────────────────────────────────────────────────

    /// <summary>Fixed encoding, empty meaning automatic detection.</summary>
    public string ForcedEncoding
    {
        get => Detection.ForcedEncoding;
        set => Set(v => Detection.ForcedEncoding = v ?? string.Empty, value, Detection.ForcedEncoding);
    }

    public bool TrustByteOrderMark
    {
        get => Detection.TrustByteOrderMark;
        set => Set(v => Detection.TrustByteOrderMark = v, value, Detection.TrustByteOrderMark);
    }

    public double MinEncodingConfidence
    {
        get => Detection.MinEncodingConfidence;
        set => Set(v => Detection.MinEncodingConfidence = v, value, Detection.MinEncodingConfidence);
    }

    public string FallbackEncoding
    {
        get => Detection.FallbackEncoding;
        set => Set(v => Detection.FallbackEncoding = v ?? "utf-8", value, Detection.FallbackEncoding);
    }

    // ── Separator ───────────────────────────────────────────────────────────

    public string CandidateSeparators
    {
        get => Detection.CandidateSeparators;
        set => Set(v => Detection.CandidateSeparators = v ?? ";,\\t|", value, Detection.CandidateSeparators);
    }

    public string ForcedSeparator
    {
        get => Detection.ForcedSeparator;
        set => Set(v => Detection.ForcedSeparator = v ?? SeparatorToken.Auto, value, Detection.ForcedSeparator);
    }

    public string QuoteCharacter
    {
        get => Detection.QuoteCharacter;
        set => Set(v => Detection.QuoteCharacter = string.IsNullOrEmpty(v) ? "\"" : v, value, Detection.QuoteCharacter);
    }

    public int AnalyzeLineCount
    {
        get => Detection.AnalyzeLineCount;
        set => Set(v => Detection.AnalyzeLineCount = Math.Clamp(v, 2, 100_000), value, Detection.AnalyzeLineCount);
    }

    public double MinSeparatorConfidence
    {
        get => Detection.MinSeparatorConfidence;
        set => Set(v => Detection.MinSeparatorConfidence = v, value, Detection.MinSeparatorConfidence);
    }

    // ── Header ──────────────────────────────────────────────────────────────

    public HeaderMode HeaderMode
    {
        get => Detection.HeaderMode;
        set => Set(v => Detection.HeaderMode = v, value, Detection.HeaderMode);
    }

    public int SkipLeadingLines
    {
        get => Detection.SkipLeadingLines;
        set => Set(v => Detection.SkipLeadingLines = Math.Max(0, v), value, Detection.SkipLeadingLines);
    }

    public bool SkipEmptyLines
    {
        get => Detection.SkipEmptyLines;
        set => Set(v => Detection.SkipEmptyLines = v, value, Detection.SkipEmptyLines);
    }

    public double HeaderScoreThreshold
    {
        get => Detection.HeaderScoreThreshold;
        set => Set(v => Detection.HeaderScoreThreshold = v, value, Detection.HeaderScoreThreshold);
    }

    public bool UseKnownFieldNames
    {
        get => Detection.UseKnownFieldNames;
        set => Set(v => Detection.UseKnownFieldNames = v, value, Detection.UseKnownFieldNames);
    }

    public double KnownFieldNameWeight
    {
        get => Detection.KnownFieldNameWeight;
        set => Set(v => Detection.KnownFieldNameWeight = v, value, Detection.KnownFieldNameWeight);
    }

    public bool UseTypeDivergence
    {
        get => Detection.UseTypeDivergence;
        set => Set(v => Detection.UseTypeDivergence = v, value, Detection.UseTypeDivergence);
    }

    public double TypeDivergenceWeight
    {
        get => Detection.TypeDivergenceWeight;
        set => Set(v => Detection.TypeDivergenceWeight = v, value, Detection.TypeDivergenceWeight);
    }

    public bool UseUniqueness
    {
        get => Detection.UseUniqueness;
        set => Set(v => Detection.UseUniqueness = v, value, Detection.UseUniqueness);
    }

    public double UniquenessWeight
    {
        get => Detection.UniquenessWeight;
        set => Set(v => Detection.UniquenessWeight = v, value, Detection.UniquenessWeight);
    }

    public bool UseNonEmptyRule
    {
        get => Detection.UseNonEmptyRule;
        set => Set(v => Detection.UseNonEmptyRule = v, value, Detection.UseNonEmptyRule);
    }

    public double NonEmptyWeight
    {
        get => Detection.NonEmptyWeight;
        set => Set(v => Detection.NonEmptyWeight = v, value, Detection.NonEmptyWeight);
    }

    // ── Value typing ────────────────────────────────────────────────────────

    public NumberFormatMode NumberFormat
    {
        get => Detection.NumberFormat;
        set => Set(v => Detection.NumberFormat = v, value, Detection.NumberFormat);
    }

    public string NumberCulture
    {
        get => Detection.NumberCulture;
        set => Set(v => Detection.NumberCulture = v ?? "de-DE", value, Detection.NumberCulture);
    }

    public bool AllowThousandsSeparator
    {
        get => Detection.AllowThousandsSeparator;
        set => Set(v => Detection.AllowThousandsSeparator = v, value, Detection.AllowThousandsSeparator);
    }

    public bool PreserveLeadingZeros
    {
        get => Detection.PreserveLeadingZeros;
        set => Set(v => Detection.PreserveLeadingZeros = v, value, Detection.PreserveLeadingZeros);
    }

    public int MaxIntegerDigits
    {
        get => Detection.MaxIntegerDigits;
        set => Set(v => Detection.MaxIntegerDigits = Math.Clamp(v, 1, 30), value, Detection.MaxIntegerDigits);
    }

    public bool DetectDates
    {
        get => Detection.DetectDates;
        set => Set(v => Detection.DetectDates = v, value, Detection.DetectDates);
    }

    public string DateCulture
    {
        get => Detection.DateCulture;
        set => Set(v => Detection.DateCulture = v ?? "de-DE", value, Detection.DateCulture);
    }

    public int MinDateYear
    {
        get => Detection.MinDateYear;
        set => Set(v => Detection.MinDateYear = v, value, Detection.MinDateYear);
    }

    public int MaxDateYear
    {
        get => Detection.MaxDateYear;
        set => Set(v => Detection.MaxDateYear = v, value, Detection.MaxDateYear);
    }

    public bool DetectBooleans
    {
        get => Detection.DetectBooleans;
        set => Set(v => Detection.DetectBooleans = v, value, Detection.DetectBooleans);
    }

    public string TrueValues
    {
        get => Detection.TrueValues;
        set => Set(v => Detection.TrueValues = v ?? string.Empty, value, Detection.TrueValues);
    }

    public string FalseValues
    {
        get => Detection.FalseValues;
        set => Set(v => Detection.FalseValues = v ?? string.Empty, value, Detection.FalseValues);
    }

    public bool QuotedFieldsAreText
    {
        get => Detection.QuotedFieldsAreText;
        set => Set(v => Detection.QuotedFieldsAreText = v, value, Detection.QuotedFieldsAreText);
    }

    public bool TrimWhitespace
    {
        get => Detection.TrimWhitespace;
        set => Set(v => Detection.TrimWhitespace = v, value, Detection.TrimWhitespace);
    }

    // ── Commands ────────────────────────────────────────────────────────────

    [RelayCommand]
    private void AddHeaderPattern()
    {
        var row = new NamePatternRow();
        HeaderPatterns.Add(row);
        SelectedHeaderPattern = row;
    }

    [RelayCommand]
    private void RemoveHeaderPattern()
    {
        if (SelectedHeaderPattern is not null)
            HeaderPatterns.Remove(SelectedHeaderPattern);
    }

    [RelayCommand]
    private void AddFieldRule()
    {
        var row = new FieldRuleRow();
        FieldRules.Add(row);
        SelectedFieldRule = row;
    }

    [RelayCommand]
    private void RemoveFieldRule()
    {
        if (SelectedFieldRule is not null)
            FieldRules.Remove(SelectedFieldRule);
    }

    [RelayCommand]
    private void AddDateFormat()
    {
        var row = new TextRow("dd.MM.yyyy");
        DateFormats.Add(row);
        SelectedDateFormat = row;
    }

    [RelayCommand]
    private void RemoveDateFormat()
    {
        if (SelectedDateFormat is not null)
            DateFormats.Remove(SelectedDateFormat);
    }

    [RelayCommand]
    private void RestoreDefaultPatterns()
    {
        _suppressTest = true;

        HeaderPatterns.Clear();
        foreach (var pattern in DefaultPatterns.CreateKnownFieldNames())
            HeaderPatterns.Add(new NamePatternRow(pattern));

        FieldRules.Clear();
        foreach (var rule in DefaultPatterns.CreateFieldRules())
            FieldRules.Add(new FieldRuleRow(rule));

        _suppressTest = false;
        _session.Log(LogLevel.Info, "Restored the built-in pattern lists.");
        RunTest();
    }

    /// <summary>
    /// Imports header patterns and type overrides from a pre-0.2
    /// <c>DataWizard.config.xml</c>.
    /// </summary>
    [RelayCommand]
    private async Task ImportLegacyConfigAsync()
    {
        var path = await _dialogs.PickOpenFileAsync(
            "Import a legacy DataWizard.config.xml", ".xml", "XML configuration");

        if (path is null)
            return;

        try
        {
            var result = LegacyConfigImporter.ImportFile(path);

            _suppressTest = true;

            if (result.HeaderPatterns.Count > 0)
            {
                HeaderPatterns.Clear();
                foreach (var pattern in result.HeaderPatterns)
                    HeaderPatterns.Add(new NamePatternRow(pattern));
            }

            if (result.FieldRules.Count > 0)
            {
                FieldRules.Clear();
                foreach (var rule in result.FieldRules)
                    FieldRules.Add(new FieldRuleRow(rule));
            }

            _suppressTest = false;

            _session.Log(LogLevel.Success,
                $"Imported {result.HeaderPatterns.Count} header pattern(s) and " +
                $"{result.FieldRules.Count} type rule(s) from {System.IO.Path.GetFileName(path)}.");

            foreach (var warning in result.Warnings)
                _session.Log(LogLevel.Warning, warning);

            RunTest();
        }
        catch (Exception ex)
        {
            _suppressTest = false;
            _session.Log(LogLevel.Error, $"Could not import '{path}': {ex.Message}");
        }
    }

    [RelayCommand]
    private void RunTest()
    {
        if (_suppressTest)
            return;

        SyncToSettings();
        TestError = null;
        TestResult = null;

        if (string.IsNullOrWhiteSpace(SampleText))
            return;

        try
        {
            var analysis = new CsvAnalyzer(Detection).AnalyzeText(SampleText, "(sample)");
            TestResult = new AnalysisPreview(analysis);
        }
        catch (Exception ex)
        {
            TestError = ex.Message;
        }
    }

    /// <summary>
    /// Copies the grid contents back into the settings object. Called before any
    /// analysis and before saving.
    /// </summary>
    public void SyncToSettings()
    {
        Detection.KnownFieldNames = HeaderPatterns
            .Where(r => !string.IsNullOrWhiteSpace(r.Pattern))
            .Select(r => r.ToModel())
            .ToList();

        Detection.FieldRules = FieldRules
            .Where(r => !string.IsNullOrWhiteSpace(r.Pattern) || r.ColumnIndex.HasValue)
            .Select(r => r.ToModel())
            .ToList();

        Detection.DateFormats = DateFormats
            .Select(r => r.Value)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .ToList();

        Detection.Normalize();
    }

    /// <summary>Rebuilds the grids from the settings, after a reset or import.</summary>
    public void LoadFromSettings()
    {
        _suppressTest = true;

        HeaderPatterns.Clear();
        foreach (var pattern in Detection.KnownFieldNames)
            HeaderPatterns.Add(new NamePatternRow(pattern));

        FieldRules.Clear();
        foreach (var rule in Detection.FieldRules)
            FieldRules.Add(new FieldRuleRow(rule));

        DateFormats.Clear();
        foreach (var format in Detection.DateFormats)
            DateFormats.Add(new TextRow(format));

        _suppressTest = false;

        OnPropertyChanged(string.Empty);
    }

    partial void OnSampleTextChanged(string value) => RunTest();

    private void OnRuleCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        // Newly added rows must re-run the test as they are filled in.
        if (e.NewItems is not null)
        {
            foreach (var item in e.NewItems.OfType<INotifyPropertyChanged>())
                item.PropertyChanged += OnRowChanged;
        }

        if (e.OldItems is not null)
        {
            foreach (var item in e.OldItems.OfType<INotifyPropertyChanged>())
                item.PropertyChanged -= OnRowChanged;
        }

        RunTest();
    }

    private void OnRowChanged(object? sender, PropertyChangedEventArgs e) => RunTest();

    /// <summary>
    /// Writes a value through to the settings and re-runs the test, but only when
    /// it actually changed.
    /// </summary>
    private void Set<T>(Action<T> apply, T value, T current, [System.Runtime.CompilerServices.CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(value, current))
            return;

        apply(value);
        OnPropertyChanged(propertyName);
        RunTest();
    }
}
