using System.Globalization;

namespace DataWizard.Core.Configuration;

/// <summary>
/// Everything that governs how a CSV file is interpreted: which encoding and
/// separator it uses, whether the first row is a header, and what data type
/// each value becomes.
/// </summary>
public sealed class DetectionSettings
{
    // ── Encoding ────────────────────────────────────────────────────────────

    /// <summary>
    /// Web name of a fixed encoding (for example <c>utf-8</c> or
    /// <c>windows-1252</c>). Empty means detect automatically.
    /// </summary>
    public string ForcedEncoding { get; set; } = string.Empty;

    /// <summary>A byte order mark, when present, always wins over detection.</summary>
    public bool TrustByteOrderMark { get; set; } = true;

    /// <summary>
    /// Detector confidence (0-1) below which <see cref="FallbackEncoding"/> is
    /// used instead of the detected result.
    /// </summary>
    public double MinEncodingConfidence { get; set; } = 0.5;

    /// <summary>Encoding used when detection fails or is not confident enough.</summary>
    public string FallbackEncoding { get; set; } = "utf-8";

    // ── Separator ───────────────────────────────────────────────────────────

    /// <summary>Characters considered during separator detection.</summary>
    public string CandidateSeparators { get; set; } = ";,\\t|";

    /// <summary>A fixed separator token, or <c>Auto</c> to detect it.</summary>
    public string ForcedSeparator { get; set; } = SeparatorToken.Auto;

    /// <summary>The character that begins and ends a quoted field.</summary>
    public string QuoteCharacter { get; set; } = "\"";

    /// <summary>How many lines are read to analyse structure.</summary>
    public int AnalyzeLineCount { get; set; } = 200;

    /// <summary>
    /// Fraction (0-1) of analysed lines that must agree on a field count before
    /// a separator is accepted. Below this the analysis is flagged as uncertain.
    /// </summary>
    public double MinSeparatorConfidence { get; set; } = 0.6;

    // ── Header ──────────────────────────────────────────────────────────────

    /// <summary>Whether the first content row is treated as a header.</summary>
    public HeaderMode HeaderMode { get; set; } = HeaderMode.Auto;

    /// <summary>Lines skipped before analysis begins, for files with a preamble.</summary>
    public int SkipLeadingLines { get; set; }

    /// <summary>Blank lines before the first content row are ignored.</summary>
    public bool SkipEmptyLines { get; set; } = true;

    /// <summary>
    /// Combined score (0-1) the first row must reach to count as a header when
    /// <see cref="HeaderMode"/> is <see cref="HeaderMode.Auto"/>.
    /// </summary>
    public double HeaderScoreThreshold { get; set; } = 0.5;

    /// <summary>Score the first row for matches against <see cref="KnownFieldNames"/>.</summary>
    public bool UseKnownFieldNames { get; set; } = true;

    /// <summary>Relative influence of the known-field-name signal.</summary>
    public double KnownFieldNameWeight { get; set; } = 1.0;

    /// <summary>
    /// Score the first row for being text where the rows below it are numeric or
    /// date values. This is the strongest signal for machine-generated exports.
    /// </summary>
    public bool UseTypeDivergence { get; set; } = true;

    /// <summary>Relative influence of the type-divergence signal.</summary>
    public double TypeDivergenceWeight { get; set; } = 1.5;

    /// <summary>Score the first row for having distinct values in every column.</summary>
    public bool UseUniqueness { get; set; } = true;

    /// <summary>Relative influence of the uniqueness signal.</summary>
    public double UniquenessWeight { get; set; } = 0.5;

    /// <summary>Score the first row for having no empty cells.</summary>
    public bool UseNonEmptyRule { get; set; } = true;

    /// <summary>Relative influence of the non-empty signal.</summary>
    public double NonEmptyWeight { get; set; } = 0.5;

    /// <summary>Field names that mark a row as a header when they appear in it.</summary>
    public List<NamePattern> KnownFieldNames { get; set; } = [];

    // ── Value typing ────────────────────────────────────────────────────────

    /// <summary>Which culture is accepted for numeric values.</summary>
    public NumberFormatMode NumberFormat { get; set; } = NumberFormatMode.Auto;

    /// <summary>Culture used for <see cref="NumberFormatMode.Culture"/> and as the second try in Auto mode.</summary>
    public string NumberCulture { get; set; } = "de-DE";

    /// <summary>
    /// Accept grouping separators such as <c>1,234.56</c>. Off by default because
    /// it makes <c>1.234</c> ambiguous between one thousand and a decimal.
    /// </summary>
    public bool AllowThousandsSeparator { get; set; }

    /// <summary>
    /// Keep values such as <c>00123</c> as text. Without this the leading zeros
    /// are lost the moment the value becomes a number.
    /// </summary>
    public bool PreserveLeadingZeros { get; set; } = true;

    /// <summary>
    /// Digit count above which a whole number is stored as text. Excel stores
    /// numbers as doubles, so anything longer silently loses precision - which is
    /// what mangles long article numbers, IBANs and EAN codes.
    /// </summary>
    public int MaxIntegerDigits { get; set; } = 15;

    /// <summary>Detect date values at all.</summary>
    public bool DetectDates { get; set; } = true;

    /// <summary>
    /// Formats accepted for dates, tried in order. Parsing is restricted to this
    /// list so that results never depend on the machine's regional settings.
    /// </summary>
    public List<string> DateFormats { get; set; } =
    [
        "yyyy-MM-dd",
        "dd.MM.yyyy",
        "dd/MM/yyyy",
        "yyyy/MM/dd",
        "yyyy-MM-dd HH:mm:ss",
        "dd.MM.yyyy HH:mm:ss",
        "dd.MM.yy"
    ];

    /// <summary>Culture used to resolve month names in <see cref="DateFormats"/>.</summary>
    public string DateCulture { get; set; } = "de-DE";

    /// <summary>Dates outside this range are treated as text rather than dates.</summary>
    public int MinDateYear { get; set; } = 1900;

    /// <inheritdoc cref="MinDateYear"/>
    public int MaxDateYear { get; set; } = 2100;

    /// <summary>Detect boolean values.</summary>
    public bool DetectBooleans { get; set; }

    /// <summary>Semicolon-separated list of values meaning true.</summary>
    public string TrueValues { get; set; } = "true;yes;ja;y;j";

    /// <summary>Semicolon-separated list of values meaning false.</summary>
    public string FalseValues { get; set; } = "false;no;nein;n";

    /// <summary>
    /// A field that was quoted in the source is always text. This is the usual
    /// way of protecting values such as <c>"007"</c>.
    /// </summary>
    public bool QuotedFieldsAreText { get; set; } = true;

    /// <summary>Strip surrounding whitespace from unquoted values.</summary>
    public bool TrimWhitespace { get; set; } = true;

    /// <summary>Per-column type overrides, evaluated in order; the first match wins.</summary>
    public List<FieldRule> FieldRules { get; set; } = [];

    // ── Helpers ─────────────────────────────────────────────────────────────

    /// <summary>The configured quote character, falling back to <c>"</c>.</summary>
    public char QuoteChar =>
        string.IsNullOrEmpty(QuoteCharacter) ? '"' : QuoteCharacter[0];

    /// <summary>The fixed separator, or <c>null</c> when it should be detected.</summary>
    public char? FixedSeparator => SeparatorToken.Parse(ForcedSeparator);

    /// <summary>The candidate separators as distinct characters.</summary>
    public char[] SeparatorCandidates => SeparatorToken.ParseCandidates(CandidateSeparators);

    /// <summary>The culture used for numeric parsing, falling back to invariant.</summary>
    public CultureInfo ResolveNumberCulture() => ResolveCulture(NumberCulture);

    /// <summary>The culture used for date parsing, falling back to invariant.</summary>
    public CultureInfo ResolveDateCulture() => ResolveCulture(DateCulture);

    private static CultureInfo ResolveCulture(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return CultureInfo.InvariantCulture;

        try
        {
            return CultureInfo.GetCultureInfo(name);
        }
        catch (CultureNotFoundException)
        {
            return CultureInfo.InvariantCulture;
        }
    }

    /// <summary>Splits <see cref="TrueValues"/> into individual tokens.</summary>
    public IReadOnlyList<string> ResolveTrueValues() => SplitList(TrueValues);

    /// <summary>Splits <see cref="FalseValues"/> into individual tokens.</summary>
    public IReadOnlyList<string> ResolveFalseValues() => SplitList(FalseValues);

    private static string[] SplitList(string? value) =>
        string.IsNullOrWhiteSpace(value)
            ? []
            : value.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    /// <summary>
    /// Clamps values that would otherwise break analysis, so a hand-edited
    /// settings file cannot put the engine into an impossible state.
    /// </summary>
    public void Normalize()
    {
        AnalyzeLineCount = Math.Clamp(AnalyzeLineCount, 2, 100_000);
        MinSeparatorConfidence = Math.Clamp(MinSeparatorConfidence, 0d, 1d);
        MinEncodingConfidence = Math.Clamp(MinEncodingConfidence, 0d, 1d);
        HeaderScoreThreshold = Math.Clamp(HeaderScoreThreshold, 0d, 1d);
        SkipLeadingLines = Math.Max(0, SkipLeadingLines);
        MaxIntegerDigits = Math.Clamp(MaxIntegerDigits, 1, 30);

        KnownFieldNameWeight = Math.Max(0, KnownFieldNameWeight);
        TypeDivergenceWeight = Math.Max(0, TypeDivergenceWeight);
        UniquenessWeight = Math.Max(0, UniquenessWeight);
        NonEmptyWeight = Math.Max(0, NonEmptyWeight);

        if (MinDateYear > MaxDateYear)
            (MinDateYear, MaxDateYear) = (MaxDateYear, MinDateYear);

        if (string.IsNullOrEmpty(QuoteCharacter))
            QuoteCharacter = "\"";

        if (DateFormats.Count == 0)
            DateFormats.Add("yyyy-MM-dd");

        KnownFieldNames ??= [];
        FieldRules ??= [];
    }

    /// <summary>Creates a copy, so the UI can edit settings without touching the live instance.</summary>
    public DetectionSettings Clone()
    {
        var clone = (DetectionSettings)MemberwiseClone();
        clone.DateFormats = [.. DateFormats];
        clone.KnownFieldNames = KnownFieldNames
            .Select(p => new NamePattern
            {
                Pattern = p.Pattern,
                Match = p.Match,
                CaseSensitive = p.CaseSensitive,
                Enabled = p.Enabled,
                Comment = p.Comment
            })
            .ToList();
        clone.FieldRules = FieldRules
            .Select(r => new FieldRule
            {
                Pattern = r.Pattern,
                Match = r.Match,
                CaseSensitive = r.CaseSensitive,
                Enabled = r.Enabled,
                Comment = r.Comment,
                DataType = r.DataType,
                ColumnIndex = r.ColumnIndex,
                DateFormat = r.DateFormat
            })
            .ToList();
        return clone;
    }
}
