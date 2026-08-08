using System.Text.RegularExpressions;

namespace DataWizard.Core.Configuration;

/// <summary>
/// A single reusable rule for recognising a field name. Used both for header
/// detection (does this row look like a header?) and for field rules (which
/// data type does this column get?).
/// </summary>
public class NamePattern
{
    /// <summary>The text or regular expression to look for.</summary>
    public string Pattern { get; set; } = string.Empty;

    /// <summary>How <see cref="Pattern"/> is compared against a field name.</summary>
    public MatchMode Match { get; set; } = MatchMode.Contains;

    /// <summary>When false (the default) comparison ignores casing.</summary>
    public bool CaseSensitive { get; set; }

    /// <summary>Lets a rule be switched off without deleting it.</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>Free-text note, shown in the settings UI only.</summary>
    public string? Comment { get; set; }

    private Regex? _compiled;
    private string? _compiledFor;
    private bool _compiledCaseSensitive;

    /// <summary>
    /// Tests a field name against this pattern. An invalid regular expression
    /// never matches rather than throwing, so one broken rule cannot break a
    /// whole conversion.
    /// </summary>
    public bool IsMatch(string? fieldName)
    {
        if (!Enabled || string.IsNullOrEmpty(Pattern) || fieldName is null)
            return false;

        var value = fieldName.Trim();
        if (value.Length == 0)
            return false;

        var comparison = CaseSensitive
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;

        return Match switch
        {
            MatchMode.Exact => value.Equals(Pattern, comparison),
            MatchMode.StartsWith => value.StartsWith(Pattern, comparison),
            MatchMode.EndsWith => value.EndsWith(Pattern, comparison),
            MatchMode.Contains => value.Contains(Pattern, comparison),
            MatchMode.Regex => RegexMatches(value),
            _ => false
        };
    }

    private bool RegexMatches(string value)
    {
        if (_compiled is null || _compiledFor != Pattern || _compiledCaseSensitive != CaseSensitive)
        {
            try
            {
                var options = RegexOptions.CultureInvariant;
                if (!CaseSensitive)
                    options |= RegexOptions.IgnoreCase;

                _compiled = new Regex(Pattern, options, TimeSpan.FromMilliseconds(250));
                _compiledFor = Pattern;
                _compiledCaseSensitive = CaseSensitive;
            }
            catch (ArgumentException)
            {
                // Invalid pattern - remember the failure so we do not retry per value.
                _compiled = null;
                _compiledFor = Pattern;
                _compiledCaseSensitive = CaseSensitive;
                return false;
            }
        }

        if (_compiled is null)
            return false;

        try
        {
            return _compiled.IsMatch(value);
        }
        catch (RegexMatchTimeoutException)
        {
            return false;
        }
    }

    /// <summary>Returns true when <see cref="Pattern"/> is a usable regular expression.</summary>
    public bool IsPatternValid(out string? error)
    {
        error = null;

        if (string.IsNullOrEmpty(Pattern))
        {
            error = "Pattern must not be empty.";
            return false;
        }

        if (Match != MatchMode.Regex)
            return true;

        try
        {
            _ = new Regex(Pattern);
            return true;
        }
        catch (ArgumentException ex)
        {
            error = ex.Message;
            return false;
        }
    }

    public override string ToString() => $"{Match}: {Pattern}";
}

/// <summary>
/// A <see cref="NamePattern"/> that forces a specific data type onto the
/// matching column, overriding automatic detection.
/// </summary>
public sealed class FieldRule : NamePattern
{
    /// <summary>The type every value in the matching column is written as.</summary>
    public FieldDataType DataType { get; set; } = FieldDataType.Text;

    /// <summary>
    /// When set, the rule targets this zero-based column position instead of
    /// matching on the field name. Useful for files without a header.
    /// </summary>
    public int? ColumnIndex { get; set; }

    /// <summary>
    /// Optional explicit date format (for example <c>dd.MM.yyyy</c>) applied when
    /// <see cref="DataType"/> is <see cref="FieldDataType.Date"/>.
    /// </summary>
    public string? DateFormat { get; set; }

    /// <summary>
    /// Tests whether this rule applies to a column, by position when
    /// <see cref="ColumnIndex"/> is set, otherwise by name.
    /// </summary>
    public bool AppliesTo(string? fieldName, int columnIndex)
    {
        if (!Enabled)
            return false;

        if (ColumnIndex.HasValue)
            return ColumnIndex.Value == columnIndex;

        return IsMatch(fieldName);
    }

    public override string ToString() =>
        ColumnIndex.HasValue
            ? $"column {ColumnIndex.Value} -> {DataType}"
            : $"{Match}: {Pattern} -> {DataType}";
}
