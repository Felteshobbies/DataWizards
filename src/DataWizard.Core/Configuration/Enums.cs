namespace DataWizard.Core.Configuration;

/// <summary>
/// How a pattern is compared against a field name.
/// </summary>
public enum MatchMode
{
    Contains,
    Exact,
    StartsWith,
    EndsWith,
    Regex
}

/// <summary>
/// The data type a CSV value is written as when it reaches Excel.
/// </summary>
public enum FieldDataType
{
    /// <summary>Detect the type from the value itself.</summary>
    Auto,
    Text,
    Integer,
    Decimal,
    Date,
    Boolean
}

/// <summary>
/// Whether the first content row is treated as a header.
/// </summary>
public enum HeaderMode
{
    /// <summary>Decide via the scoring heuristics in <see cref="DetectionSettings"/>.</summary>
    Auto,
    Always,
    Never
}

/// <summary>
/// Which culture is used to interpret numeric CSV values.
/// </summary>
public enum NumberFormatMode
{
    /// <summary>Try invariant (<c>1234.56</c>) first, then the configured culture (<c>1234,56</c>).</summary>
    Auto,
    /// <summary>Only accept <c>1234.56</c>.</summary>
    Invariant,
    /// <summary>Only accept the format of <see cref="DetectionSettings.NumberCulture"/>.</summary>
    Culture
}

/// <summary>
/// How aggressively fields are wrapped in quotes when writing CSV.
/// </summary>
public enum QuoteMode
{
    /// <summary>Quote only where the value would otherwise break the format.</summary>
    Minimal,
    /// <summary>Quote text fields, leave numbers and dates bare.</summary>
    TextFields,
    /// <summary>Quote every field.</summary>
    All
}

/// <summary>
/// Line terminator used when writing CSV.
/// </summary>
public enum LineEndingStyle
{
    /// <summary>Whatever the current platform uses.</summary>
    Platform,
    /// <summary>Windows (<c>\r\n</c>).</summary>
    Crlf,
    /// <summary>Unix (<c>\n</c>).</summary>
    Lf
}

/// <summary>
/// What happens to a source file once a watch rule has converted it.
/// </summary>
public enum SourceFileAction
{
    Keep,
    Move,
    Delete
}

/// <summary>
/// Application colour scheme.
/// </summary>
public enum ThemeMode
{
    System,
    Light,
    Dark
}
