using System.Globalization;
using DataWizard.Core.Configuration;

namespace DataWizard.Core.Csv;

/// <summary>The type a CSV value was recognised as, plus the converted value.</summary>
public readonly record struct TypedValue
{
    /// <summary>The recognised type.</summary>
    public FieldDataType Type { get; init; }

    /// <summary>The original text.</summary>
    public string Text { get; init; }

    /// <summary>Set for <see cref="FieldDataType.Integer"/> and <see cref="FieldDataType.Decimal"/>.</summary>
    public double Number { get; init; }

    /// <summary>Set for <see cref="FieldDataType.Date"/>.</summary>
    public DateTime Date { get; init; }

    /// <summary>Set for <see cref="FieldDataType.Boolean"/>.</summary>
    public bool Boolean { get; init; }

    /// <summary>Creates a text result.</summary>
    public static TypedValue AsText(string text) => new() { Type = FieldDataType.Text, Text = text };
}

/// <summary>
/// Decides what a CSV value actually is: text, a whole number, a decimal, a date
/// or a boolean.
/// </summary>
/// <remarks>
/// Everything this class does is driven by <see cref="DetectionSettings"/>, and
/// nothing consults the machine's regional settings. Parsing a date with the
/// ambient culture is how the same file ends up with different results on a
/// German and an American machine.
/// </remarks>
public sealed class ValueTypeDetector
{
    private readonly DetectionSettings _settings;
    private readonly CultureInfo _numberCulture;
    private readonly CultureInfo _dateCulture;
    private readonly string[] _dateFormats;
    private readonly HashSet<string> _trueValues;
    private readonly HashSet<string> _falseValues;

    public ValueTypeDetector(DetectionSettings settings)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _numberCulture = settings.ResolveNumberCulture();
        _dateCulture = settings.ResolveDateCulture();
        _dateFormats = [.. settings.DateFormats];
        _trueValues = new HashSet<string>(settings.ResolveTrueValues(), StringComparer.OrdinalIgnoreCase);
        _falseValues = new HashSet<string>(settings.ResolveFalseValues(), StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Determines the type of a field, taking the quoting flag into account.
    /// </summary>
    public TypedValue Detect(CsvField field)
    {
        if (_settings.QuotedFieldsAreText && field.WasQuoted)
            return TypedValue.AsText(field.Value);

        return Detect(field.Value);
    }

    /// <summary>Determines the type of a raw string.</summary>
    public TypedValue Detect(string? value)
    {
        var text = value ?? string.Empty;
        var trimmed = _settings.TrimWhitespace ? text.Trim() : text;

        if (trimmed.Length == 0)
            return TypedValue.AsText(text);

        if (_settings.DetectBooleans && TryParseBoolean(trimmed, out var boolean))
            return new TypedValue { Type = FieldDataType.Boolean, Text = text, Boolean = boolean };

        // Leading zeros carry meaning in article numbers, postal codes and
        // account numbers. Once the value becomes a number they are gone, so the
        // check happens before any numeric parse.
        if (_settings.PreserveLeadingZeros && HasSignificantLeadingZero(trimmed))
            return TypedValue.AsText(text);

        if (TryParseNumber(trimmed, out var number, out var isInteger))
        {
            if (isInteger)
            {
                // Excel stores numbers as doubles. Beyond ~15 digits the value
                // silently changes, which is what turns EAN codes into rounded
                // nonsense, so keep long digit strings as text.
                if (CountDigits(trimmed) > _settings.MaxIntegerDigits)
                    return TypedValue.AsText(text);

                return new TypedValue { Type = FieldDataType.Integer, Text = text, Number = number };
            }

            return new TypedValue { Type = FieldDataType.Decimal, Text = text, Number = number };
        }

        if (_settings.DetectDates && TryParseDate(trimmed, out var date))
            return new TypedValue { Type = FieldDataType.Date, Text = text, Date = date };

        return TypedValue.AsText(text);
    }

    /// <summary>
    /// Parses a numeric value according to the configured number format.
    /// </summary>
    /// <param name="value">The text to parse.</param>
    /// <param name="number">The parsed value.</param>
    /// <param name="isInteger">True when the value has no fractional part.</param>
    public bool TryParseNumber(string value, out double number, out bool isInteger)
    {
        number = 0;
        isInteger = false;

        if (value.Length == 0)
            return false;

        // A value that contains both a dot and a comma (for example
        // <c>1.234,56</c> or <c>1,234.56</c>) uses one of them for grouping and
        // the other for the decimal point. There is no other reading, so the
        // grouping separator is accepted even when the setting is off. A bare
        // <c>1.234</c> or <c>1,234</c> is genuinely ambiguous and stays governed
        // by the setting.
        var allowThousands = _settings.AllowThousandsSeparator
                             || (value.Contains('.') && value.Contains(','));

        var styles = NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint |
                     NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite;

        if (allowThousands)
            styles |= NumberStyles.AllowThousands;

        var parsed = _settings.NumberFormat switch
        {
            NumberFormatMode.Invariant =>
                double.TryParse(value, styles, CultureInfo.InvariantCulture, out number),

            NumberFormatMode.Culture =>
                double.TryParse(value, styles, _numberCulture, out number),

            _ => double.TryParse(value, styles, CultureInfo.InvariantCulture, out number)
                  || double.TryParse(value, styles, _numberCulture, out number)
        };

        if (!parsed || double.IsNaN(number) || double.IsInfinity(number))
            return false;

        isInteger = number == Math.Truncate(number) && !HasFractionalPart(value, allowThousands);
        return true;
    }

    /// <summary>
    /// Parses a date using only the configured formats.
    /// </summary>
    public bool TryParseDate(string value, out DateTime date)
    {
        date = default;

        if (value.Length == 0 || _dateFormats.Length == 0)
            return false;

        // A bare run of digits is a number, never a date - otherwise "20250131"
        // and even "6" start turning into dates.
        if (value.All(char.IsDigit))
            return false;

        const DateTimeStyles styles = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault;

        if (!DateTime.TryParseExact(value, _dateFormats, _dateCulture, styles, out date))
            return false;

        return date.Year >= _settings.MinDateYear && date.Year <= _settings.MaxDateYear;
    }

    /// <summary>
    /// Parses a date for a column that a rule has forced to
    /// <see cref="FieldDataType.Date"/>, optionally with a rule-specific format.
    /// </summary>
    public bool TryParseDate(string value, string? explicitFormat, out DateTime date)
    {
        date = default;

        if (string.IsNullOrWhiteSpace(explicitFormat))
            return TryParseDate(value, out date);

        const DateTimeStyles styles = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault;

        return DateTime.TryParseExact(value.Trim(), explicitFormat, _dateCulture, styles, out date)
               && date.Year >= _settings.MinDateYear
               && date.Year <= _settings.MaxDateYear;
    }

    /// <summary>Parses a boolean using the configured true and false words.</summary>
    public bool TryParseBoolean(string value, out bool result)
    {
        if (_trueValues.Contains(value))
        {
            result = true;
            return true;
        }

        if (_falseValues.Contains(value))
        {
            result = false;
            return true;
        }

        result = false;
        return false;
    }

    /// <summary>
    /// True for values such as <c>007</c> or <c>-0042</c>, where the zeros are
    /// part of the identifier. A single leading zero before a decimal separator
    /// (<c>0.5</c>) does not count.
    /// </summary>
    /// <remarks>
    /// Public because the diff engine needs the same rule: <c>00123</c> and
    /// <c>123</c> are different values, and comparing them as numbers would
    /// hide the difference.
    /// </remarks>
    public static bool HasSignificantLeadingZero(string value)
    {
        var index = 0;

        if (index < value.Length && (value[index] == '+' || value[index] == '-'))
            index++;

        if (index >= value.Length || value[index] != '0')
            return false;

        // "0", "0.5" and "0,5" are ordinary numbers.
        if (index + 1 >= value.Length)
            return false;

        var next = value[index + 1];
        if (!char.IsDigit(next))
            return false;

        // Every remaining character must be a digit for this to be an identifier
        // rather than something like "01.02.2025".
        for (var i = index; i < value.Length; i++)
        {
            if (!char.IsDigit(value[i]))
                return false;
        }

        return true;
    }

    /// <summary>
    /// Whether the text spells out decimal places, which decides between the
    /// integer and the decimal Excel format. Only the last separator matters, and
    /// a group of exactly three trailing digits is read as thousands grouping
    /// (<c>1,234</c>) rather than as a fraction (<c>1,23</c>).
    /// </summary>
    private static bool HasFractionalPart(string value, bool allowThousands)
    {
        var index = value.LastIndexOfAny(['.', ',']);
        if (index < 0)
            return false;

        var digitsAfter = value.Length - index - 1;
        if (digitsAfter == 0)
            return false;

        for (var i = index + 1; i < value.Length; i++)
        {
            if (!char.IsDigit(value[i]))
                return false;
        }

        return !(allowThousands && digitsAfter == 3);
    }

    private static int CountDigits(string value)
    {
        var count = 0;
        foreach (var c in value)
        {
            if (char.IsDigit(c))
                count++;
        }

        return count;
    }
}
