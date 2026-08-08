namespace DataWizard.Core.Csv;

/// <summary>
/// One parsed CSV value together with the knowledge of whether the source wrote
/// it in quotes. That flag is what lets <c>"007"</c> survive as text.
/// </summary>
/// <param name="Value">The unescaped value, without the surrounding quotes.</param>
/// <param name="WasQuoted">True when the source wrapped the value in quotes.</param>
public readonly record struct CsvField(string Value, bool WasQuoted)
{
    /// <summary>An unquoted empty field.</summary>
    public static readonly CsvField Empty = new(string.Empty, false);

    /// <summary>True when the value contains nothing but whitespace.</summary>
    public bool IsEmpty => string.IsNullOrWhiteSpace(Value);

    public override string ToString() => WasQuoted ? $"\"{Value}\"" : Value;
}

/// <summary>
/// A complete CSV record. A record is not the same as a line: a quoted value may
/// contain line breaks and then spans several physical lines.
/// </summary>
public sealed class CsvRecord
{
    /// <summary>The values of this record.</summary>
    public required CsvField[] Fields { get; init; }

    /// <summary>One-based physical line on which the record starts.</summary>
    public required int StartLine { get; init; }

    /// <summary>How many physical lines the record spans.</summary>
    public required int LineCount { get; init; }

    /// <summary>True for a line that holds no value at all.</summary>
    public bool IsBlank => Fields.Length == 0 ||
                           (Fields.Length == 1 && !Fields[0].WasQuoted && Fields[0].Value.Length == 0);

    /// <summary>Number of values in this record.</summary>
    public int FieldCount => Fields.Length;

    /// <summary>Returns the value at <paramref name="index"/>, or an empty string when out of range.</summary>
    public string ValueAt(int index) =>
        index >= 0 && index < Fields.Length ? Fields[index].Value : string.Empty;

    public override string ToString() => string.Join(" | ", Fields.Select(f => f.Value));
}
