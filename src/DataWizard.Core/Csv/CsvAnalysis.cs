using System.Text;
using DataWizard.Core.Configuration;

namespace DataWizard.Core.Csv;

/// <summary>What analysis worked out about one column.</summary>
public sealed class ColumnInfo
{
    /// <summary>Zero-based position of the column.</summary>
    public required int Index { get; init; }

    /// <summary>Header name, or a generated <c>Column n</c>.</summary>
    public required string Name { get; init; }

    /// <summary>The type that dominates the sampled values.</summary>
    public required FieldDataType DetectedType { get; init; }

    /// <summary>The type actually used, after field rules have been applied.</summary>
    public required FieldDataType EffectiveType { get; set; }

    /// <summary>The rule that overrode the detected type, when one did.</summary>
    public FieldRule? AppliedRule { get; set; }

    /// <summary>How many non-empty values were sampled.</summary>
    public required int SampleCount { get; init; }

    /// <summary>How many sampled values were empty.</summary>
    public required int EmptyCount { get; init; }

    /// <summary>Distribution of detected types across the sample.</summary>
    public required IReadOnlyDictionary<FieldDataType, int> TypeCounts { get; init; }

    /// <summary>True when a rule changed the type away from what was detected.</summary>
    public bool WasOverridden => AppliedRule is not null && EffectiveType != DetectedType;

    /// <summary>
    /// True when the sample contained more than one type, which usually means the
    /// column is inconsistent and worth a look.
    /// </summary>
    public bool IsMixed => TypeCounts.Count(kvp => kvp.Value > 0) > 1;

    /// <summary>A short summary for the analysis panel.</summary>
    public string Describe()
    {
        var text = $"{Name}: {EffectiveType}";

        if (WasOverridden)
            text += $" (rule '{AppliedRule!.Pattern}' overrides {DetectedType})";
        else if (IsMixed)
            text += " (mixed values)";

        return text;
    }
}

/// <summary>
/// The complete result of inspecting a CSV file: how to read it, what is in it,
/// and anything questionable that was noticed on the way.
/// </summary>
public sealed class CsvAnalysis
{
    /// <summary>The analysed file.</summary>
    public required string FilePath { get; init; }

    /// <summary>The encoding the file is read with.</summary>
    public required Encoding Encoding { get; init; }

    /// <summary>Details of how the encoding was arrived at.</summary>
    public required EncodingDetectionResult EncodingResult { get; init; }

    /// <summary>The field separator.</summary>
    public required char Separator { get; init; }

    /// <summary>Details of how the separator was arrived at.</summary>
    public required SeparatorDetectionResult SeparatorResult { get; init; }

    /// <summary>The header decision and its reasoning.</summary>
    public required HeaderDetectionResult HeaderResult { get; init; }

    /// <summary>Column names and types.</summary>
    public required IReadOnlyList<ColumnInfo> Columns { get; init; }

    /// <summary>The first records of the file, for the preview grid.</summary>
    public required IReadOnlyList<CsvRecord> SampleRecords { get; init; }

    /// <summary>One-based physical line on which the first content record starts.</summary>
    public required int FirstContentLine { get; init; }

    /// <summary>How many records were read for analysis.</summary>
    public required int AnalyzedRecordCount { get; init; }

    /// <summary>
    /// Physical lines in the whole file. Counted separately from analysis, so
    /// unlike before this is the real figure and not the sample size.
    /// </summary>
    public required long TotalLineCount { get; init; }

    /// <summary>True when every analysed record had the same number of fields.</summary>
    public required bool IsFieldCountConsistent { get; init; }

    /// <summary>Things worth telling the user about.</summary>
    public required IReadOnlyList<string> Warnings { get; init; }

    /// <summary>Whether the first record names the columns.</summary>
    public bool HasHeader => HeaderResult.HasHeader;

    /// <summary>The resolved column names.</summary>
    public string[] FieldNames => HeaderResult.FieldNames;

    /// <summary>Expected number of fields per record.</summary>
    public int FieldCount => Columns.Count;

    /// <summary>Fraction of analysed records agreeing on the field count.</summary>
    public double SeparatorConfidence => SeparatorResult.Confidence;

    /// <summary>Renders the analysis as the lines shown in the log.</summary>
    public IEnumerable<string> ToLogLines()
    {
        yield return $"Encoding: {EncodingResult.Describe()}";
        yield return $"Separator: {SeparatorToken.Describe(Separator)}" +
                     (SeparatorResult.WasForced
                         ? " (forced)"
                         : $" ({SeparatorConfidence:P0} consistent)");
        yield return $"Columns: {FieldCount}, {HeaderResult.Describe()}";
        yield return $"Lines: {TotalLineCount} total, {AnalyzedRecordCount} analysed" +
                     (IsFieldCountConsistent ? "" : ", field count varies");

        foreach (var column in Columns)
            yield return $"  [{column.Index}] {column.Describe()}";

        foreach (var warning in Warnings)
            yield return $"Warning: {warning}";
    }
}
