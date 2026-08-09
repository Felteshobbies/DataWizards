using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;

namespace DataWizard.App.ViewModels;

/// <summary>One column as shown in the preview grid.</summary>
public sealed class ColumnPreviewRow
{
    public required int Index { get; init; }
    public required string Name { get; init; }
    public required string DetectedType { get; init; }
    public required string EffectiveType { get; init; }
    public required string Rule { get; init; }
    public required string Sample { get; init; }
    public required string Notes { get; init; }
}

/// <summary>One header signal as shown in the preview.</summary>
public sealed class HeaderSignalRow
{
    public required string Name { get; init; }
    public required string Score { get; init; }
    public required string Weight { get; init; }
    public required string Detail { get; init; }
}

/// <summary>One separator candidate as shown in the preview.</summary>
public sealed class SeparatorCandidateRow
{
    public required string Separator { get; init; }
    public required string Fields { get; init; }
    public required string Consistency { get; init; }
    public required bool IsChosen { get; init; }
}

/// <summary>
/// Presents a <see cref="CsvAnalysis"/> for the preview panel.
/// </summary>
/// <remarks>
/// The point of this panel is to answer "why did it read my file that way?".
/// Every decision the analyser made is shown together with the evidence behind it,
/// which is the difference between a converter you can debug and one you have to
/// guess at.
/// </remarks>
public sealed class AnalysisPreview
{
    private const int PreviewRowLimit = 12;
    private const int PreviewCellWidth = 18;

    public AnalysisPreview(CsvAnalysis analysis)
    {
        Encoding = analysis.EncodingResult.Describe();
        Separator = SeparatorToken.Describe(analysis.Separator) +
                    (analysis.SeparatorResult.WasForced
                        ? " (forced)"
                        : $" - {analysis.SeparatorConfidence:P0} of rows agree");

        Header = analysis.HeaderResult.Describe();
        Structure = $"{analysis.FieldCount} columns, {analysis.TotalLineCount} lines " +
                    $"({analysis.AnalyzedRecordCount} analysed)" +
                    (analysis.IsFieldCountConsistent ? string.Empty : ", field count varies");

        FirstContentLine = analysis.FirstContentLine;
        Warnings = analysis.Warnings.ToArray();
        HasWarnings = Warnings.Count > 0;

        Columns = analysis.Columns.Select(column => new ColumnPreviewRow
        {
            Index = column.Index,
            Name = column.Name,
            DetectedType = column.DetectedType.ToString(),
            EffectiveType = column.EffectiveType.ToString(),
            Rule = column.AppliedRule is { } rule
                ? (rule.ColumnIndex.HasValue ? $"column {rule.ColumnIndex}" : $"{rule.Match}: {rule.Pattern}")
                : string.Empty,
            Sample = FirstSample(analysis, column.Index),
            Notes = DescribeNotes(column)
        }).ToArray();

        HeaderSignals = analysis.HeaderResult.Signals.Select(signal => new HeaderSignalRow
        {
            Name = signal.Name,
            Score = signal.Applied ? signal.Value.ToString("P0") : "n/a",
            Weight = signal.Weight.ToString("F2"),
            Detail = signal.Detail
        }).ToArray();

        SeparatorCandidates = analysis.SeparatorResult.Candidates.Select(candidate => new SeparatorCandidateRow
        {
            Separator = SeparatorToken.Describe(candidate.Separator),
            Fields = candidate.FieldCount.ToString(),
            Consistency = candidate.Confidence.ToString("P0"),
            IsChosen = candidate.Separator == analysis.Separator
        }).ToArray();

        SampleText = BuildSampleText(analysis);

        Compact = BuildCompactSummary(analysis);
        AttentionReasons = CollectAttentionReasons(analysis);
        NeedsAttention = AttentionReasons.Count > 0;
        AttentionHeadline = NeedsAttention ? AttentionReasons[0] : "Detection looks unambiguous.";
    }

    /// <summary>
    /// A single line naming the separator, the header verdict and the column
    /// count. Enough to confirm the file was read as expected without opening the
    /// full panel.
    /// </summary>
    public string Compact { get; }

    /// <summary>
    /// True when something about the detection is worth a second look. This is
    /// what decides whether the analysis panel opens by itself.
    /// </summary>
    public bool NeedsAttention { get; }

    /// <summary>Why the analysis wants attention, most important first.</summary>
    public IReadOnlyList<string> AttentionReasons { get; }

    /// <summary>The first attention reason, or a reassurance when there are none.</summary>
    public string AttentionHeadline { get; }

    private static string BuildCompactSummary(CsvAnalysis analysis)
    {
        var separator = SeparatorToken.Describe(analysis.Separator);
        var header = analysis.HasHeader ? "header" : "no header";

        return $"{separator} · {analysis.FieldCount} columns · {header} · " +
               $"{analysis.TotalLineCount} lines · {analysis.Encoding.WebName}";
    }

    /// <summary>
    /// Works out whether the detection was clear-cut.
    /// </summary>
    /// <remarks>
    /// Warnings are the obvious trigger, but a header score sitting right on the
    /// threshold is the more insidious case: nothing is wrong, yet the decision
    /// could just as easily have gone the other way, and that is precisely when a
    /// human should look.
    /// </remarks>
    private static List<string> CollectAttentionReasons(CsvAnalysis analysis)
    {
        var reasons = new List<string>();

        if (analysis.Warnings.Count > 0)
            reasons.AddRange(analysis.Warnings);

        if (analysis.HeaderResult.Mode == HeaderMode.Auto)
        {
            var margin = Math.Abs(analysis.HeaderResult.Score - analysis.HeaderResult.Threshold);

            if (margin < BorderlineHeaderMargin)
            {
                reasons.Add(
                    $"The header decision was close: score {analysis.HeaderResult.Score:F2} against a " +
                    $"threshold of {analysis.HeaderResult.Threshold:F2}. Check that row {analysis.FirstContentLine} " +
                    $"really is {(analysis.HasHeader ? "a header" : "data")}.");
            }
        }

        var mixed = analysis.Columns.Where(IsGenuinelyInconsistent).ToList();
        if (mixed.Count > 0)
        {
            reasons.Add(
                $"{mixed.Count} column(s) hold more than one kind of value: " +
                string.Join(", ", mixed.Take(4).Select(c => c.Name)) +
                (mixed.Count > 4 ? ", ..." : string.Empty));
        }

        return reasons;
    }

    /// <summary>
    /// Whether a column's values genuinely disagree about what they are.
    /// </summary>
    /// <remarks>
    /// <see cref="ColumnInfo.IsMixed"/> is too blunt to drive an interruption. It
    /// counts two ordinary situations as mixed: whole numbers sitting among
    /// decimals, which is simply what a numeric column looks like, and a column
    /// that a rule has already settled - such as a postal code column holding
    /// both <c>01067</c> and <c>80331</c>, where the rule makes both text and the
    /// outcome is not in doubt. Flagging either would fire on almost every file.
    /// </remarks>
    private static bool IsGenuinelyInconsistent(ColumnInfo column)
    {
        // A rule is the user having already decided; there is nothing to review.
        if (column.AppliedRule is not null)
            return false;

        var kinds = column.TypeCounts
            .Where(kvp => kvp.Value > 0)
            .Select(kvp => CollapseToKind(kvp.Key))
            .Distinct()
            .Count();

        return kinds > 1;
    }

    /// <summary>
    /// Groups the numeric types together, so an integer among decimals does not
    /// read as a disagreement.
    /// </summary>
    private static FieldDataType CollapseToKind(FieldDataType type) =>
        type == FieldDataType.Integer ? FieldDataType.Decimal : type;

    /// <summary>
    /// How close a header score may sit to the threshold before the decision is
    /// treated as borderline. Deliberately narrow: this is meant to catch a
    /// decision that was within a hair of flipping, not merely one that was not
    /// emphatic.
    /// </summary>
    private const double BorderlineHeaderMargin = 0.05;

    /// <summary>How the file is decoded.</summary>
    public string Encoding { get; }

    /// <summary>Which separator was chosen and how confidently.</summary>
    public string Separator { get; }

    /// <summary>The header verdict.</summary>
    public string Header { get; }

    /// <summary>Shape of the file.</summary>
    public string Structure { get; }

    /// <summary>Line the data starts on.</summary>
    public int FirstContentLine { get; }

    /// <summary>Per-column findings.</summary>
    public IReadOnlyList<ColumnPreviewRow> Columns { get; }

    /// <summary>Why the header verdict came out the way it did.</summary>
    public IReadOnlyList<HeaderSignalRow> HeaderSignals { get; }

    /// <summary>How each candidate separator scored.</summary>
    public IReadOnlyList<SeparatorCandidateRow> SeparatorCandidates { get; }

    /// <summary>Anything questionable about the file.</summary>
    public IReadOnlyList<string> Warnings { get; }

    /// <summary>Whether there is anything in <see cref="Warnings"/>.</summary>
    public bool HasWarnings { get; }

    /// <summary>The first rows, aligned into columns for reading.</summary>
    public string SampleText { get; }

    private static string DescribeNotes(ColumnInfo column)
    {
        var notes = new List<string>();

        if (column.WasOverridden)
            notes.Add($"rule overrides {column.DetectedType}");

        if (column.IsMixed)
        {
            var breakdown = string.Join(", ",
                column.TypeCounts
                    .Where(kvp => kvp.Value > 0)
                    .OrderByDescending(kvp => kvp.Value)
                    .Select(kvp => $"{kvp.Key} x{kvp.Value}"));

            notes.Add($"mixed: {breakdown}");
        }

        if (column.EmptyCount > 0)
            notes.Add($"{column.EmptyCount} empty");

        return string.Join("; ", notes);
    }

    private static string FirstSample(CsvAnalysis analysis, int columnIndex)
    {
        var records = analysis.HasHeader ? analysis.SampleRecords.Skip(1) : analysis.SampleRecords;

        foreach (var record in records)
        {
            if (columnIndex >= record.FieldCount)
                continue;

            var field = record.Fields[columnIndex];

            if (!field.IsEmpty)
                return field.WasQuoted ? $"\"{field.Value}\"" : field.Value;
        }

        return string.Empty;
    }

    /// <summary>
    /// Renders the first rows as a fixed-width table. Reading the values in aligned
    /// columns is what makes a shifted field obvious at a glance.
    /// </summary>
    private static string BuildSampleText(CsvAnalysis analysis)
    {
        if (analysis.SampleRecords.Count == 0)
            return "(no rows)";

        var builder = new StringBuilder();
        var rows = analysis.SampleRecords.Take(PreviewRowLimit).ToList();
        var columnCount = rows.Max(r => r.FieldCount);

        if (analysis.HasHeader)
        {
            builder.Append("     ");
            for (var i = 0; i < columnCount; i++)
                builder.Append(Pad(i < analysis.FieldNames.Length ? analysis.FieldNames[i] : $"Column {i + 1}"));

            builder.AppendLine();
            builder.Append("     ");
            builder.Append(new string('-', Math.Min(columnCount * PreviewCellWidth, 200)));
            builder.AppendLine();
        }

        var lineNumber = analysis.FirstContentLine;
        var startIndex = analysis.HasHeader ? 1 : 0;

        for (var r = startIndex; r < rows.Count; r++)
        {
            var record = rows[r];
            builder.Append(record.StartLine.ToString().PadLeft(4)).Append(' ');

            for (var c = 0; c < columnCount; c++)
            {
                var value = c < record.FieldCount ? record.Fields[c].Value : string.Empty;
                builder.Append(Pad(value.Replace("\n", "\\n").Replace("\r", string.Empty)));
            }

            builder.AppendLine();
            lineNumber++;
        }

        if (analysis.SampleRecords.Count > PreviewRowLimit)
            builder.Append($"... {analysis.AnalyzedRecordCount - PreviewRowLimit} more analysed row(s)");

        return builder.ToString().TrimEnd();
    }

    private static string Pad(string value)
    {
        if (value.Length >= PreviewCellWidth)
            return value[..(PreviewCellWidth - 2)] + ".." + " ";

        return value.PadRight(PreviewCellWidth);
    }
}
