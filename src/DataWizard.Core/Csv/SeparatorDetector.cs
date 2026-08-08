using DataWizard.Core.Configuration;

namespace DataWizard.Core.Csv;

/// <summary>How well one candidate separator explains the sampled text.</summary>
/// <param name="Separator">The candidate character.</param>
/// <param name="FieldCount">The most common field count it produces.</param>
/// <param name="Confidence">Fraction of records (0-1) that agree on <paramref name="FieldCount"/>.</param>
/// <param name="RecordCount">How many records were scored.</param>
public readonly record struct SeparatorCandidate(
    char Separator,
    int FieldCount,
    double Confidence,
    int RecordCount)
{
    /// <summary>A description for the log and the analysis panel.</summary>
    public string Describe() =>
        $"{SeparatorToken.Describe(Separator)}: {FieldCount} fields, {Confidence:P0} consistent";
}

/// <summary>The chosen separator plus the scores of everything that lost.</summary>
/// <param name="Separator">The separator to use.</param>
/// <param name="FieldCount">Expected number of fields per record.</param>
/// <param name="Confidence">Fraction of records that agree on the field count.</param>
/// <param name="WasForced">True when the separator came from the settings rather than detection.</param>
/// <param name="Candidates">All candidates, best first, so the UI can explain the choice.</param>
public readonly record struct SeparatorDetectionResult(
    char Separator,
    int FieldCount,
    double Confidence,
    bool WasForced,
    IReadOnlyList<SeparatorCandidate> Candidates);

/// <summary>
/// Works out which character separates the fields of a CSV file.
/// </summary>
/// <remarks>
/// The decision is made on structural consistency, not on raw frequency. Simply
/// counting characters - the approach used before 0.2 - picks the wrong answer
/// whenever a column contains prose: a single description field full of commas
/// outvotes the semicolons that actually delimit the file. Here each candidate is
/// used to parse the sample, and the winner is the one that yields the same field
/// count on the most records.
/// </remarks>
public static class SeparatorDetector
{
    /// <summary>
    /// Scores every candidate separator against <paramref name="sample"/> and
    /// returns the best one.
    /// </summary>
    public static SeparatorDetectionResult Detect(string sample, DetectionSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        var candidates = ScoreAll(sample, settings);

        if (settings.FixedSeparator is { } forced)
        {
            var match = candidates.FirstOrDefault(c => c.Separator == forced);

            // Score the forced separator too, so the analysis panel can still warn
            // when the manual choice does not fit the file.
            if (match.Separator == '\0')
                match = Score(sample, forced, settings);

            return new SeparatorDetectionResult(forced, match.FieldCount, match.Confidence, true, candidates);
        }

        var best = candidates.FirstOrDefault();

        // Nothing produced more than one column: the file is single-column, so any
        // separator is as good as none. Report the first candidate at zero
        // confidence rather than pretending to know.
        if (best.FieldCount <= 1)
        {
            var fallback = settings.SeparatorCandidates.FirstOrDefault(';');
            return new SeparatorDetectionResult(fallback, 1, 0d, false, candidates);
        }

        return new SeparatorDetectionResult(best.Separator, best.FieldCount, best.Confidence, false, candidates);
    }

    /// <summary>
    /// Scores every configured candidate, ordered by how well it fits.
    /// </summary>
    public static IReadOnlyList<SeparatorCandidate> ScoreAll(string sample, DetectionSettings settings)
    {
        var candidates = settings.SeparatorCandidates;
        var scored = new List<SeparatorCandidate>(candidates.Length);

        foreach (var candidate in candidates)
            scored.Add(Score(sample, candidate, settings));

        return scored
            .OrderByDescending(c => c.FieldCount > 1)
            .ThenByDescending(c => c.Confidence)
            .ThenByDescending(c => c.FieldCount)
            .ToList();
    }

    /// <summary>Scores a single candidate separator.</summary>
    public static SeparatorCandidate Score(string sample, char separator, DetectionSettings settings)
    {
        if (string.IsNullOrEmpty(sample))
            return new SeparatorCandidate(separator, 0, 0d, 0);

        using var reader = new CsvRecordReader(
            new StringReader(sample),
            separator,
            settings.QuoteChar,
            settings.TrimWhitespace);

        var counts = new Dictionary<int, int>();
        var records = 0;

        foreach (var record in reader.ReadRecords())
        {
            if (record.IsBlank)
                continue;

            records++;
            counts[record.FieldCount] = counts.GetValueOrDefault(record.FieldCount) + 1;

            if (records >= settings.AnalyzeLineCount)
                break;
        }

        if (records == 0)
            return new SeparatorCandidate(separator, 0, 0d, 0);

        // The modal field count wins; ties go to the wider layout, because a
        // separator that occasionally appears inside text produces a few wide
        // records rather than a consistent one.
        var mode = counts
            .OrderByDescending(kvp => kvp.Value)
            .ThenByDescending(kvp => kvp.Key)
            .First();

        var confidence = (double)mode.Value / records;
        return new SeparatorCandidate(separator, mode.Key, confidence, records);
    }
}
