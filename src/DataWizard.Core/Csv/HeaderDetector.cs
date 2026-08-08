using DataWizard.Core.Configuration;

namespace DataWizard.Core.Csv;

/// <summary>One scored signal contributing to the header decision.</summary>
/// <param name="Name">Display name of the signal.</param>
/// <param name="Value">Signal strength between 0 and 1.</param>
/// <param name="Weight">Configured influence of the signal.</param>
/// <param name="Applied">False when the signal is switched off or could not be evaluated.</param>
/// <param name="Detail">Human-readable explanation, shown in the analysis panel.</param>
public readonly record struct HeaderSignal(
    string Name,
    double Value,
    double Weight,
    bool Applied,
    string Detail);

/// <summary>The header decision, with the reasoning behind it.</summary>
public sealed class HeaderDetectionResult
{
    /// <summary>Whether the first content record is a header.</summary>
    public required bool HasHeader { get; init; }

    /// <summary>Combined weighted score between 0 and 1.</summary>
    public required double Score { get; init; }

    /// <summary>The score the file had to reach.</summary>
    public required double Threshold { get; init; }

    /// <summary>Whether the decision came from a setting or from scoring.</summary>
    public required HeaderMode Mode { get; init; }

    /// <summary>The individual signals, for display.</summary>
    public required IReadOnlyList<HeaderSignal> Signals { get; init; }

    /// <summary>
    /// Column names: the header values when a header was found, otherwise
    /// generated names such as <c>Column 1</c>.
    /// </summary>
    public required string[] FieldNames { get; init; }

    /// <summary>A one-line summary for the log.</summary>
    public string Describe() => Mode switch
    {
        HeaderMode.Always => "header: yes (forced)",
        HeaderMode.Never => "header: no (forced)",
        _ => $"header: {(HasHeader ? "yes" : "no")} (score {Score:F2} vs threshold {Threshold:F2})"
    };
}

/// <summary>
/// Decides whether the first record of a CSV file names the columns.
/// </summary>
/// <remarks>
/// Four independent signals are scored and combined with configurable weights, so
/// a file that only satisfies one of them does not automatically count as having a
/// header. Each signal is reported back, which means the analysis panel can show
/// why a file was classified the way it was instead of leaving the user guessing.
/// </remarks>
public sealed class HeaderDetector
{
    private readonly DetectionSettings _settings;
    private readonly ValueTypeDetector _typeDetector;

    public HeaderDetector(DetectionSettings settings, ValueTypeDetector? typeDetector = null)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _typeDetector = typeDetector ?? new ValueTypeDetector(settings);
    }

    /// <summary>
    /// Scores <paramref name="records"/>, whose first entry is the candidate header
    /// row and the rest sample data.
    /// </summary>
    public HeaderDetectionResult Detect(IReadOnlyList<CsvRecord> records)
    {
        var first = records.Count > 0 ? records[0] : null;
        var columnCount = first?.FieldCount ?? 0;

        if (_settings.HeaderMode == HeaderMode.Never || first is null || columnCount == 0)
        {
            return new HeaderDetectionResult
            {
                HasHeader = false,
                Score = 0d,
                Threshold = _settings.HeaderScoreThreshold,
                Mode = _settings.HeaderMode == HeaderMode.Never ? HeaderMode.Never : HeaderMode.Auto,
                Signals = [],
                FieldNames = GenerateNames(columnCount)
            };
        }

        var signals = ScoreSignals(first, records);

        if (_settings.HeaderMode == HeaderMode.Always)
        {
            return new HeaderDetectionResult
            {
                HasHeader = true,
                Score = 1d,
                Threshold = _settings.HeaderScoreThreshold,
                Mode = HeaderMode.Always,
                Signals = signals,
                FieldNames = ResolveNames(first)
            };
        }

        var totalWeight = signals.Where(s => s.Applied).Sum(s => s.Weight);
        var score = totalWeight > 0
            ? signals.Where(s => s.Applied).Sum(s => s.Value * s.Weight) / totalWeight
            : 0d;

        var hasHeader = score >= _settings.HeaderScoreThreshold;

        return new HeaderDetectionResult
        {
            HasHeader = hasHeader,
            Score = score,
            Threshold = _settings.HeaderScoreThreshold,
            Mode = HeaderMode.Auto,
            Signals = signals,
            FieldNames = hasHeader ? ResolveNames(first) : GenerateNames(columnCount)
        };
    }

    private List<HeaderSignal> ScoreSignals(CsvRecord first, IReadOnlyList<CsvRecord> records)
    {
        var signals = new List<HeaderSignal>(4);
        var columnCount = first.FieldCount;

        // ── Known field names ───────────────────────────────────────────────
        if (_settings.UseKnownFieldNames)
        {
            var matches = 0;
            for (var i = 0; i < columnCount; i++)
            {
                var name = first.Fields[i].Value;
                if (_settings.KnownFieldNames.Any(p => p.IsMatch(name)))
                    matches++;
            }

            var value = columnCount > 0 ? (double)matches / columnCount : 0d;
            signals.Add(new HeaderSignal(
                "Known field names",
                value,
                _settings.KnownFieldNameWeight,
                _settings.KnownFieldNames.Count > 0,
                $"{matches} of {columnCount} column names matched a known pattern"));
        }

        // ── Type divergence ─────────────────────────────────────────────────
        if (_settings.UseTypeDivergence)
        {
            var dataRecords = records.Skip(1).Where(r => !r.IsBlank).Take(50).ToList();

            if (dataRecords.Count == 0)
            {
                signals.Add(new HeaderSignal(
                    "Type divergence",
                    0d,
                    _settings.TypeDivergenceWeight,
                    false,
                    "no data rows below the first row to compare against"));
            }
            else
            {
                var diverging = 0;
                var comparable = 0;

                for (var column = 0; column < columnCount; column++)
                {
                    var headerType = _typeDetector.Detect(first.Fields[column]).Type;

                    var typedBelow = 0;
                    var samples = 0;

                    foreach (var record in dataRecords)
                    {
                        if (column >= record.FieldCount)
                            continue;

                        var field = record.Fields[column];
                        if (field.IsEmpty)
                            continue;

                        samples++;
                        if (_typeDetector.Detect(field).Type != FieldDataType.Text)
                            typedBelow++;
                    }

                    if (samples == 0)
                        continue;

                    comparable++;

                    // The tell-tale shape of a header: a text label sitting on top
                    // of a column that is consistently numeric or date-valued.
                    if (headerType == FieldDataType.Text && (double)typedBelow / samples >= 0.7)
                        diverging++;
                }

                var value = comparable > 0 ? (double)diverging / comparable : 0d;
                signals.Add(new HeaderSignal(
                    "Type divergence",
                    value,
                    _settings.TypeDivergenceWeight,
                    comparable > 0,
                    $"{diverging} of {comparable} columns are text on top of typed data"));
            }
        }

        // ── Uniqueness ──────────────────────────────────────────────────────
        if (_settings.UseUniqueness)
        {
            var names = first.Fields
                .Select(f => f.Value.Trim())
                .Where(v => v.Length > 0)
                .ToList();

            var distinct = names.Distinct(StringComparer.OrdinalIgnoreCase).Count();
            var value = names.Count > 0 ? (double)distinct / names.Count : 0d;

            signals.Add(new HeaderSignal(
                "Unique values",
                value,
                _settings.UniquenessWeight,
                names.Count > 0,
                $"{distinct} of {names.Count} non-empty values are distinct"));
        }

        // ── No empty cells ──────────────────────────────────────────────────
        if (_settings.UseNonEmptyRule)
        {
            var filled = first.Fields.Count(f => !f.IsEmpty);
            var value = columnCount > 0 ? (double)filled / columnCount : 0d;

            signals.Add(new HeaderSignal(
                "No empty cells",
                value,
                _settings.NonEmptyWeight,
                true,
                $"{filled} of {columnCount} cells carry a value"));
        }

        return signals;
    }

    /// <summary>
    /// Turns header values into usable column names, filling in blanks and
    /// disambiguating duplicates so downstream rules always have something to
    /// match against.
    /// </summary>
    private static string[] ResolveNames(CsvRecord header)
    {
        var names = new string[header.FieldCount];
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < header.FieldCount; i++)
        {
            var name = header.Fields[i].Value.Trim();

            if (name.Length == 0)
                name = $"Column {i + 1}";

            if (seen.TryGetValue(name, out var count))
            {
                seen[name] = count + 1;
                name = $"{name} ({count + 1})";
            }
            else
            {
                seen[name] = 1;
            }

            names[i] = name;
        }

        return names;
    }

    private static string[] GenerateNames(int count)
    {
        var names = new string[Math.Max(0, count)];
        for (var i = 0; i < names.Length; i++)
            names[i] = $"Column {i + 1}";

        return names;
    }
}
