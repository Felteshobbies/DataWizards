using System.Globalization;
using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;

namespace DataWizard.Core.Diff;

/// <summary>How rows of the two tables are matched to each other.</summary>
public enum DiffMatchMode
{
    /// <summary>Row 1 of the reference is compared with row 1 of the candidate, and so on.</summary>
    Positional,

    /// <summary>Rows are matched by one or more key columns.</summary>
    Keyed
}

/// <summary>What one difference consists of.</summary>
public enum DiffKind
{
    /// <summary>A cell holds different values in the two tables.</summary>
    ValueChanged,

    /// <summary>A row (or a duplicate of a key) exists only in the reference.</summary>
    RowOnlyInReference,

    /// <summary>A row (or a duplicate of a key) exists only in the candidate.</summary>
    RowOnlyInCandidate,

    /// <summary>A column exists only in the reference.</summary>
    ColumnOnlyInReference,

    /// <summary>A column exists only in the candidate.</summary>
    ColumnOnlyInCandidate
}

/// <summary>
/// The parameters of one comparison. Everything is optional; the defaults give
/// an exact, position-based comparison of two tables.
/// </summary>
public sealed class DiffOptions
{
    /// <summary>Whether rows are matched by position or by key columns.</summary>
    public DiffMatchMode MatchMode { get; set; } = DiffMatchMode.Positional;

    /// <summary>
    /// Key columns for <see cref="DiffMatchMode.Keyed"/>. When both tables have
    /// a header the entries are column names (case-insensitive); otherwise they
    /// are one-based column positions.
    /// </summary>
    public string[] KeyColumns { get; set; } = [];

    /// <summary>
    /// Two numbers are equal when they differ by at most this amount, regardless
    /// of their size.
    /// </summary>
    public double AbsoluteTolerance { get; set; }

    /// <summary>
    /// Two numbers are equal when they differ by at most this fraction of the
    /// larger of the two, for example 0.001 is a per-mille tolerance.
    /// </summary>
    public double RelativeTolerance { get; set; }

    /// <summary>Two dates are equal when they differ by at most this span.</summary>
    public TimeSpan DateTolerance { get; set; } = TimeSpan.Zero;

    /// <summary>Compare text without regard to upper and lower case.</summary>
    public bool IgnoreCase { get; set; }

    /// <summary>Strip surrounding whitespace before comparing.</summary>
    public bool TrimValues { get; set; } = true;

    /// <summary>
    /// Column names excluded from the comparison. A column that is ignored on
    /// one side does not produce a structural difference either.
    /// </summary>
    public string[] IgnoredColumns { get; set; } = [];

    /// <summary>Clamps values that would make the comparison meaningless.</summary>
    public void Normalize()
    {
        AbsoluteTolerance = Math.Max(0d, AbsoluteTolerance);
        RelativeTolerance = Math.Max(0d, RelativeTolerance);
        DateTolerance = DateTolerance < TimeSpan.Zero ? TimeSpan.Zero : DateTolerance;
        KeyColumns ??= [];
        IgnoredColumns ??= [];
    }

    /// <summary>Creates a copy, so the UI can edit options without touching the live instance.</summary>
    public DiffOptions Clone() => new()
    {
        MatchMode = MatchMode,
        KeyColumns = [.. KeyColumns],
        AbsoluteTolerance = AbsoluteTolerance,
        RelativeTolerance = RelativeTolerance,
        DateTolerance = DateTolerance,
        IgnoreCase = IgnoreCase,
        TrimValues = TrimValues,
        IgnoredColumns = [.. IgnoredColumns]
    };
}

/// <summary>
/// One difference between the two tables. Row numbers are one-based data rows;
/// zero means the row does not exist on that side.
/// </summary>
public sealed class DiffEntry
{
    /// <summary>What kind of difference this is.</summary>
    public required DiffKind Kind { get; init; }

    /// <summary>Data row in the reference, one-based; zero when not applicable.</summary>
    public int ReferenceRow { get; init; }

    /// <summary>Data row in the candidate, one-based; zero when not applicable.</summary>
    public int CandidateRow { get; init; }

    /// <summary>The column the difference is in; empty for row-level differences.</summary>
    public string Column { get; init; } = string.Empty;

    /// <summary>The value in the reference; empty when the row or column does not exist there.</summary>
    public string ReferenceValue { get; init; } = string.Empty;

    /// <summary>The value in the candidate; empty when the row or column does not exist there.</summary>
    public string CandidateValue { get; init; } = string.Empty;

    /// <summary>Extra context, for example the key of a row that exists on one side only.</summary>
    public string Detail { get; init; } = string.Empty;
}

/// <summary>
/// The result of comparing two tables: the individual differences, counts per
/// kind and everything worth telling the user.
/// </summary>
public sealed class DiffReport
{
    /// <summary>When the comparison was made.</summary>
    public required DateTime CreatedAt { get; init; }

    /// <summary>The reference file.</summary>
    public required string ReferencePath { get; init; }

    /// <summary>The candidate file.</summary>
    public required string CandidatePath { get; init; }

    /// <summary>The worksheet the reference was read from, for workbooks.</summary>
    public string? ReferenceSheet { get; init; }

    /// <summary>The worksheet the candidate was read from, for workbooks.</summary>
    public string? CandidateSheet { get; init; }

    /// <summary>The options the comparison ran with.</summary>
    public required DiffOptions Options { get; init; }

    /// <summary>Data rows in the reference table.</summary>
    public int ReferenceRowCount { get; init; }

    /// <summary>Data rows in the candidate table.</summary>
    public int CandidateRowCount { get; init; }

    /// <summary>Columns in the reference table.</summary>
    public int ReferenceColumnCount { get; init; }

    /// <summary>Columns in the candidate table.</summary>
    public int CandidateColumnCount { get; init; }

    /// <summary>Every difference found, in a stable order.</summary>
    public required IReadOnlyList<DiffEntry> Entries { get; init; }

    /// <summary>Cells with different values.</summary>
    public int ValueChangedCount { get; init; }

    /// <summary>Rows that exist only in the reference.</summary>
    public int RowOnlyInReferenceCount { get; init; }

    /// <summary>Rows that exist only in the candidate.</summary>
    public int RowOnlyInCandidateCount { get; init; }

    /// <summary>Columns that exist only in the reference.</summary>
    public int ColumnOnlyInReferenceCount { get; init; }

    /// <summary>Columns that exist only in the candidate.</summary>
    public int ColumnOnlyInCandidateCount { get; init; }

    /// <summary>Rows that share a key with another row of the same table.</summary>
    public int DuplicateKeyCount { get; init; }

    /// <summary>Things worth telling the user, from loading and from the comparison.</summary>
    public required IReadOnlyList<string> Warnings { get; init; }

    /// <summary>Whether anything differs at all.</summary>
    public bool HasDifferences => Entries.Count > 0;

    /// <summary>Total number of differences of every kind.</summary>
    public int DifferenceCount => Entries.Count;
}

/// <summary>
/// Compares two loaded tables and produces a <see cref="DiffReport"/>.
/// </summary>
/// <remarks>
/// Values are compared in a fixed order: numbers with a tolerance, then dates
/// with a tolerance, then text. A value with significant leading zeros is
/// always text, so <c>00123</c> and <c>123</c> are a difference. Keys are
/// compared as plain text, never with a numeric tolerance, because a key is an
/// identity and identities do not have tolerances.
/// </remarks>
public sealed class DiffEngine
{
    private readonly DetectionSettings _settings;
    private readonly ValueTypeDetector _detector;

    public DiffEngine(DetectionSettings settings)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _settings.Normalize();
        _detector = new ValueTypeDetector(_settings);
    }

    /// <summary>
    /// Compares <paramref name="reference"/> with <paramref name="candidate"/>
    /// and reports every difference.
    /// </summary>
    public DiffReport Compare(LoadedTable reference, LoadedTable candidate, DiffOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(reference);
        ArgumentNullException.ThrowIfNull(candidate);
        options ??= new DiffOptions();
        options.Normalize();

        var warnings = new List<string>(reference.Warnings);
        warnings.AddRange(candidate.Warnings);

        var trim = options.TrimValues;
        var comparer = options.IgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var ignored = new HashSet<string>(options.IgnoredColumns, StringComparer.OrdinalIgnoreCase);

        var keyPairs = options.MatchMode == DiffMatchMode.Keyed
            ? ResolveKeyColumns(reference, candidate, options.KeyColumns)
            : Array.Empty<(int RefIndex, int CandidateIndex, string Name)>();

        var refKeyColumns = keyPairs.Select(p => p.RefIndex).ToHashSet();
        var candKeyColumns = keyPairs.Select(p => p.CandidateIndex).ToHashSet();

        var (pairs, refOnly, candOnly) = PairColumns(
            reference, candidate, refKeyColumns, candKeyColumns, ignored, options.MatchMode == DiffMatchMode.Positional);

        var entries = new List<DiffEntry>();
        var duplicateKeyCount = 0;

        foreach (var name in refOnly)
            entries.Add(new DiffEntry { Kind = DiffKind.ColumnOnlyInReference, Column = name });

        foreach (var name in candOnly)
            entries.Add(new DiffEntry { Kind = DiffKind.ColumnOnlyInCandidate, Column = name });

        if (options.MatchMode == DiffMatchMode.Keyed)
            duplicateKeyCount = CompareKeyed(reference, candidate, keyPairs, pairs, entries, warnings, trim, comparer, options);
        else
            ComparePositional(reference, candidate, pairs, entries, trim, comparer, options);

        return new DiffReport
        {
            CreatedAt = DateTime.Now,
            ReferencePath = reference.SourcePath,
            CandidatePath = candidate.SourcePath,
            ReferenceSheet = reference.SheetName,
            CandidateSheet = candidate.SheetName,
            Options = options.Clone(),
            ReferenceRowCount = reference.RowCount,
            CandidateRowCount = candidate.RowCount,
            ReferenceColumnCount = reference.ColumnCount,
            CandidateColumnCount = candidate.ColumnCount,
            Entries = entries,
            ValueChangedCount = entries.Count(e => e.Kind == DiffKind.ValueChanged),
            RowOnlyInReferenceCount = entries.Count(e => e.Kind == DiffKind.RowOnlyInReference),
            RowOnlyInCandidateCount = entries.Count(e => e.Kind == DiffKind.RowOnlyInCandidate),
            ColumnOnlyInReferenceCount = entries.Count(e => e.Kind == DiffKind.ColumnOnlyInReference),
            ColumnOnlyInCandidateCount = entries.Count(e => e.Kind == DiffKind.ColumnOnlyInCandidate),
            DuplicateKeyCount = duplicateKeyCount,
            Warnings = warnings
        };
    }

    private void ComparePositional(
        LoadedTable reference,
        LoadedTable candidate,
        List<(int Ref, int Cand, string Name)> pairs,
        List<DiffEntry> entries,
        bool trim,
        StringComparison comparer,
        DiffOptions options)
    {
        var maxRows = Math.Max(reference.RowCount, candidate.RowCount);

        for (var i = 0; i < maxRows; i++)
        {
            if (i >= reference.RowCount)
            {
                entries.Add(new DiffEntry
                {
                    Kind = DiffKind.RowOnlyInCandidate,
                    CandidateRow = i + 1
                });
                continue;
            }

            if (i >= candidate.RowCount)
            {
                entries.Add(new DiffEntry
                {
                    Kind = DiffKind.RowOnlyInReference,
                    ReferenceRow = i + 1
                });
                continue;
            }

            var refRow = reference.Rows[i];
            var candRow = candidate.Rows[i];

            foreach (var (refColumn, candColumn, name) in pairs)
                CompareCell(refRow[refColumn], candRow[candColumn], name, i + 1, i + 1, entries, trim, comparer, options);
        }
    }

    /// <summary>
    /// Compares the rows of both tables matched by their key and returns the
    /// number of rows that share a key with another row of the same table.
    /// </summary>
    private int CompareKeyed(
        LoadedTable reference,
        LoadedTable candidate,
        (int RefIndex, int CandidateIndex, string Name)[] keyPairs,
        List<(int Ref, int Cand, string Name)> pairs,
        List<DiffEntry> entries,
        List<string> warnings,
        bool trim,
        StringComparison comparer,
        DiffOptions options)
    {
        var (refByKey, refOrder) = GroupByKey(reference, keyPairs, trim, comparer);
        var (candByKey, candOrder) = GroupByKey(candidate, keyPairs, trim, comparer);

        var duplicateKeyCount = 0;

        foreach (var key in refOrder.Concat(candOrder.Where(k => !refByKey.ContainsKey(k))))
        {
            var refRows = refByKey.TryGetValue(key, out var r) ? r : Array.Empty<int>();
            var candRows = candByKey.TryGetValue(key, out var c) ? c : Array.Empty<int>();

            var display = DisplayKey(key, keyPairs);

            if (refRows.Length > 1)
            {
                duplicateKeyCount += refRows.Length - 1;
                warnings.Add($"Key '{display}' occurs {refRows.Length} times in the reference; rows are compared in order of appearance.");
            }

            if (candRows.Length > 1)
            {
                duplicateKeyCount += candRows.Length - 1;
                warnings.Add($"Key '{display}' occurs {candRows.Length} times in the candidate; rows are compared in order of appearance.");
            }

            var max = Math.Max(refRows.Length, candRows.Length);

            for (var k = 0; k < max; k++)
            {
                if (k >= refRows.Length)
                {
                    entries.Add(new DiffEntry
                    {
                        Kind = DiffKind.RowOnlyInCandidate,
                        CandidateRow = candRows[k] + 1,
                        Detail = display
                    });
                    continue;
                }

                if (k >= candRows.Length)
                {
                    entries.Add(new DiffEntry
                    {
                        Kind = DiffKind.RowOnlyInReference,
                        ReferenceRow = refRows[k] + 1,
                        Detail = display
                    });
                    continue;
                }

                var refRow = reference.Rows[refRows[k]];
                var candRow = candidate.Rows[candRows[k]];

                foreach (var (refColumn, candColumn, name) in pairs)
                    CompareCell(refRow[refColumn], candRow[candColumn], name, refRows[k] + 1, candRows[k] + 1, entries, trim, comparer, options);
            }
        }

        return duplicateKeyCount;
    }

    /// <summary>
    /// Compares one cell pair and appends a <see cref="DiffKind.ValueChanged"/>
    /// entry when the values differ beyond the configured tolerances.
    /// </summary>
    private void CompareCell(
        string a,
        string b,
        string column,
        int referenceRow,
        int candidateRow,
        List<DiffEntry> entries,
        bool trim,
        StringComparison comparer,
        DiffOptions options)
    {
        var x = trim ? a.Trim() : a;
        var y = trim ? b.Trim() : b;

        if (x.Equals(y, comparer))
            return;

        // A significant leading zero makes the value an identifier. The numbers
        // 00123 and 123 are the same number and different values, so this is a
        // difference and no tolerance may hide it.
        if (ValueTypeDetector.HasSignificantLeadingZero(x) || ValueTypeDetector.HasSignificantLeadingZero(y))
        {
            entries.Add(new DiffEntry
            {
                Kind = DiffKind.ValueChanged,
                ReferenceRow = referenceRow,
                CandidateRow = candidateRow,
                Column = column,
                ReferenceValue = x,
                CandidateValue = y
            });
            return;
        }

        if (_detector.TryParseNumber(x, out var nx, out _) && _detector.TryParseNumber(y, out var ny, out _))
        {
            var tolerance = Math.Max(
                options.AbsoluteTolerance,
                options.RelativeTolerance * Math.Max(Math.Abs(nx), Math.Abs(ny)));

            if (Math.Abs(nx - ny) <= tolerance)
                return;
        }
        else if (TryParseDateValue(x, out var dx) && TryParseDateValue(y, out var dy))
        {
            if (Math.Abs((dx - dy).TotalSeconds) <= options.DateTolerance.TotalSeconds)
                return;
        }

        entries.Add(new DiffEntry
        {
            Kind = DiffKind.ValueChanged,
            ReferenceRow = referenceRow,
            CandidateRow = candidateRow,
            Column = column,
            ReferenceValue = x,
            CandidateValue = y
        });
    }

    /// <summary>
    /// Parses a date with the configured formats, falling back to the invariant
    /// round-trip format that date-styled workbook cells are normalised to.
    /// </summary>
    private bool TryParseDateValue(string value, out DateTime date)
    {
        if (_detector.TryParseDate(value, out date))
            return true;

        return DateTime.TryParseExact(value, "o", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
    }

    private (Dictionary<string, int[]> ByKey, List<string> Order) GroupByKey(
        LoadedTable table,
        (int RefIndex, int CandidateIndex, string Name)[] keyPairs,
        bool trim,
        StringComparison comparer)
    {
        var byKey = new Dictionary<string, List<int>>(
            comparer == StringComparison.OrdinalIgnoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);

        var order = new List<string>();

        for (var i = 0; i < table.RowCount; i++)
        {
            var key = BuildKey(table.Rows[i], keyPairs, trim);

            if (!byKey.TryGetValue(key, out var list))
            {
                list = [];
                byKey[key] = list;
                order.Add(key);
            }

            list.Add(i);
        }

        var arrays = byKey.ToDictionary(kv => kv.Key, kv => kv.Value.ToArray());
        return (arrays, order);
    }

    private static string BuildKey(string[] row, (int RefIndex, int CandidateIndex, string Name)[] keyPairs, bool trim)
    {
        var parts = new string[keyPairs.Length];

        for (var i = 0; i < keyPairs.Length; i++)
        {
            var value = row.Length > keyPairs[i].RefIndex ? row[keyPairs[i].RefIndex] : string.Empty;
            parts[i] = trim ? value.Trim() : value;
        }

        return string.Join('\u0001', parts);
    }

    private static string DisplayKey(string key, (int RefIndex, int CandidateIndex, string Name)[] keyPairs)
    {
        var parts = key.Split('\u0001');
        var display = new string[parts.Length];

        for (var i = 0; i < parts.Length; i++)
            display[i] = keyPairs.Length > i ? $"{keyPairs[i].Name}: {parts[i]}" : parts[i];

        return string.Join(", ", display);
    }

    /// <summary>
    /// Resolves the key columns to column indices of both tables.
    /// </summary>
    /// <remarks>
    /// When both tables have a header the key entries are column names, matched
    /// case-insensitively. Without a header they are one-based positions,
    /// because there is no other way to name a column.
    /// </remarks>
    private static (int RefIndex, int CandidateIndex, string Name)[] ResolveKeyColumns(
        LoadedTable reference,
        LoadedTable candidate,
        string[] keyColumns)
    {
        if (keyColumns.Length == 0)
            throw new ArgumentException("Keyed comparison needs at least one key column.", nameof(keyColumns));

        var byName = reference.HasHeader && candidate.HasHeader;
        var resolved = new (int RefIndex, int CandidateIndex, string Name)[keyColumns.Length];
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < keyColumns.Length; i++)
        {
            var entry = keyColumns[i].Trim();

            if (!seen.Add(entry))
                throw new ArgumentException($"Key column '{entry}' is listed more than once.", nameof(keyColumns));

            int refIndex, candIndex;
            var name = entry;

            if (byName)
            {
                refIndex = FindColumn(reference.Header, entry);
                candIndex = FindColumn(candidate.Header, entry);

                if (refIndex < 0)
                    throw new ArgumentException($"Key column '{entry}' was not found in the reference table.", nameof(keyColumns));

                if (candIndex < 0)
                    throw new ArgumentException($"Key column '{entry}' was not found in the candidate table.", nameof(keyColumns));
            }
            else
            {
                var limit = Math.Min(reference.ColumnCount, candidate.ColumnCount);

                if (!int.TryParse(entry, NumberStyles.Integer, CultureInfo.InvariantCulture, out var position)
                    || position < 1
                    || position > limit)
                {
                    throw new ArgumentException(
                        $"Key column '{entry}' is not a position between 1 and {limit}; both tables lack a header, so keys are given as positions.",
                        nameof(keyColumns));
                }

                refIndex = candIndex = position - 1;
                name = reference.Header[refIndex];
            }

            resolved[i] = (refIndex, candIndex, name);
        }

        return resolved;
    }

    private static int FindColumn(string[] header, string entry) =>
        header.ToList().FindIndex(h => h.Trim().Equals(entry, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Pairs the columns that will be compared, and collects the columns that
    /// exist on one side only.
    /// </summary>
    /// <remarks>
    /// Positional comparison pairs by position. Keyed comparison pairs by name
    /// when both tables have a header, and by position otherwise. Key columns
    /// are never compared as cells, because the key is what matched the rows.
    /// </remarks>
    private static (List<(int Ref, int Cand, string Name)> Pairs, List<string> ReferenceOnly, List<string> CandidateOnly)
        PairColumns(
            LoadedTable reference,
            LoadedTable candidate,
            HashSet<int> refKeyColumns,
            HashSet<int> candKeyColumns,
            HashSet<string> ignored,
            bool positional)
    {
        var pairs = new List<(int, int, string)>();
        var referenceOnly = new List<string>();
        var candidateOnly = new List<string>();

        var refCount = reference.ColumnCount;
        var candCount = candidate.ColumnCount;
        var shared = Math.Min(refCount, candCount);

        if (!positional && reference.HasHeader && candidate.HasHeader)
        {
            var usedCandidate = new bool[candCount];

            for (var r = 0; r < refCount; r++)
            {
                if (refKeyColumns.Contains(r) || ignored.Contains(reference.Header[r]))
                    continue;

                var match = -1;

                for (var c = 0; c < candCount; c++)
                {
                    if (usedCandidate[c] || candKeyColumns.Contains(c) || ignored.Contains(candidate.Header[c]))
                        continue;

                    if (reference.Header[r].Trim().Equals(candidate.Header[c].Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        match = c;
                        break;
                    }
                }

                if (match < 0)
                {
                    referenceOnly.Add(reference.Header[r]);
                    continue;
                }

                usedCandidate[match] = true;
                pairs.Add((r, match, reference.Header[r]));
            }

            for (var c = 0; c < candCount; c++)
            {
                if (!usedCandidate[c] && !candKeyColumns.Contains(c) && !ignored.Contains(candidate.Header[c]))
                    candidateOnly.Add(candidate.Header[c]);
            }

            return (pairs, referenceOnly, candidateOnly);
        }

        for (var c = 0; c < shared; c++)
        {
            if (refKeyColumns.Contains(c) || candKeyColumns.Contains(c))
                continue;

            if (ignored.Contains(reference.Header[c]))
                continue;

            pairs.Add((c, c, reference.Header[c]));
        }

        for (var c = shared; c < refCount; c++)
        {
            if (!refKeyColumns.Contains(c) && !ignored.Contains(reference.Header[c]))
                referenceOnly.Add(reference.Header[c]);
        }

        for (var c = shared; c < candCount; c++)
        {
            if (!candKeyColumns.Contains(c) && !ignored.Contains(candidate.Header[c]))
                candidateOnly.Add(candidate.Header[c]);
        }

        return (pairs, referenceOnly, candidateOnly);
    }
}
