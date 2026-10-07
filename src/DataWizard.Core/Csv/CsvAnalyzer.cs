using System.Text;
using DataWizard.Core.Configuration;

namespace DataWizard.Core.Csv;

/// <summary>
/// Inspects a CSV file and reports everything needed to convert it: encoding,
/// separator, header, and the type of each column.
/// </summary>
/// <remarks>
/// Analysis reads a bounded sample of the file - <see cref="DetectionSettings.AnalyzeLineCount"/>
/// records - and each stage is re-run from a fresh reader rather than by rewinding
/// one. Seeking a <see cref="StreamReader"/> back to zero re-reads the byte order
/// mark as content, which used to leave an invisible character glued to the first
/// column name.
/// </remarks>
public sealed class CsvAnalyzer
{
    /// <summary>
    /// Detector confidence below which the encoding is reported as uncertain.
    /// Above the setting's own threshold the file is still read, but the user is
    /// told the choice was a guess.
    /// </summary>
    /// <remarks>
    /// Statistical detection is inherently a guess, and short files give it little
    /// to work with - the single-byte code pages in particular are hard to tell
    /// apart, so a file can decode "successfully" into the wrong accented
    /// characters with nothing to show for it. The bar is set low enough that a
    /// merely unemphatic UTF-8 verdict does not raise a false alarm.
    /// </remarks>
    private const double UncertainEncodingConfidence = 0.7;

    private readonly DetectionSettings _settings;

    public CsvAnalyzer(DetectionSettings settings)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _settings.Normalize();
    }

    /// <summary>Analyses a file on disk.</summary>
    /// <exception cref="FileNotFoundException">The file does not exist.</exception>
    public CsvAnalysis Analyze(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("CSV file not found.", path);

        var warnings = new List<string>();
        var encodingResult = EncodingDetector.Detect(path, _settings);

        if (encodingResult.Source == EncodingSource.Fallback)
        {
            warnings.Add(
                $"Encoding could not be detected with sufficient confidence; " +
                $"falling back to {encodingResult.Encoding.WebName}. Non-ASCII characters may be wrong.");
        }

        var sample = ReadSample(path, encodingResult.Encoding, out var firstContentLine, out var skippedLines);

        if (encodingResult.Source == EncodingSource.Detected &&
            encodingResult.Confidence < UncertainEncodingConfidence &&
            ContainsNonAscii(sample))
        {
            // Only worth saying when it could actually matter. A pure ASCII file
            // decodes identically under every encoding on offer, so a low
            // confidence score there has no consequence to warn about.
            warnings.Add(
                $"Encoding was detected as {encodingResult.Encoding.WebName} with only " +
                $"{encodingResult.Confidence:P0} confidence. If accented characters look wrong, " +
                "set the encoding explicitly (windows-1252 for most Western European exports).");
        }

        if (skippedLines > 0)
            warnings.Add($"Skipped {skippedLines} leading line(s) before the first content row.");

        if (sample.Length == 0)
        {
            return EmptyAnalysis(path, encodingResult, warnings);
        }

        var separatorResult = SeparatorDetector.Detect(sample, _settings);

        if (!separatorResult.WasForced && separatorResult.Confidence < _settings.MinSeparatorConfidence)
        {
            warnings.Add(
                $"Separator {SeparatorToken.Describe(separatorResult.Separator)} only explains " +
                $"{separatorResult.Confidence:P0} of the rows (threshold {_settings.MinSeparatorConfidence:P0}). " +
                "The file may use a different separator or have irregular rows.");
        }

        if (separatorResult.FieldCount <= 1)
            warnings.Add("No separator produced more than one column; the file is being treated as single-column.");

        var records = ReadRecords(sample, separatorResult.Separator);

        if (records.Count == 0)
            return EmptyAnalysis(path, encodingResult, warnings);

        var typeDetector = new ValueTypeDetector(_settings);
        var headerResult = new HeaderDetector(_settings, typeDetector).Detect(records);

        var fieldCounts = records.Select(r => r.FieldCount).ToList();
        var isConsistent = fieldCounts.Distinct().Count() == 1;

        if (!isConsistent)
        {
            warnings.Add(
                $"Records have between {fieldCounts.Min()} and {fieldCounts.Max()} fields. " +
                "Short rows are padded and long rows keep their extra values.");
        }

        var columnCount = fieldCounts.Max();
        var columns = BuildColumns(records, headerResult, columnCount, typeDetector);

        var unterminated = HasUnterminatedQuote(sample, separatorResult.Separator);
        if (unterminated)
            warnings.Add("A quoted value is never closed. The file does not follow RFC 4180 and may be truncated.");

        return new CsvAnalysis
        {
            FilePath = path,
            Encoding = encodingResult.Encoding,
            EncodingResult = encodingResult,
            Separator = separatorResult.Separator,
            SeparatorResult = separatorResult,
            HeaderResult = headerResult,
            Columns = columns,
            SampleRecords = records,
            FirstContentLine = firstContentLine,
            AnalyzedRecordCount = records.Count,
            TotalLineCount = CountLines(path),
            IsFieldCountConsistent = isConsistent,
            Warnings = warnings
        };
    }

    /// <summary>
    /// Reads the opening lines of a file as text, decoded the way analysis would
    /// decode it.
    /// </summary>
    /// <remarks>
    /// Lines are returned exactly as they appear, including any preamble that
    /// <see cref="DetectionSettings.SkipLeadingLines"/> would drop. Seeing the
    /// preamble is the point: it is what explains why the setting is needed.
    /// </remarks>
    /// <param name="path">The file to read.</param>
    /// <param name="maxLines">How many lines to return.</param>
    public string ReadSampleText(string path, int maxLines = 25)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("File not found.", path);

        var encoding = EncodingDetector.Detect(path, _settings).Encoding;

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true);

        var builder = new StringBuilder();
        var lines = 0;

        while (lines < maxLines && reader.ReadLine() is { } line)
        {
            if (lines > 0)
                builder.Append('\n');

            builder.Append(line);
            lines++;
        }

        return builder.ToString();
    }

    /// <summary>
    /// Analyses text held in memory. Used by the settings screen, which lets the
    /// user paste a few rows and watch the detection react.
    /// </summary>
    public CsvAnalysis AnalyzeText(string text, string label = "(text)")
    {
        var warnings = new List<string>();
        var encodingResult = new EncodingDetectionResult(
            new UTF8Encoding(false), 1d, EncodingSource.Forced, false);

        var sample = TrimToSample(text, out var firstContentLine, out _);

        if (sample.Length == 0)
            return EmptyAnalysis(label, encodingResult, warnings);

        var separatorResult = SeparatorDetector.Detect(sample, _settings);
        var records = ReadRecords(sample, separatorResult.Separator);

        if (records.Count == 0)
            return EmptyAnalysis(label, encodingResult, warnings);

        var typeDetector = new ValueTypeDetector(_settings);
        var headerResult = new HeaderDetector(_settings, typeDetector).Detect(records);
        var fieldCounts = records.Select(r => r.FieldCount).ToList();
        var columns = BuildColumns(records, headerResult, fieldCounts.Max(), typeDetector);

        return new CsvAnalysis
        {
            FilePath = label,
            Encoding = encodingResult.Encoding,
            EncodingResult = encodingResult,
            Separator = separatorResult.Separator,
            SeparatorResult = separatorResult,
            HeaderResult = headerResult,
            Columns = columns,
            SampleRecords = records,
            FirstContentLine = firstContentLine,
            AnalyzedRecordCount = records.Count,
            TotalLineCount = records.Sum(r => (long)r.LineCount),
            IsFieldCountConsistent = fieldCounts.Distinct().Count() == 1,
            Warnings = warnings
        };
    }

    private List<CsvRecord> ReadRecords(string sample, char separator)
    {
        using var reader = new CsvRecordReader(
            new StringReader(sample),
            separator,
            _settings.QuoteChar,
            _settings.TrimWhitespace);

        var records = new List<CsvRecord>();

        foreach (var record in reader.ReadRecords())
        {
            if (_settings.SkipEmptyLines && record.IsBlank)
                continue;

            records.Add(record);

            if (records.Count >= _settings.AnalyzeLineCount)
                break;
        }

        return records;
    }

    private List<ColumnInfo> BuildColumns(
        IReadOnlyList<CsvRecord> records,
        HeaderDetectionResult headerResult,
        int columnCount,
        ValueTypeDetector typeDetector)
    {
        var dataRecords = headerResult.HasHeader ? records.Skip(1).ToList() : records.ToList();
        var columns = new List<ColumnInfo>(columnCount);

        for (var index = 0; index < columnCount; index++)
        {
            var counts = new Dictionary<FieldDataType, int>();
            var empty = 0;
            var samples = 0;

            foreach (var record in dataRecords)
            {
                if (index >= record.FieldCount)
                {
                    empty++;
                    continue;
                }

                var field = record.Fields[index];

                if (field.IsEmpty)
                {
                    empty++;
                    continue;
                }

                samples++;
                var type = typeDetector.Detect(field).Type;
                counts[type] = counts.GetValueOrDefault(type) + 1;
            }

            var detected = ResolveDominantType(counts);
            var name = index < headerResult.FieldNames.Length
                ? headerResult.FieldNames[index]
                : $"Column {index + 1}";

            var column = new ColumnInfo
            {
                Index = index,
                Name = name,
                DetectedType = detected,
                EffectiveType = detected,
                SampleCount = samples,
                EmptyCount = empty,
                TypeCounts = counts
            };

            var rule = FindRule(name, index);
            if (rule is not null)
            {
                column.AppliedRule = rule;
                column.EffectiveType = rule.DataType == FieldDataType.Auto ? detected : rule.DataType;
            }

            columns.Add(column);
        }

        return columns;
    }

    /// <summary>
    /// Finds the first field rule that applies to a column. Position rules are
    /// considered before name rules, because a rule that names an exact column is
    /// the more specific statement of intent.
    /// </summary>
    public FieldRule? FindRule(string? fieldName, int columnIndex)
    {
        foreach (var rule in _settings.FieldRules)
        {
            if (rule.Enabled && rule.ColumnIndex == columnIndex)
                return rule;
        }

        foreach (var rule in _settings.FieldRules)
        {
            if (!rule.ColumnIndex.HasValue && rule.AppliesTo(fieldName, columnIndex))
                return rule;
        }

        return null;
    }

    /// <summary>
    /// Picks the type that best describes a column. A column is only numeric or a
    /// date if the great majority of its values are, because one stray value must
    /// not drag the whole column into a type that then fails to convert.
    /// </summary>
    private static FieldDataType ResolveDominantType(Dictionary<FieldDataType, int> counts)
    {
        var total = counts.Values.Sum();
        if (total == 0)
            return FieldDataType.Text;

        if (counts.TryGetValue(FieldDataType.Text, out var textCount) && textCount > 0)
            return FieldDataType.Text;

        // Integer and decimal are one family: Excel stores both as numbers, and a
        // decimal format renders whole numbers unchanged. A column of "12" and
        // "0,45" is a decimal column, not a text column - even when the integers
        // are the majority.
        var numeric = counts.GetValueOrDefault(FieldDataType.Integer)
                    + counts.GetValueOrDefault(FieldDataType.Decimal);

        if (numeric == total)
            return counts.GetValueOrDefault(FieldDataType.Decimal) > 0
                ? FieldDataType.Decimal
                : FieldDataType.Integer;

        var dominant = counts.OrderByDescending(kvp => kvp.Value).First();

        if ((double)dominant.Value / total < 0.9)
            return FieldDataType.Text;

        return dominant.Key;
    }

    /// <summary>
    /// Reads the leading portion of a file as text, skipping the configured
    /// preamble and any blank lines before the first content row.
    /// </summary>
    private string ReadSample(string path, Encoding encoding, out int firstContentLine, out int skippedLines)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true);

        var builder = new StringBuilder();
        var line = 0;
        var kept = 0;

        firstContentLine = 1;
        skippedLines = 0;
        var foundContent = false;

        // Read a few extra records so header scoring still has data rows even when
        // the file starts with blank lines.
        var limit = _settings.AnalyzeLineCount + _settings.SkipLeadingLines + 8;

        while (kept < limit && reader.ReadLine() is { } text)
        {
            line++;

            if (line <= _settings.SkipLeadingLines)
            {
                skippedLines++;
                continue;
            }

            if (!foundContent)
            {
                if (_settings.SkipEmptyLines && string.IsNullOrWhiteSpace(text))
                {
                    skippedLines++;
                    continue;
                }

                foundContent = true;
                firstContentLine = line;
            }

            builder.Append(text).Append('\n');
            kept++;
        }

        return builder.ToString();
    }

    private string TrimToSample(string text, out int firstContentLine, out int skippedLines)
    {
        using var reader = new StringReader(text);
        var builder = new StringBuilder();
        var line = 0;
        var kept = 0;

        firstContentLine = 1;
        skippedLines = 0;
        var foundContent = false;
        var limit = _settings.AnalyzeLineCount + _settings.SkipLeadingLines + 8;

        while (kept < limit && reader.ReadLine() is { } current)
        {
            line++;

            if (line <= _settings.SkipLeadingLines)
            {
                skippedLines++;
                continue;
            }

            if (!foundContent)
            {
                if (_settings.SkipEmptyLines && string.IsNullOrWhiteSpace(current))
                {
                    skippedLines++;
                    continue;
                }

                foundContent = true;
                firstContentLine = line;
            }

            builder.Append(current).Append('\n');
            kept++;
        }

        return builder.ToString();
    }

    private bool HasUnterminatedQuote(string sample, char separator)
    {
        using var reader = new CsvRecordReader(
            new StringReader(sample), separator, _settings.QuoteChar, _settings.TrimWhitespace);

        while (reader.ReadRecord() is not null)
        {
        }

        return reader.HasUnterminatedQuote;
    }

    private CsvAnalysis EmptyAnalysis(
        string path,
        EncodingDetectionResult encodingResult,
        List<string> warnings)
    {
        warnings.Add("The file contains no data rows.");

        return new CsvAnalysis
        {
            FilePath = path,
            Encoding = encodingResult.Encoding,
            EncodingResult = encodingResult,
            Separator = _settings.FixedSeparator ?? ';',
            SeparatorResult = new SeparatorDetectionResult(
                _settings.FixedSeparator ?? ';', 0, 0d, _settings.FixedSeparator.HasValue, []),
            HeaderResult = new HeaderDetectionResult
            {
                HasHeader = false,
                Score = 0d,
                Threshold = _settings.HeaderScoreThreshold,
                Mode = _settings.HeaderMode,
                Signals = [],
                FieldNames = []
            },
            Columns = [],
            SampleRecords = [],
            FirstContentLine = 1,
            AnalyzedRecordCount = 0,
            TotalLineCount = File.Exists(path) ? CountLines(path) : 0,
            IsFieldCountConsistent = true,
            Warnings = warnings
        };
    }

    /// <summary>
    /// Whether the sample holds any character outside plain ASCII. Below that
    /// boundary every candidate encoding decodes identically.
    /// </summary>
    private static bool ContainsNonAscii(string sample)
    {
        foreach (var c in sample)
        {
            if (!char.IsAscii(c))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Counts physical lines by scanning bytes. This is what makes the reported
    /// line count the real one rather than the size of the analysis sample.
    /// </summary>
    public static long CountLines(string path)
    {
        try
        {
            using var stream = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 64 * 1024, FileOptions.SequentialScan);

            var buffer = new byte[64 * 1024];
            long lines = 0;
            var lastByte = 0;
            int read;

            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
            {
                for (var i = 0; i < read; i++)
                {
                    if (buffer[i] == (byte)'\n')
                        lines++;
                }

                lastByte = buffer[read - 1];
            }

            // A final line without a trailing newline still counts.
            if (lastByte != 0 && lastByte != (byte)'\n')
                lines++;

            return lines;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return 0;
        }
    }
}
