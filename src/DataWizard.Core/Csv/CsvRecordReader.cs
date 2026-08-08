using System.Text;

namespace DataWizard.Core.Csv;

/// <summary>
/// Reads CSV records from a <see cref="TextReader"/> following RFC 4180, with the
/// leniencies real-world files need.
/// </summary>
/// <remarks>
/// <para>
/// Unlike a plain <c>string.Split</c> this understands quoted values, so a
/// separator, a doubled quote or a line break inside quotes no longer splits the
/// record apart.
/// </para>
/// <para>
/// Where a file breaks the specification the reader keeps going rather than
/// throwing: an unterminated quote runs to the end of the input, and stray
/// characters after a closing quote are appended to the value.
/// </para>
/// </remarks>
public sealed class CsvRecordReader : IDisposable
{
    private readonly TextReader _reader;
    private readonly char _separator;
    private readonly char _quote;
    private readonly bool _trimUnquoted;
    private readonly bool _leaveOpen;

    private readonly StringBuilder _value = new();
    private readonly List<CsvField> _fields = [];

    private int _pendingChar = -1;
    private int _line = 1;
    private bool _endOfStream;

    /// <summary>Creates a reader over <paramref name="reader"/>.</summary>
    /// <param name="reader">Source of the CSV text.</param>
    /// <param name="separator">The character between fields.</param>
    /// <param name="quote">The character that delimits a quoted field.</param>
    /// <param name="trimUnquoted">Strip surrounding whitespace from unquoted values.</param>
    /// <param name="leaveOpen">Do not dispose <paramref name="reader"/> with this instance.</param>
    public CsvRecordReader(
        TextReader reader,
        char separator,
        char quote = '"',
        bool trimUnquoted = true,
        bool leaveOpen = false)
    {
        _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        _separator = separator;
        _quote = quote;
        _trimUnquoted = trimUnquoted;
        _leaveOpen = leaveOpen;
    }

    /// <summary>One-based physical line the reader has consumed up to.</summary>
    public int CurrentLine => _line;

    /// <summary>
    /// True when a quoted field ran to the end of the input without a closing
    /// quote. The data was still returned, but the file is malformed.
    /// </summary>
    public bool HasUnterminatedQuote { get; private set; }

    /// <summary>
    /// Reads the next record, or returns <c>null</c> at the end of the input.
    /// </summary>
    public CsvRecord? ReadRecord()
    {
        if (_endOfStream)
            return null;

        var startLine = _line;
        _fields.Clear();
        _value.Clear();

        var inQuotes = false;
        var fieldWasQuoted = false;
        var sawAnyContent = false;

        while (true)
        {
            var read = ReadChar();

            if (read == -1)
            {
                _endOfStream = true;

                if (inQuotes)
                    HasUnterminatedQuote = true;

                // Trailing newline at the end of the file is not a record.
                if (!sawAnyContent && _fields.Count == 0 && _value.Length == 0)
                    return null;

                CommitField(fieldWasQuoted);
                return BuildRecord(startLine);
            }

            var c = (char)read;

            if (inQuotes)
            {
                if (c == _quote)
                {
                    // A doubled quote is a literal quote inside the value.
                    if (Peek() == _quote)
                    {
                        ReadChar();
                        _value.Append(_quote);
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    if (c == '\n')
                        _line++;

                    _value.Append(c);
                }

                sawAnyContent = true;
                continue;
            }

            if (c == _quote && _value.Length == 0 && !fieldWasQuoted)
            {
                inQuotes = true;
                fieldWasQuoted = true;
                sawAnyContent = true;
                continue;
            }

            if (c == _separator)
            {
                CommitField(fieldWasQuoted);
                fieldWasQuoted = false;
                sawAnyContent = true;
                continue;
            }

            if (c is '\r' or '\n')
            {
                // Consume the second half of a CRLF pair.
                if (c == '\r' && Peek() == '\n')
                    ReadChar();

                _line++;

                CommitField(fieldWasQuoted);
                return BuildRecord(startLine);
            }

            _value.Append(c);
            sawAnyContent = true;
        }
    }

    /// <summary>Reads every remaining record.</summary>
    public IEnumerable<CsvRecord> ReadRecords()
    {
        while (ReadRecord() is { } record)
            yield return record;
    }

    /// <summary>
    /// Reads at most <paramref name="maxRecords"/> records. Used by analysis, which
    /// only needs a sample of the file.
    /// </summary>
    public List<CsvRecord> ReadRecords(int maxRecords)
    {
        var records = new List<CsvRecord>();

        while (records.Count < maxRecords && ReadRecord() is { } record)
            records.Add(record);

        return records;
    }

    private void CommitField(bool wasQuoted)
    {
        var text = _value.ToString();

        if (!wasQuoted && _trimUnquoted)
            text = text.Trim();

        _fields.Add(new CsvField(text, wasQuoted));
        _value.Clear();
    }

    private CsvRecord BuildRecord(int startLine) => new()
    {
        Fields = _fields.ToArray(),
        StartLine = startLine,
        LineCount = Math.Max(1, _line - startLine)
    };

    private int ReadChar()
    {
        if (_pendingChar != -1)
        {
            var pending = _pendingChar;
            _pendingChar = -1;
            return pending;
        }

        return _reader.Read();
    }

    private int Peek()
    {
        if (_pendingChar == -1)
            _pendingChar = _reader.Read();

        return _pendingChar;
    }

    /// <summary>
    /// Splits a single line. Convenience for callers that already hold a string,
    /// such as previews; it cannot represent values containing line breaks.
    /// </summary>
    public static CsvField[] SplitLine(string line, char separator, char quote = '"', bool trimUnquoted = true)
    {
        using var reader = new CsvRecordReader(new StringReader(line), separator, quote, trimUnquoted);
        return reader.ReadRecord()?.Fields ?? [];
    }

    public void Dispose()
    {
        if (!_leaveOpen)
            _reader.Dispose();
    }
}
