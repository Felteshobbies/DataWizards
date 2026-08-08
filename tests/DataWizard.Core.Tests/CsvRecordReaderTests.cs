using DataWizard.Core.Csv;

namespace DataWizard.Core.Tests;

public class CsvRecordReaderTests
{
    private static CsvField[] Split(string line, char separator = ';') =>
        CsvRecordReader.SplitLine(line, separator);

    private static List<CsvRecord> ReadAll(string text, char separator = ';')
    {
        using var reader = new CsvRecordReader(new StringReader(text), separator);
        return reader.ReadRecords().ToList();
    }

    [Fact]
    public void SplitsPlainFields()
    {
        var fields = Split("a;b;c");

        Assert.Equal(3, fields.Length);
        Assert.Equal(["a", "b", "c"], fields.Select(f => f.Value));
        Assert.All(fields, f => Assert.False(f.WasQuoted));
    }

    [Fact]
    public void KeepsEmptyFieldsInTheMiddleAndAtTheEnd()
    {
        var fields = Split("a;;b;");

        Assert.Equal(4, fields.Length);
        Assert.Equal(["a", "", "b", ""], fields.Select(f => f.Value));
    }

    [Fact]
    public void MarksQuotedFieldsSoLeadingZerosCanBePreserved()
    {
        var fields = Split("\"007\";42");

        Assert.True(fields[0].WasQuoted);
        Assert.Equal("007", fields[0].Value);
        Assert.False(fields[1].WasQuoted);
    }

    [Fact]
    public void SeparatorInsideQuotesDoesNotSplitTheField()
    {
        var fields = Split("\"a;b\";c");

        Assert.Equal(2, fields.Length);
        Assert.Equal("a;b", fields[0].Value);
        Assert.Equal("c", fields[1].Value);
    }

    [Fact]
    public void DoubledQuoteBecomesALiteralQuote()
    {
        var fields = Split("\"say \"\"hi\"\"\";x");

        Assert.Equal("say \"hi\"", fields[0].Value);
        Assert.Equal("x", fields[1].Value);
    }

    [Fact]
    public void QuotedFieldMaySpanSeveralLines()
    {
        var records = ReadAll("a;\"line1\nline2\";c\nd;e;f");

        Assert.Equal(2, records.Count);
        Assert.Equal("line1\nline2", records[0].Fields[1].Value);
        Assert.Equal(3, records[0].FieldCount);
        Assert.Equal(["d", "e", "f"], records[1].Fields.Select(f => f.Value));
    }

    [Fact]
    public void HandlesCrLfAndLfInTheSameFile()
    {
        var records = ReadAll("a;b\r\nc;d\ne;f");

        Assert.Equal(3, records.Count);
        Assert.Equal(["e", "f"], records[2].Fields.Select(f => f.Value));
    }

    [Fact]
    public void TrailingNewlineDoesNotProduceAnExtraRecord()
    {
        var records = ReadAll("a;b\r\n");

        Assert.Single(records);
    }

    [Fact]
    public void BlankLineIsReportedAsBlank()
    {
        var records = ReadAll("a;b\n\nc;d");

        Assert.Equal(3, records.Count);
        Assert.True(records[1].IsBlank);
        Assert.False(records[0].IsBlank);
    }

    [Fact]
    public void UnterminatedQuoteIsReportedButDataIsStillReturned()
    {
        using var reader = new CsvRecordReader(new StringReader("a;\"unclosed"), ';');
        var record = reader.ReadRecord();

        Assert.NotNull(record);
        Assert.Equal("unclosed", record.Fields[1].Value);
        Assert.True(reader.HasUnterminatedQuote);
    }

    [Fact]
    public void UnquotedWhitespaceIsTrimmedButQuotedWhitespaceIsKept()
    {
        var fields = Split("  a  ;\"  b  \"");

        Assert.Equal("a", fields[0].Value);
        Assert.Equal("  b  ", fields[1].Value);
    }

    [Fact]
    public void TracksTheLineARecordStartsOn()
    {
        var records = ReadAll("h1;h2\na;\"x\ny\"\nb;c");

        Assert.Equal(1, records[0].StartLine);
        Assert.Equal(2, records[1].StartLine);
        Assert.Equal(4, records[2].StartLine);
    }

    [Fact]
    public void TextAfterAClosingQuoteIsKeptRatherThanDiscarded()
    {
        var fields = Split("\"a\"tail;b");

        Assert.Equal("atail", fields[0].Value);
        Assert.True(fields[0].WasQuoted);
        Assert.Equal("b", fields[1].Value);
    }

    [Fact]
    public void TabSeparatedInputIsSupported()
    {
        var fields = Split("a\tb\tc", '\t');

        Assert.Equal(3, fields.Length);
    }
}
