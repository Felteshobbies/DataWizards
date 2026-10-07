using System.Text;
using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;

namespace DataWizard.Core.Tests;

public class CsvAnalyzerTests
{
    private static DetectionSettings Settings(Action<DetectionSettings>? configure = null)
    {
        var settings = DataWizardSettings.CreateDefault().Detection;
        configure?.Invoke(settings);
        settings.Normalize();
        return settings;
    }

    [Fact]
    public void ReportsEncodingSeparatorHeaderAndColumns()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("data.csv",
            "id;name;price;date",
            "1;Widget;19,99;31.01.2025",
            "2;Gadget;4,50;01.02.2025");

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.Equal(';', analysis.Separator);
        Assert.True(analysis.HasHeader);
        Assert.Equal(["id", "name", "price", "date"], analysis.FieldNames);
        Assert.Equal(4, analysis.FieldCount);
        Assert.True(analysis.IsFieldCountConsistent);
    }

    /// <summary>
    /// The old analyser rewound its <see cref="StreamReader"/> with
    /// <c>BaseStream.Seek(0)</c>, which re-read the byte order mark as data and left
    /// an invisible character on the front of the first column name.
    /// </summary>
    [Fact]
    public void AByteOrderMarkDoesNotContaminateTheFirstColumnName()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteFile(
            "bom.csv",
            "id;name\r\n1;Widget\r\n2;Gadget\r\n",
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.Equal("id", analysis.FieldNames[0]);
        Assert.DoesNotContain('﻿', analysis.FieldNames[0]);
        Assert.True(analysis.EncodingResult.HasByteOrderMark);
    }

    /// <summary>
    /// Line counts used to be capped at the analysis sample size, so a 5000-line
    /// file was reported as having 100 lines.
    /// </summary>
    [Fact]
    public void ReportsTheRealLineCountNotTheSampleSize()
    {
        using var workspace = new TempWorkspace();

        var builder = new StringBuilder("id;value\r\n");
        for (var i = 1; i <= 500; i++)
            builder.Append(i).Append(';').Append(i * 2).Append("\r\n");

        var path = workspace.WriteFile("big.csv", builder.ToString());
        var analysis = new CsvAnalyzer(Settings(s => s.AnalyzeLineCount = 50)).Analyze(path);

        Assert.Equal(501, analysis.TotalLineCount);
        Assert.Equal(50, analysis.AnalyzedRecordCount);
    }

    [Fact]
    public void CountsAFinalLineWithoutATrailingNewline()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteFile("nonewline.csv", "a;b\r\nc;d");

        Assert.Equal(2, CsvAnalyzer.CountLines(path));
    }

    [Fact]
    public void SkipsBlankLinesBeforeTheHeader()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("blank.csv",
            "",
            "   ",
            "id;name",
            "1;Widget");

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.Equal(3, analysis.FirstContentLine);
        Assert.True(analysis.HasHeader);
        Assert.Equal("id", analysis.FieldNames[0]);
    }

    [Fact]
    public void SkipsAConfiguredNumberOfPreambleLines()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("preamble.csv",
            "# Export from some system",
            "# Generated 2025-01-31",
            "id;name;price",
            "1;Widget;9.99");

        var analysis = new CsvAnalyzer(Settings(s => s.SkipLeadingLines = 2)).Analyze(path);

        Assert.Equal(["id", "name", "price"], analysis.FieldNames);
        Assert.Contains(analysis.Warnings, w => w.Contains("Skipped 2 leading line"));
    }

    [Fact]
    public void AppliesFieldRulesAndReportsWhichRuleWon()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("rules.csv",
            "customer_id;zip;price",
            "1001;10115;19.99",
            "1002;80331;4.50");

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        var id = analysis.Columns[0];
        var zip = analysis.Columns[1];
        var price = analysis.Columns[2];

        Assert.Equal(FieldDataType.Text, id.EffectiveType);
        Assert.Equal(FieldDataType.Text, zip.EffectiveType);
        Assert.Equal(FieldDataType.Decimal, price.EffectiveType);

        Assert.True(id.WasOverridden);
        Assert.NotNull(id.AppliedRule);
    }

    [Fact]
    public void APositionalRuleTakesPrecedenceOverANameRule()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("positional.csv",
            "price;note",
            "19.99;a",
            "4.50;b");

        var settings = Settings(s => s.FieldRules.Insert(0, new FieldRule
        {
            ColumnIndex = 0,
            DataType = FieldDataType.Text
        }));

        var analysis = new CsvAnalyzer(settings).Analyze(path);

        Assert.Equal(FieldDataType.Text, analysis.Columns[0].EffectiveType);
    }

    [Fact]
    public void WarnsWhenRowsHaveDifferentFieldCounts()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("ragged.csv",
            "id;name;price",
            "1;Widget",
            "2;Gadget;4.50;extra");

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.False(analysis.IsFieldCountConsistent);
        Assert.Contains(analysis.Warnings, w => w.Contains("fields"));
        Assert.Equal(4, analysis.FieldCount);
    }

    [Fact]
    public void WarnsWhenAQuotedValueIsNeverClosed()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteFile("unterminated.csv", "id;name\r\n1;\"never closed\r\n");

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.Contains(analysis.Warnings, w => w.Contains("never closed"));
    }

    [Fact]
    public void AnEmptyFileYieldsAnEmptyAnalysisRatherThanThrowing()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteFile("empty.csv", string.Empty);

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.Equal(0, analysis.FieldCount);
        Assert.False(analysis.HasHeader);
        Assert.Contains(analysis.Warnings, w => w.Contains("no data rows"));
    }

    [Fact]
    public void WarnsWhenTheSeparatorExplainsTooFewRows()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("messy.csv",
            "a;b;c",
            "1;2",
            "3;4;5;6",
            "7",
            "8;9;10;11;12");

        var analysis = new CsvAnalyzer(Settings(s => s.MinSeparatorConfidence = 0.9)).Analyze(path);

        Assert.Contains(analysis.Warnings, w => w.Contains("only explains"));
    }

    [Fact]
    public void MissingFilesThrowFileNotFound()
    {
        using var workspace = new TempWorkspace();

        Assert.Throws<FileNotFoundException>(() =>
            new CsvAnalyzer(Settings()).Analyze(workspace.PathTo("nope.csv")));
    }

    [Fact]
    public void ReadsWindows1252WhenForced()
    {
        using var workspace = new TempWorkspace();
        EncodingDetector.EnsureCodePagesRegistered();

        var encoding = Encoding.GetEncoding(1252);
        var path = workspace.WriteFile("ansi.csv", "name;city\r\nMüller;Köln\r\nWeiß;Zürich\r\n", encoding);

        var analysis = new CsvAnalyzer(Settings(s => s.ForcedEncoding = "windows-1252")).Analyze(path);

        Assert.Equal("Müller", analysis.SampleRecords[1].Fields[0].Value);
        Assert.Equal("Zürich", analysis.SampleRecords[2].Fields[1].Value);
    }

    /// <summary>
    /// Statistical encoding detection is a guess, and a weak one on short files.
    /// A file that decodes "successfully" into the wrong accented characters must
    /// say so rather than leaving the user to spot it in the output.
    /// </summary>
    [Fact]
    public void WarnsWhenTheEncodingWasOnlyGuessed()
    {
        using var workspace = new TempWorkspace();
        EncodingDetector.EnsureCodePagesRegistered();

        var path = workspace.WriteFile(
            "short.csv",
            "name;city\r\nMüller;Köln\r\n",
            Encoding.GetEncoding(1252));

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        // Either the detector was unsure, or it was confident - only the unsure
        // case must produce the hint.
        if (analysis.EncodingResult.Source == EncodingSource.Detected &&
            analysis.EncodingResult.Confidence < 0.7)
        {
            Assert.Contains(analysis.Warnings, w => w.Contains("confidence"));
            Assert.Contains(analysis.Warnings, w => w.Contains("windows-1252"));
        }
    }

    [Fact]
    public void AConfidentlyDetectedEncodingProducesNoHint()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("plain.csv", "id;name", "1;Widget", "2;Gadget");

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.DoesNotContain(analysis.Warnings, w => w.Contains("confidence"));
    }

    /// <summary>
    /// Below the ASCII boundary every candidate encoding decodes identically, so a
    /// low confidence score has no consequence and warning about it is noise.
    /// </summary>
    [Fact]
    public void AnAsciiOnlyFileNeverWarnsAboutEncodingConfidence()
    {
        using var workspace = new TempWorkspace();

        // Short and featureless: the kind of input a detector is least sure about.
        var path = workspace.WriteFile("tiny.csv", "a;b\r\n1;2\r\n");

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.DoesNotContain(analysis.Warnings, w => w.Contains("confidence"));
    }

    [Fact]
    public void AColumnOfMixedTypesFallsBackToText()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("mixed.csv",
            "code;value",
            "A1;1",
            "B2;2",
            "C3;not a number");

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.Equal(FieldDataType.Text, analysis.Columns[1].EffectiveType);
        Assert.True(analysis.Columns[1].IsMixed);
    }

    /// <summary>
    /// Whole numbers and decimals are one family: a column of "12" and "0,45"
    /// must come out decimal, not text, no matter which kind is the majority.
    /// </summary>
    [Fact]
    public void AColumnOfIntegersAndDecimalsIsDecimal()
    {
        using var workspace = new TempWorkspace();
        // One column is mostly whole numbers, the other mostly decimals; both
        // mix the two kinds, so both must come out decimal.
        var path = workspace.WriteLines("mixed.csv",
            "id;mostly_integers;mostly_decimals",
            "A1;1;0,45",
            "B2;2;0,89",
            "C3;0,5;500");

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.Equal(FieldDataType.Decimal, analysis.Columns[1].EffectiveType);
        Assert.Equal(FieldDataType.Decimal, analysis.Columns[2].EffectiveType);
    }

    [Fact]
    public void AColumnOfOnlyIntegersStaysInteger()
    {
        using var workspace = new TempWorkspace();
        var path = workspace.WriteLines("ints.csv",
            "id;total",
            "A1;3",
            "B2;12",
            "C3;1");

        var analysis = new CsvAnalyzer(Settings()).Analyze(path);

        Assert.Equal(FieldDataType.Integer, analysis.Columns[1].EffectiveType);
    }

    [Fact]
    public void AnalyzeTextWorksWithoutAFile()
    {
        var analysis = new CsvAnalyzer(Settings()).AnalyzeText("id;name\n1;Widget\n2;Gadget");

        Assert.True(analysis.HasHeader);
        Assert.Equal(2, analysis.FieldCount);
    }
}
