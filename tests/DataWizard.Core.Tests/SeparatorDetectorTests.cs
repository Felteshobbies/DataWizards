using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;

namespace DataWizard.Core.Tests;

public class SeparatorDetectorTests
{
    private static DetectionSettings Settings(Action<DetectionSettings>? configure = null)
    {
        var settings = DataWizardSettings.CreateDefault().Detection;
        configure?.Invoke(settings);
        settings.Normalize();
        return settings;
    }

    [Theory]
    [InlineData("a;b;c\n1;2;3", ';')]
    [InlineData("a,b,c\n1,2,3", ',')]
    [InlineData("a\tb\tc\n1\t2\t3", '\t')]
    [InlineData("a|b|c\n1|2|3", '|')]
    public void FindsTheSeparatorInAnUnambiguousFile(string sample, char expected)
    {
        var result = SeparatorDetector.Detect(sample, Settings());

        Assert.Equal(expected, result.Separator);
        Assert.Equal(3, result.FieldCount);
        Assert.Equal(1d, result.Confidence, precision: 3);
    }

    /// <summary>
    /// The old detector simply counted characters, so a prose column full of
    /// commas outvoted the semicolons that actually delimited the file. Scoring on
    /// structural consistency fixes that.
    /// </summary>
    [Fact]
    public void ProseFullOfCommasDoesNotOutvoteTheRealSeparator()
    {
        var sample = string.Join('\n',
            "id;description;price",
            "1;\"red, round, and, rather, large\";9.99",
            "2;\"blue, square, small, light, cheap\";4.50",
            "3;\"green, oval, medium, sturdy, nice\";7.25");

        var result = SeparatorDetector.Detect(sample, Settings());

        Assert.Equal(';', result.Separator);
        Assert.Equal(3, result.FieldCount);
    }

    [Fact]
    public void SeparatorsInsideQuotesAreNotCounted()
    {
        var sample = "a;b\n\"x;y;z\";c";
        var result = SeparatorDetector.Detect(sample, Settings());

        Assert.Equal(';', result.Separator);
        Assert.Equal(2, result.FieldCount);
    }

    [Fact]
    public void ConfidenceDropsWhenRowsDisagree()
    {
        var consistent = SeparatorDetector.Detect("a;b;c\n1;2;3\n4;5;6", Settings());
        var ragged = SeparatorDetector.Detect("a;b;c\n1;2\n4;5;6;7\n8", Settings());

        Assert.Equal(1d, consistent.Confidence, precision: 3);
        Assert.True(ragged.Confidence < consistent.Confidence);
    }

    [Fact]
    public void ASingleColumnFileIsReportedWithoutConfidence()
    {
        var result = SeparatorDetector.Detect("alpha\nbeta\ngamma", Settings());

        Assert.Equal(1, result.FieldCount);
        Assert.Equal(0d, result.Confidence);
    }

    [Fact]
    public void AForcedSeparatorIsUsedAndFlagged()
    {
        var settings = Settings(s => s.ForcedSeparator = "Comma");
        var result = SeparatorDetector.Detect("a;b;c\n1;2;3", settings);

        Assert.Equal(',', result.Separator);
        Assert.True(result.WasForced);
    }

    [Fact]
    public void AForcedSeparatorIsStillScoredSoAWrongChoiceCanBeReported()
    {
        var settings = Settings(s => s.ForcedSeparator = "Comma");
        var result = SeparatorDetector.Detect("a;b;c\n1;2;3", settings);

        // Comma yields one column here, which the analysis surfaces as a warning.
        Assert.Equal(1, result.FieldCount);
    }

    [Fact]
    public void EveryCandidateIsReportedSoTheChoiceCanBeExplained()
    {
        var result = SeparatorDetector.Detect("a;b;c\n1;2;3", Settings());

        Assert.Equal(4, result.Candidates.Count);
        Assert.Equal(';', result.Candidates[0].Separator);
        Assert.Contains(result.Candidates, c => c.Separator == ',');
    }

    [Fact]
    public void TheCandidateListIsConfigurable()
    {
        var settings = Settings(s => s.CandidateSeparators = "#~");
        var result = SeparatorDetector.Detect("a#b#c\n1#2#3", settings);

        Assert.Equal('#', result.Separator);
        Assert.Equal(2, result.Candidates.Count);
    }

    [Fact]
    public void EmptyInputDoesNotThrow()
    {
        var result = SeparatorDetector.Detect(string.Empty, Settings());

        Assert.Equal(0d, result.Confidence);
    }
}
