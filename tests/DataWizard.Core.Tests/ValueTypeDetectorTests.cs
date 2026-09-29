using System.Globalization;
using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;

namespace DataWizard.Core.Tests;

public class ValueTypeDetectorTests
{
    private static ValueTypeDetector Detector(Action<DetectionSettings>? configure = null)
    {
        var settings = DataWizardSettings.CreateDefault().Detection;
        configure?.Invoke(settings);
        settings.Normalize();
        return new ValueTypeDetector(settings);
    }

    [Theory]
    [InlineData("42", FieldDataType.Integer)]
    [InlineData("-42", FieldDataType.Integer)]
    [InlineData("0", FieldDataType.Integer)]
    [InlineData("19.99", FieldDataType.Decimal)]
    [InlineData("19,99", FieldDataType.Decimal)]
    [InlineData("-0.5", FieldDataType.Decimal)]
    [InlineData("2025-01-31", FieldDataType.Date)]
    [InlineData("31.01.2025", FieldDataType.Date)]
    [InlineData("hello", FieldDataType.Text)]
    [InlineData("", FieldDataType.Text)]
    public void RecognisesTheObviousCases(string value, FieldDataType expected)
    {
        Assert.Equal(expected, Detector().Detect(value).Type);
    }

    [Theory]
    [InlineData("19.99", 19.99)]
    [InlineData("19,99", 19.99)]
    public void AcceptsBothEnglishAndGermanDecimals(string value, double expected)
    {
        var result = Detector().Detect(value);

        Assert.Equal(FieldDataType.Decimal, result.Type);
        Assert.Equal(expected, result.Number, precision: 6);
    }

    [Theory]
    [InlineData("007")]
    [InlineData("00123")]
    [InlineData("-0042")]
    public void KeepsLeadingZerosAsText(string value)
    {
        Assert.Equal(FieldDataType.Text, Detector().Detect(value).Type);
    }

    [Theory]
    [InlineData("0")]
    [InlineData("0.5")]
    [InlineData("0,5")]
    public void ASingleZeroBeforeADecimalIsStillANumber(string value)
    {
        Assert.NotEqual(FieldDataType.Text, Detector().Detect(value).Type);
    }

    [Fact]
    public void LeadingZeroPreservationCanBeSwitchedOff()
    {
        var result = Detector(s => s.PreserveLeadingZeros = false).Detect("007");

        Assert.Equal(FieldDataType.Integer, result.Type);
        Assert.Equal(7d, result.Number);
    }

    [Fact]
    public void LongDigitRunsStayTextBecauseADoubleWouldLosePrecision()
    {
        // A 13-digit EAN survives; an 18-digit identifier does not fit a double.
        Assert.Equal(FieldDataType.Text, Detector().Detect("123456789012345678").Type);
        Assert.Equal(FieldDataType.Integer, Detector().Detect("4006381333931").Type);
    }

    [Fact]
    public void TheDigitLimitIsConfigurable()
    {
        Assert.Equal(FieldDataType.Text, Detector(s => s.MaxIntegerDigits = 4).Detect("12345").Type);
        Assert.Equal(FieldDataType.Integer, Detector(s => s.MaxIntegerDigits = 5).Detect("12345").Type);
    }

    [Fact]
    public void QuotedValuesAreTextEvenWhenTheyLookNumeric()
    {
        var detector = Detector();

        Assert.Equal(FieldDataType.Text, detector.Detect(new CsvField("42", WasQuoted: true)).Type);
        Assert.Equal(FieldDataType.Integer, detector.Detect(new CsvField("42", WasQuoted: false)).Type);
    }

    [Fact]
    public void QuotedValuesCanBeAllowedToTypeNormally()
    {
        var detector = Detector(s => s.QuotedFieldsAreText = false);

        Assert.Equal(FieldDataType.Integer, detector.Detect(new CsvField("42", WasQuoted: true)).Type);
    }

    /// <summary>
    /// The old implementation called <c>DateTime.TryParse</c> without a culture, so
    /// the same file produced different results depending on the machine's regional
    /// settings. Parsing is now restricted to the configured formats.
    /// </summary>
    [Fact]
    public void DateParsingDoesNotDependOnTheAmbientCulture()
    {
        var original = CultureInfo.CurrentCulture;

        try
        {
            var results = new List<DateTime>();

            foreach (var culture in new[] { "en-US", "de-DE", "ja-JP" })
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(culture);
                var result = Detector().Detect("31.01.2025");

                Assert.Equal(FieldDataType.Date, result.Type);
                results.Add(result.Date);
            }

            Assert.Single(results.Distinct());
            Assert.Equal(new DateTime(2025, 1, 31), results[0]);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Theory]
    [InlineData("45.66")]
    [InlineData("12.5")]
    [InlineData("6")]
    public void PlausibleLookingNumbersDoNotBecomeDates(string value)
    {
        Assert.NotEqual(FieldDataType.Date, Detector().Detect(value).Type);
    }

    [Fact]
    public void ARunOfDigitsIsNeverADate()
    {
        Assert.Equal(FieldDataType.Integer, Detector().Detect("20250131").Type);
    }

    [Fact]
    public void DatesOutsideTheConfiguredRangeAreText()
    {
        var detector = Detector(s =>
        {
            s.MinDateYear = 2000;
            s.MaxDateYear = 2030;
        });

        Assert.Equal(FieldDataType.Text, detector.Detect("1899-05-04").Type);
        Assert.Equal(FieldDataType.Date, detector.Detect("2025-05-04").Type);
    }

    [Fact]
    public void OnlyTheConfiguredDateFormatsAreAccepted()
    {
        var detector = Detector(s => s.DateFormats = ["yyyy-MM-dd"]);

        Assert.Equal(FieldDataType.Date, detector.Detect("2025-01-31").Type);
        Assert.Equal(FieldDataType.Text, detector.Detect("31.01.2025").Type);
    }

    [Fact]
    public void BooleanDetectionIsOffByDefaultAndCanBeEnabled()
    {
        Assert.Equal(FieldDataType.Text, Detector().Detect("yes").Type);

        var detector = Detector(s => s.DetectBooleans = true);
        var result = detector.Detect("yes");

        Assert.Equal(FieldDataType.Boolean, result.Type);
        Assert.True(result.Boolean);
    }

    [Fact]
    public void NumberModeCanBeRestrictedToOneCulture()
    {
        var invariantOnly = Detector(s => s.NumberFormat = NumberFormatMode.Invariant);

        Assert.Equal(FieldDataType.Decimal, invariantOnly.Detect("19.99").Type);
        Assert.Equal(FieldDataType.Text, invariantOnly.Detect("19,99").Type);
    }

    [Fact]
    public void GroupingSeparatorsAreRejectedUnlessEnabled()
    {
        // "1,234,567" can only be read with a grouping separator, so without the
        // setting it stays text. (A lone "1,234" is a German decimal and parses
        // either way, so it is not the right probe for this setting.)
        Assert.Equal(FieldDataType.Text, Detector().Detect("1,234,567").Type);

        var lenient = Detector(s =>
        {
            s.AllowThousandsSeparator = true;
            s.NumberFormat = NumberFormatMode.Invariant;
        });

        Assert.Equal(FieldDataType.Integer, lenient.Detect("1,234,567").Type);
    }

    [Fact]
    public void ValuesWithBothSeparatorsAreNumbersEvenByDefault()
    {
        // "1.234,56" and "1,234.56" hold a dot and a comma, so one of them is
        // grouping and the other the decimal point. There is no other reading,
        // so they are numbers even with the setting off.
        var german = Detector().Detect("1.234,56");
        Assert.Equal(FieldDataType.Decimal, german.Type);
        Assert.Equal(1234.56, german.Number, precision: 6);

        var english = Detector().Detect("1,234.56");
        Assert.Equal(FieldDataType.Decimal, english.Type);
        Assert.Equal(1234.56, english.Number, precision: 6);
    }

    [Fact]
    public void ValuesWithExplicitDecimalPlacesStayDecimalEvenWhenWhole()
    {
        // Keeps a money column on one Excel format instead of alternating.
        Assert.Equal(FieldDataType.Decimal, Detector().Detect("45.00").Type);
    }

    [Fact]
    public void ARuleSpecificDateFormatOverridesTheGlobalList()
    {
        var detector = Detector(s => s.DateFormats = ["yyyy-MM-dd"]);

        Assert.True(detector.TryParseDate("01/31/2025", "MM/dd/yyyy", out var date));
        Assert.Equal(new DateTime(2025, 1, 31), date);
    }
}
