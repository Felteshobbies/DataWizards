using DataWizard.Core.Configuration;
using DataWizard.Core.Csv;

namespace DataWizard.Core.Tests;

/// <summary>
/// The shipped rule set, checked against the column names it is meant to catch and
/// the ones it must leave alone.
/// </summary>
public class DefaultPatternTests
{
    private static readonly DetectionSettings Settings = CreateSettings();

    private static DetectionSettings CreateSettings()
    {
        var settings = DataWizardSettings.CreateDefault().Detection;
        settings.Normalize();
        return settings;
    }

    /// <summary>Resolves a column name the way the analyser does.</summary>
    private static FieldDataType? RuleTypeFor(string fieldName)
    {
        var analyzer = new CsvAnalyzer(Settings);
        return analyzer.FindRule(fieldName, columnIndex: 0)?.DataType;
    }

    private static bool IsKnownHeaderName(string fieldName) =>
        Settings.KnownFieldNames.Any(p => p.IsMatch(fieldName));

    // ── Identifiers that must stay text ─────────────────────────────────────

    [Theory]
    [InlineData("id")]
    [InlineData("customer_id")]
    [InlineData("kundennummer")]
    [InlineData("customerno")]
    [InlineData("sku")]
    [InlineData("ean")]
    [InlineData("iban")]
    [InlineData("plz")]
    [InlineData("postleitzahl")]
    [InlineData("zip")]
    [InlineData("telefon")]
    public void IdentifierColumnsAreForcedToText(string name)
    {
        Assert.Equal(FieldDataType.Text, RuleTypeFor(name));
    }

    /// <summary>
    /// These came from the pre-0.2 configuration, where they had evidently been
    /// added after real data went wrong.
    /// </summary>
    [Theory]
    [InlineData("artikel")]
    [InlineData("artikelnr")]
    [InlineData("artikelnummer")]
    [InlineData("artikel_nr")]
    [InlineData("article")]
    [InlineData("article_no")]
    [InlineData("part")]
    [InlineData("partno")]
    [InlineData("part-nr")]
    [InlineData("teil")]
    [InlineData("teilnummer")]
    [InlineData("teilenummer")]
    [InlineData("matnr")]
    [InlineData("material_nr")]
    [InlineData("materialnummer")]
    public void ArticlePartAndMaterialNumbersAreForcedToText(string name)
    {
        Assert.Equal(FieldDataType.Text, RuleTypeFor(name));
    }

    /// <summary>
    /// The original patterns were <c>.*artikel.*</c> and <c>.*teil.*</c>. Copied
    /// verbatim they would force these numeric columns to text, which is worse
    /// than having no rule at all.
    /// </summary>
    [Theory]
    [InlineData("artikelpreis", FieldDataType.Decimal)]
    [InlineData("artikel_preis", FieldDataType.Decimal)]
    [InlineData("article_price", FieldDataType.Decimal)]
    [InlineData("artikelmenge", FieldDataType.Integer)]
    public void TheLooseArticlePatternsDoNotSwallowPricesAndQuantities(string name, FieldDataType expected)
    {
        Assert.Equal(expected, RuleTypeFor(name));
    }

    [Theory]
    [InlineData("anteil")]      // "share" - a number, matched by the old .*teil.*
    [InlineData("verteiler")]
    [InlineData("urteil")]
    [InlineData("bestellteilmenge")]
    public void WordsMerelyContainingTeilAreLeftAlone(string name)
    {
        Assert.NotEqual(FieldDataType.Text, RuleTypeFor(name) ?? FieldDataType.Auto);
    }

    /// <summary>
    /// The old header list used <c>.*id.*</c>, which matched almost any text and
    /// made nearly every first row look like a header.
    /// </summary>
    [Theory]
    [InlineData("bildname")]
    [InlineData("guide")]
    [InlineData("identisch")]
    [InlineData("hybrid")]
    public void WordsMerelyContainingIdAreNotIdentifiers(string name)
    {
        Assert.NotEqual(FieldDataType.Text, RuleTypeFor(name) ?? FieldDataType.Auto);
    }

    // ── Typed columns ───────────────────────────────────────────────────────

    [Theory]
    [InlineData("preis", FieldDataType.Decimal)]
    [InlineData("price", FieldDataType.Decimal)]
    [InlineData("betrag", FieldDataType.Decimal)]
    [InlineData("amount", FieldDataType.Decimal)]
    [InlineData("kosten", FieldDataType.Decimal)]
    [InlineData("menge", FieldDataType.Integer)]
    [InlineData("quantity", FieldDataType.Integer)]
    [InlineData("anzahl", FieldDataType.Integer)]
    [InlineData("qty", FieldDataType.Integer)]
    [InlineData("datum", FieldDataType.Date)]
    [InlineData("date", FieldDataType.Date)]
    [InlineData("geburtsdatum", FieldDataType.Date)]
    [InlineData("birthdate", FieldDataType.Date)]
    public void TypedColumnsKeepTheTypeTheOldConfigurationGaveThem(string name, FieldDataType expected)
    {
        Assert.Equal(expected, RuleTypeFor(name));
    }

    // ── Header names ────────────────────────────────────────────────────────

    [Theory]
    [InlineData("id")]
    [InlineData("artikel")]
    [InlineData("produkt")]
    [InlineData("bezeichnung")]
    [InlineData("preis")]
    [InlineData("datum")]
    [InlineData("strasse")]
    [InlineData("ort")]
    [InlineData("plz")]
    [InlineData("email")]
    [InlineData("e-mail")]
    [InlineData("mail")]
    [InlineData("telefon")]
    [InlineData("menge")]
    [InlineData("anzahl")]
    [InlineData("count")]
    [InlineData("matnr")]
    [InlineData("teilnummer")]
    public void EveryHeaderNameFromTheOldConfigurationIsStillRecognised(string name)
    {
        Assert.True(IsKnownHeaderName(name), $"'{name}' is no longer a known header name");
    }

    [Theory]
    [InlineData("department")]
    [InlineData("abteilung")]
    [InlineData("participant")]
    public void HeaderPatternsDoNotFireOnWordsThatMerelyContainThem(string name)
    {
        Assert.False(IsKnownHeaderName(name), $"'{name}' should not count as a known header name");
    }

    [Fact]
    public void EveryShippedPatternCompiles()
    {
        foreach (var pattern in Settings.KnownFieldNames)
        {
            Assert.True(pattern.IsPatternValid(out var error),
                $"header pattern '{pattern.Pattern}' is invalid: {error}");
        }

        foreach (var rule in Settings.FieldRules)
        {
            Assert.True(rule.IsPatternValid(out var error),
                $"field rule '{rule.Pattern}' is invalid: {error}");
        }
    }

    /// <summary>
    /// Rules are first-match-wins, so a rule that is shadowed by an earlier one can
    /// never fire and is a latent mistake.
    /// </summary>
    [Fact]
    public void NoShippedRuleIsCompletelyShadowedByAnEarlierOne()
    {
        var rules = Settings.FieldRules;

        for (var i = 1; i < rules.Count; i++)
        {
            // Only literal Contains rules can be checked cheaply for shadowing.
            if (rules[i].Match != MatchMode.Contains)
                continue;

            var sample = rules[i].Pattern;
            var winner = rules.FirstOrDefault(r => r.AppliesTo(sample, 0));

            Assert.True(
                winner is not null && winner.DataType == rules[i].DataType,
                $"rule '{rules[i].Pattern}' -> {rules[i].DataType} is shadowed by " +
                $"'{winner?.Pattern}' -> {winner?.DataType}");
        }
    }
}
