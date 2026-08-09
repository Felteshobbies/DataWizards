namespace DataWizard.Core.Configuration;

/// <summary>
/// The patterns shipped with the application. They cover common English and
/// German column names found in exports from ERP and shop systems.
/// </summary>
public static class DefaultPatterns
{
    private static NamePattern Word(string pattern, string? comment = null) => new()
    {
        Pattern = pattern,
        Match = MatchMode.Regex,
        Comment = comment
    };

    private static NamePattern Text(string pattern, string? comment = null) => new()
    {
        Pattern = pattern,
        Match = MatchMode.Contains,
        Comment = comment
    };

    /// <summary>
    /// Column names that identify a row as a header.
    /// </summary>
    /// <remarks>
    /// The identifier patterns are anchored on word boundaries on purpose. The
    /// previous <c>.*id.*</c> also matched "Bildname", "Identity" and "Guide",
    /// which made almost any text row look like a header.
    /// </remarks>
    public static List<NamePattern> CreateKnownFieldNames() =>
    [
        // Identifiers and numbers
        Word(@"^id$", "Exactly 'id'"),
        Word(@"\bid\b", "'id' as a separate word"),
        Word(@"(^|[_\s-])(id|nr|no|num)([_\s-]|$)", "id/nr/no as a name part"),
        Word(@"(number|nummer)$"),
        Word(@"^(pos|position)$"),

        // Article and product
        Text("article"),
        Text("artikel"),
        Text("product"),
        Text("produkt"),
        Text("sku"),
        Text("ean"),
        Text("gtin"),
        Word(@"\bpart([\s_-]?(no|nr|number))?\b", "part / part-no / partnumber, but not 'department'"),
        Word(@"teile?[\s_-]?(nr|no|nummer|number)", "teilnummer and teilenummer"),
        Word(@"^mat(erial)?[\s_-]?(nr|no|nummer|number)$", "Material number"),

        // Names and descriptions
        Word(@"^name$"),
        Word(@"(first|last|full|company|customer|supplier)[\s_-]?name"),
        Text("bezeichnung"),
        Text("description"),
        Text("beschreibung"),
        Word(@"^titl?e$"),
        Text("titel"),

        // Money
        Text("price"),
        Text("preis"),
        Text("amount"),
        Text("betrag"),
        Text("cost"),
        Text("kosten"),
        Text("currency"),
        Text("waehrung"),
        Text("währung"),
        Word(@"(net|gross|netto|brutto)"),
        Word(@"(tax|vat|mwst|ust)"),
        Text("discount"),
        Text("rabatt"),

        // Dates
        Text("date"),
        Text("datum"),
        Word(@"^(from|to|von|bis)$"),
        Word(@"(created|modified|updated|erstellt|geaendert|geändert)"),
        Text("timestamp"),
        Text("zeitstempel"),

        // Addresses
        Text("street"),
        Text("strasse"),
        Text("straße"),
        Text("city"),
        Word(@"^ort$"),
        Text("country"),
        Text("land"),
        Text("zip"),
        Text("plz"),
        Text("postal"),
        Text("postleitzahl"),
        Text("address"),
        Text("adresse"),

        // Contact
        Text("email"),
        Text("e-mail"),
        Word(@"^mail$"),
        Text("phone"),
        Text("telefon"),
        Word(@"^(tel|fax|mobile|mobil)$"),

        // Quantities
        Text("quantity"),
        Text("menge"),
        Text("anzahl"),
        Word(@"^count$"),
        Word(@"^(qty|stk|pcs)$"),
        Text("weight"),
        Text("gewicht"),
        Text("unit"),
        Text("einheit"),

        // Status
        Text("status"),
        Text("state"),
        Word(@"^(active|aktiv)$"),
        Text("category"),
        Text("kategorie")
    ];

    private static FieldRule Rule(string pattern, MatchMode match, FieldDataType type, string? comment = null) => new()
    {
        Pattern = pattern,
        Match = match,
        DataType = type,
        Comment = comment
    };

    /// <summary>
    /// Type overrides for columns whose correct type cannot be told from the
    /// values alone. Identifiers and postal codes are the classic case: they look
    /// like numbers, but turning them into numbers destroys leading zeros.
    /// </summary>
    public static List<FieldRule> CreateFieldRules() =>
    [
        // Identifiers stay text so leading zeros survive
        Rule(@"^id$", MatchMode.Regex, FieldDataType.Text, "Leading zeros must survive"),
        Rule(@"(^|[_\s-])id$", MatchMode.Regex, FieldDataType.Text),
        Rule(@"(number|nummer|_no|_nr)$", MatchMode.Regex, FieldDataType.Text),
        Rule("customerno", MatchMode.Contains, FieldDataType.Text),
        Rule("kundennummer", MatchMode.Contains, FieldDataType.Text),

        // Article, part and material numbers. Carried over from the pre-0.2
        // configuration, where they had clearly been added in response to real
        // data, but anchored rather than copied verbatim: the originals were
        // ".*artikel.*" and ".*teil.*", which also match "Artikelpreis" and
        // "Anteil" and would turn those numeric columns into text.
        Rule(@"^(artikel|article)([\s_-]?(nr|no|num|nummer|number))?$", MatchMode.Regex, FieldDataType.Text,
            "Article number - leading zeros must survive"),
        Rule(@"^(teil|teile|part)([\s_-]?(nr|no|num|nummer|number))?$", MatchMode.Regex, FieldDataType.Text,
            "Part number"),
        Rule(@"^mat(erial)?[\s_-]?(nr|no|num|nummer|number)$", MatchMode.Regex, FieldDataType.Text,
            "Material number, as used by SAP and similar systems"),

        Rule("sku", MatchMode.Contains, FieldDataType.Text),
        Rule("ean", MatchMode.Contains, FieldDataType.Text, "13 digits - too long for a double"),
        Rule("gtin", MatchMode.Contains, FieldDataType.Text),
        Rule("iban", MatchMode.Contains, FieldDataType.Text),
        Rule("bic", MatchMode.Contains, FieldDataType.Text),

        // Postal codes and phone numbers are text, never numbers
        Rule("plz", MatchMode.Contains, FieldDataType.Text),
        Rule("postleitzahl", MatchMode.Contains, FieldDataType.Text),
        Rule("zip", MatchMode.Contains, FieldDataType.Text),
        Rule("postal", MatchMode.Contains, FieldDataType.Text),
        Rule("phone", MatchMode.Contains, FieldDataType.Text),
        Rule("telefon", MatchMode.Contains, FieldDataType.Text),
        Rule(@"^(tel|fax|mobile|mobil)$", MatchMode.Regex, FieldDataType.Text),

        // Money
        Rule("price", MatchMode.Contains, FieldDataType.Decimal),
        Rule("preis", MatchMode.Contains, FieldDataType.Decimal),
        Rule("amount", MatchMode.Contains, FieldDataType.Decimal),
        Rule("betrag", MatchMode.Contains, FieldDataType.Decimal),
        Rule("cost", MatchMode.Contains, FieldDataType.Decimal),
        Rule("kosten", MatchMode.Contains, FieldDataType.Decimal),

        // Quantities
        Rule("quantity", MatchMode.Contains, FieldDataType.Integer),
        Rule("menge", MatchMode.Contains, FieldDataType.Integer),
        Rule("anzahl", MatchMode.Contains, FieldDataType.Integer),
        Rule(@"^(qty|stk|pcs)$", MatchMode.Regex, FieldDataType.Integer),

        // Dates
        Rule("date", MatchMode.Contains, FieldDataType.Date),
        Rule("datum", MatchMode.Contains, FieldDataType.Date)
    ];
}
