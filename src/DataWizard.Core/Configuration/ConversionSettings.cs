using System.Globalization;
using System.Text;

namespace DataWizard.Core.Configuration;

/// <summary>
/// Output options for both conversion directions, plus the shared file handling
/// rules (where results are written and whether they may replace existing files).
/// </summary>
public sealed class ConversionSettings
{
    // ── Shared ──────────────────────────────────────────────────────────────

    /// <summary>Target folder. Empty means "next to the source file".</summary>
    public string OutputFolder { get; set; } = string.Empty;

    /// <summary>Replace an existing output file instead of adding a numeric suffix.</summary>
    public bool Overwrite { get; set; }

    // ── Excel to CSV ────────────────────────────────────────────────────────

    /// <summary>Separator token written between fields.</summary>
    public string CsvSeparator { get; set; } = "Semicolon";

    /// <summary>Web name of the encoding used for the CSV file.</summary>
    public string CsvEncoding { get; set; } = "utf-8";

    /// <summary>
    /// Emit a byte order mark. Excel needs one to open a UTF-8 CSV with the right
    /// encoding by double-click; most import scripts prefer none.
    /// </summary>
    public bool WriteByteOrderMark { get; set; } = true;

    /// <summary>How aggressively fields are quoted.</summary>
    public QuoteMode QuoteMode { get; set; } = QuoteMode.TextFields;

    /// <summary>Line terminator.</summary>
    public LineEndingStyle LineEnding { get; set; } = LineEndingStyle.Crlf;

    /// <summary>Write every worksheet to its own file.</summary>
    public bool ExportAllSheets { get; set; }

    /// <summary>Zero-based worksheet index used when <see cref="ExportAllSheets"/> is false.</summary>
    public int WorksheetIndex { get; set; }

    /// <summary>
    /// Culture for numbers in the CSV output. Empty means invariant
    /// (<c>1234.56</c>); <c>de-DE</c> produces <c>1234,56</c>.
    /// </summary>
    public string NumberOutputCulture { get; set; } = string.Empty;

    /// <summary>Format applied to date cells.</summary>
    public string DateOutputFormat { get; set; } = "yyyy-MM-dd";

    /// <summary>
    /// Decimal places kept in the output. The default preserves the value as
    /// stored; lowering it rounds and therefore loses data.
    /// </summary>
    public int MaxDecimalPlaces { get; set; } = 15;

    /// <summary>Write a line for worksheet rows that contain no cells at all.</summary>
    public bool KeepEmptyRows { get; set; } = true;

    // ── CSV to Excel ────────────────────────────────────────────────────────

    /// <summary>Name given to the generated worksheet.</summary>
    public string SheetName { get; set; } = "Sheet1";

    /// <summary>Render the header row in bold.</summary>
    public bool BoldHeaderRow { get; set; } = true;

    /// <summary>Freeze the header row so it stays visible while scrolling.</summary>
    public bool FreezeHeaderRow { get; set; } = true;

    /// <summary>Add an auto-filter to the header row.</summary>
    public bool AutoFilterHeaderRow { get; set; }

    /// <summary>Size columns to their content.</summary>
    public bool AutoFitColumns { get; set; } = true;

    /// <summary>Excel number format code applied to date cells.</summary>
    public string DateNumberFormat { get; set; } = "yyyy\\-mm\\-dd";

    /// <summary>Excel number format code applied to decimal cells.</summary>
    public string DecimalNumberFormat { get; set; } = "#,##0.00";

    /// <summary>Excel number format code applied to whole-number cells.</summary>
    public string IntegerNumberFormat { get; set; } = "0";

    // ── Helpers ─────────────────────────────────────────────────────────────

    /// <summary>The CSV separator as a character, defaulting to <c>;</c>.</summary>
    public char ResolveCsvSeparator() => SeparatorToken.Parse(CsvSeparator) ?? ';';

    /// <summary>
    /// Builds the output encoding. The byte order mark is controlled explicitly
    /// rather than inherited from a shared <see cref="Encoding"/> instance.
    /// </summary>
    public Encoding ResolveCsvEncoding()
    {
        var name = string.IsNullOrWhiteSpace(CsvEncoding) ? "utf-8" : CsvEncoding.Trim();

        Encoding encoding;
        try
        {
            encoding = Encoding.GetEncoding(name);
        }
        catch (ArgumentException)
        {
            encoding = new UTF8Encoding(WriteByteOrderMark);
        }

        // Encoding.UTF8 always emits a preamble, so build the variant we want.
        if (encoding.CodePage == Encoding.UTF8.CodePage)
            return new UTF8Encoding(WriteByteOrderMark);

        if (encoding.CodePage == Encoding.Unicode.CodePage)
            return new UnicodeEncoding(bigEndian: false, byteOrderMark: WriteByteOrderMark);

        return encoding;
    }

    /// <summary>The literal line terminator for <see cref="LineEnding"/>.</summary>
    public string ResolveLineEnding() => LineEnding switch
    {
        LineEndingStyle.Crlf => "\r\n",
        LineEndingStyle.Lf => "\n",
        _ => Environment.NewLine
    };

    /// <summary>Culture used to format numbers on the way out.</summary>
    public CultureInfo ResolveNumberOutputCulture()
    {
        if (string.IsNullOrWhiteSpace(NumberOutputCulture))
            return CultureInfo.InvariantCulture;

        try
        {
            return CultureInfo.GetCultureInfo(NumberOutputCulture);
        }
        catch (CultureNotFoundException)
        {
            return CultureInfo.InvariantCulture;
        }
    }

    /// <summary>Clamps values that would otherwise produce invalid output.</summary>
    public void Normalize()
    {
        MaxDecimalPlaces = Math.Clamp(MaxDecimalPlaces, 0, 15);
        WorksheetIndex = Math.Max(0, WorksheetIndex);

        if (string.IsNullOrWhiteSpace(SheetName))
            SheetName = "Sheet1";

        if (string.IsNullOrWhiteSpace(DateOutputFormat))
            DateOutputFormat = "yyyy-MM-dd";
    }

    /// <summary>Creates a copy for editing.</summary>
    public ConversionSettings Clone() => (ConversionSettings)MemberwiseClone();
}
