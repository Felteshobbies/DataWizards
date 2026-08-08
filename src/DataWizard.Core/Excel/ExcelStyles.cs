using DocumentFormat.OpenXml.Spreadsheet;
using DataWizard.Core.Configuration;

namespace DataWizard.Core.Excel;

/// <summary>
/// The fixed set of cell formats written into every generated workbook.
/// </summary>
/// <remarks>
/// The indices are part of the file format: a cell refers to its format by
/// position in <c>cellXfs</c>. They are named here rather than written as bare
/// numbers at the call sites.
/// </remarks>
public static class ExcelStyles
{
    /// <summary>General format, used for text.</summary>
    public const uint Text = 0;

    /// <summary>Date format.</summary>
    public const uint Date = 1;

    /// <summary>Decimal format.</summary>
    public const uint Decimal = 2;

    /// <summary>Whole number format.</summary>
    public const uint Integer = 3;

    /// <summary>Bold, used for the header row.</summary>
    public const uint Header = 4;

    private const uint DateFormatId = 164;
    private const uint DecimalFormatId = 165;
    private const uint IntegerFormatId = 166;

    /// <summary>
    /// Builds the stylesheet. The number format codes come from the conversion
    /// settings, so date and decimal presentation is configurable.
    /// </summary>
    public static Stylesheet Create(ConversionSettings settings)
    {
        var stylesheet = new Stylesheet
        {
            NumberingFormats = new NumberingFormats(
                new NumberingFormat
                {
                    NumberFormatId = DateFormatId,
                    FormatCode = Escape(settings.DateNumberFormat, "yyyy\\-mm\\-dd")
                },
                new NumberingFormat
                {
                    NumberFormatId = DecimalFormatId,
                    FormatCode = Escape(settings.DecimalNumberFormat, "#,##0.00")
                },
                new NumberingFormat
                {
                    NumberFormatId = IntegerFormatId,
                    FormatCode = Escape(settings.IntegerNumberFormat, "0")
                }),

            Fonts = new Fonts(
                new Font(
                    new FontSize { Val = 11D },
                    new FontName { Val = "Calibri" }),
                new Font(
                    new Bold(),
                    new FontSize { Val = 11D },
                    new FontName { Val = "Calibri" })),

            Fills = new Fills(
                new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } },
                new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }),

            Borders = new Borders(new Border()),

            CellStyleFormats = new CellStyleFormats(
                new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 })
        };

        stylesheet.CellFormats = new CellFormats(
            // 0 - text / general
            new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 },
            // 1 - date
            new CellFormat
            {
                NumberFormatId = DateFormatId,
                FontId = 0, FillId = 0, BorderId = 0, FormatId = 0,
                ApplyNumberFormat = true
            },
            // 2 - decimal
            new CellFormat
            {
                NumberFormatId = DecimalFormatId,
                FontId = 0, FillId = 0, BorderId = 0, FormatId = 0,
                ApplyNumberFormat = true
            },
            // 3 - integer
            new CellFormat
            {
                NumberFormatId = IntegerFormatId,
                FontId = 0, FillId = 0, BorderId = 0, FormatId = 0,
                ApplyNumberFormat = true
            },
            // 4 - bold header
            new CellFormat
            {
                NumberFormatId = 0,
                FontId = 1, FillId = 0, BorderId = 0, FormatId = 0,
                ApplyFont = true
            });

        stylesheet.NumberingFormats.Count = 3;
        stylesheet.Fonts.Count = 2;
        stylesheet.Fills.Count = 2;
        stylesheet.Borders.Count = 1;
        stylesheet.CellStyleFormats.Count = 1;
        stylesheet.CellFormats.Count = 5;

        return stylesheet;
    }

    /// <summary>Maps a detected data type onto a cell format index.</summary>
    public static uint ForDataType(FieldDataType type) => type switch
    {
        FieldDataType.Date => Date,
        FieldDataType.Decimal => Decimal,
        FieldDataType.Integer => Integer,
        _ => Text
    };

    private static string Escape(string? code, string fallback) =>
        string.IsNullOrWhiteSpace(code) ? fallback : code;

    /// <summary>
    /// Number format identifiers that Excel treats as dates. Used when reading a
    /// workbook to tell a date apart from the plain number it is stored as.
    /// </summary>
    public static bool IsDateFormatId(uint numberFormatId) =>
        numberFormatId is (>= 14 and <= 22) or (>= 27 and <= 36) or (>= 45 and <= 47)
            or (>= 50 and <= 58) or DateFormatId;

    /// <summary>
    /// Recognises a date from a custom format code, for workbooks written by other
    /// tools whose format identifiers are not in the built-in range.
    /// </summary>
    public static bool IsDateFormatCode(string? formatCode)
    {
        if (string.IsNullOrEmpty(formatCode))
            return false;

        var stripped = new string(formatCode
            .Where(c => c is not ('"' or '\\' or '[' or ']'))
            .ToArray())
            .ToLowerInvariant();

        // A format is a date format if it mentions date or time parts. Exclude
        // codes that only use 'm' as a thousands/decimal artefact by requiring a
        // recognisable date or time token.
        return stripped.Contains("yy")
               || stripped.Contains("dd")
               || stripped.Contains("mmm")
               || stripped.Contains("hh")
               || stripped.Contains("ss")
               || stripped.Contains("am/pm");
    }
}
