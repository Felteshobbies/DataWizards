namespace DataWizard.Core.Diff;

/// <summary>
/// One table fully read into memory: a header row (real or generated) and the
/// data rows beneath it, with every cell already normalised to a comparable
/// string.
/// </summary>
/// <remarks>
/// Numbers come out invariant (<c>19.99</c>), date-styled cells as ISO
/// timestamps, and booleans as <c>TRUE</c>/<c>FALSE</c>, so a value loaded from a
/// workbook and the same value loaded from a CSV compare equal regardless of the
/// culture each file was written in.
/// </remarks>
public sealed class LoadedTable
{
    /// <summary>The file the table was read from.</summary>
    public required string SourcePath { get; init; }

    /// <summary>The worksheet it came from, for workbooks; <c>null</c> for delimited text.</summary>
    public string? SheetName { get; init; }

    /// <summary>Whether the first row names the columns.</summary>
    public required bool HasHeader { get; init; }

    /// <summary>
    /// Column names: the header values when one was found, otherwise generated
    /// names such as <c>Column 1</c>. Always as wide as the widest row.
    /// </summary>
    public required string[] Header { get; init; }

    /// <summary>The data rows, without the header row.</summary>
    public required string[][] Rows { get; init; }

    /// <summary>Things worth telling the user about, for example a low-confidence encoding.</summary>
    public required IReadOnlyList<string> Warnings { get; init; }

    /// <summary>How many data rows the table holds.</summary>
    public int RowCount => Rows.Length;

    /// <summary>How many columns the table has.</summary>
    public int ColumnCount => Header.Length;
}
