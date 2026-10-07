using DataWizard.Core.Configuration;
using DataWizard.Core.Diff;

namespace DataWizard.Core.Tests;

public class DiffEngineTests
{
    private readonly DiffEngine _engine = new(new DetectionSettings());

    private static LoadedTable Table(string[] header, string[][] rows, string path = "reference.csv") => new()
    {
        SourcePath = path,
        SheetName = null,
        HasHeader = true,
        Header = header,
        Rows = rows,
        Warnings = []
    };

    private static LoadedTable TableNoHeader(string[][] rows, string path = "reference.csv")
    {
        var width = rows.Length > 0 ? rows.Max(r => r.Length) : 0;

        return new LoadedTable
        {
            SourcePath = path,
            SheetName = null,
            HasHeader = false,
            Header = Enumerable.Range(1, width).Select(i => $"Column {i}").ToArray(),
            Rows = rows,
            Warnings = []
        };
    }

    // ── Positional ────────────────────────────────────────────────────────────

    [Fact]
    public void IdenticalTablesHaveNoDifferences()
    {
        var reference = Table(["id", "name"], [["1", "Widget"], ["2", "Gadget"]]);
        var candidate = Table(["id", "name"], [["1", "Widget"], ["2", "Gadget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        Assert.False(report.HasDifferences);
        Assert.Empty(report.Entries);
    }

    [Fact]
    public void AChangedCellIsReportedWithBothValuesAndItsRow()
    {
        var reference = Table(["id", "price"], [["1", "19.99"]]);
        var candidate = Table(["id", "price"], [["1", "24.99"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        var entry = Assert.Single(report.Entries);
        Assert.Equal(DiffKind.ValueChanged, entry.Kind);
        Assert.Equal(1, entry.ReferenceRow);
        Assert.Equal(1, entry.CandidateRow);
        Assert.Equal("price", entry.Column);
        Assert.Equal("19.99", entry.ReferenceValue);
        Assert.Equal("24.99", entry.CandidateValue);
        Assert.Equal(1, report.ValueChangedCount);
    }

    [Fact]
    public void ARowOnlyInTheReferenceIsReported()
    {
        var reference = Table(["id"], [["1"], ["2"]]);
        var candidate = Table(["id"], [["1"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        var entry = Assert.Single(report.Entries);
        Assert.Equal(DiffKind.RowOnlyInReference, entry.Kind);
        Assert.Equal(2, entry.ReferenceRow);
        Assert.Equal(0, entry.CandidateRow);
        Assert.Equal(1, report.RowOnlyInReferenceCount);
    }

    [Fact]
    public void ARowOnlyInTheCandidateIsReported()
    {
        var reference = Table(["id"], [["1"]]);
        var candidate = Table(["id"], [["1"], ["2"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        var entry = Assert.Single(report.Entries);
        Assert.Equal(DiffKind.RowOnlyInCandidate, entry.Kind);
        Assert.Equal(0, entry.ReferenceRow);
        Assert.Equal(2, entry.CandidateRow);
    }

    [Fact]
    public void AColumnOnlyOnOneSideIsReportedStructurally()
    {
        var reference = Table(["id", "extra"], [["1", "x"]]);
        var candidate = Table(["id"], [["1"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        var entry = Assert.Single(report.Entries);
        Assert.Equal(DiffKind.ColumnOnlyInReference, entry.Kind);
        Assert.Equal("extra", entry.Column);
    }

    [Fact]
    public void IgnoredColumnsAreSkippedAndDoNotReportStructuralDifferences()
    {
        var reference = Table(["id", "noise"], [["1", "a"]]);
        var candidate = Table(["id"], [["1"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            IgnoredColumns = ["noise"]
        });

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void IgnoringAColumnByPositionNameWorksForHeaderlessTables()
    {
        var reference = TableNoHeader([["1", "a"], ["2", "b"]]);
        var candidate = TableNoHeader([["1", "x"], ["2", "x"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            IgnoredColumns = ["Column 2"]
        });

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void EmptyTablesHaveNoDifferences()
    {
        var reference = Table(["id"], []);
        var candidate = Table(["id"], [], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        Assert.False(report.HasDifferences);
    }

    // ── Numbers ───────────────────────────────────────────────────────────────

    [Fact]
    public void NumbersWithinTheAbsoluteToleranceAreEqual()
    {
        var reference = Table(["price"], [["100"]]);
        var candidate = Table(["price"], [["100.5"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            AbsoluteTolerance = 1d
        });

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void NumbersBeyondTheAbsoluteToleranceDiffer()
    {
        var reference = Table(["price"], [["100"]]);
        var candidate = Table(["price"], [["101.5"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            AbsoluteTolerance = 1d
        });

        Assert.Single(report.Entries);
    }

    [Fact]
    public void NumbersWithinTheRelativeToleranceAreEqual()
    {
        var reference = Table(["price"], [["1000"]]);
        var candidate = Table(["price"], [["1005"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            RelativeTolerance = 0.01
        });

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void TheSameNumberSpelledInDifferentCulturesIsEqual()
    {
        var reference = Table(["price"], [["19.99"]]);
        var candidate = Table(["price"], [["19,99"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void LeadingZerosAreNeverComparedAsNumbers()
    {
        var reference = Table(["article"], [["00123"]]);
        var candidate = Table(["article"], [["123"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            AbsoluteTolerance = 1000d
        });

        var entry = Assert.Single(report.Entries);
        Assert.Equal(DiffKind.ValueChanged, entry.Kind);
        Assert.Equal("00123", entry.ReferenceValue);
        Assert.Equal("123", entry.CandidateValue);
    }

    // ── Dates ─────────────────────────────────────────────────────────────────

    [Fact]
    public void DatesWithinTheToleranceAreEqual()
    {
        var reference = Table(["date"], [["31.01.2025"]]);
        var candidate = Table(["date"], [["01.02.2025"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            DateTolerance = TimeSpan.FromDays(1)
        });

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void DatesBeyondTheToleranceDiffer()
    {
        var reference = Table(["date"], [["31.01.2025"]]);
        var candidate = Table(["date"], [["01.02.2025"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        Assert.Single(report.Entries);
    }

    [Fact]
    public void AConfiguredDateFormatAndTheIsoFormOfTheSameDateAreEqual()
    {
        var reference = Table(["date"], [["31.01.2025"]]);
        var candidate = Table(["date"], [["2025-01-31T00:00:00.0000000"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        Assert.Empty(report.Entries);
    }

    // ── Text ──────────────────────────────────────────────────────────────────

    [Fact]
    public void CaseIsIgnoredWhenConfigured()
    {
        var reference = Table(["name"], [["Widget"]]);
        var candidate = Table(["name"], [["widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            IgnoreCase = true
        });

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void CaseMattersByDefault()
    {
        var reference = Table(["name"], [["Widget"]]);
        var candidate = Table(["name"], [["widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        Assert.Single(report.Entries);
    }

    [Fact]
    public void SurroundingWhitespaceIsTrimmedByDefault()
    {
        var reference = Table(["name"], [[" Widget "]]);
        var candidate = Table(["name"], [["Widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void WhitespaceIsKeptWhenTrimmingIsOff()
    {
        var reference = Table(["name"], [[" Widget "]]);
        var candidate = Table(["name"], [["Widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            TrimValues = false
        });

        Assert.Single(report.Entries);
    }

    [Fact]
    public void EmptyCellsOnBothSidesAreEqual()
    {
        var reference = Table(["a", "b"], [["1", ""]]);
        var candidate = Table(["a", "b"], [["1", "  "]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        Assert.Empty(report.Entries);
    }

    // ── Keyed ─────────────────────────────────────────────────────────────────

    [Fact]
    public void KeyedComparisonMatchesRowsOutOfOrder()
    {
        var reference = Table(["id", "name"], [["1", "Widget"], ["2", "Gadget"]]);
        var candidate = Table(["id", "name"], [["2", "Gadget"], ["1", "Widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["id"]
        });

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void KeyedComparisonReportsValueChangesWithBothRowNumbers()
    {
        var reference = Table(["id", "name"], [["1", "Widget"], ["2", "Gadget"]]);
        var candidate = Table(["id", "name"], [["2", "Gizmo"], ["1", "Widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["id"]
        });

        var entry = Assert.Single(report.Entries);
        Assert.Equal(DiffKind.ValueChanged, entry.Kind);
        Assert.Equal("name", entry.Column);
        Assert.Equal(2, entry.ReferenceRow);
        Assert.Equal(1, entry.CandidateRow);
    }

    [Fact]
    public void KeyedComparisonReportsARowMissingFromTheCandidateWithItsKey()
    {
        var reference = Table(["id", "name"], [["1", "Widget"], ["2", "Gadget"]]);
        var candidate = Table(["id", "name"], [["1", "Widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["id"]
        });

        var entry = Assert.Single(report.Entries);
        Assert.Equal(DiffKind.RowOnlyInReference, entry.Kind);
        Assert.Equal(2, entry.ReferenceRow);
        Assert.Equal(0, entry.CandidateRow);
        Assert.Contains("2", entry.Detail);
    }

    [Fact]
    public void KeysAreComparedAsTextNeverAsNumbers()
    {
        var reference = Table(["id", "name"], [["1", "Widget"]]);
        var candidate = Table(["id", "name"], [["1.0", "Widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["id"]
        });

        Assert.Equal(2, report.DifferenceCount);
        Assert.Equal(1, report.RowOnlyInReferenceCount);
        Assert.Equal(1, report.RowOnlyInCandidateCount);
    }

    [Fact]
    public void DuplicateKeysArePairedInOrderOfAppearanceAndCounted()
    {
        var reference = Table(["id", "name"], [["1", "a"], ["1", "b"], ["2", "c"]]);
        var candidate = Table(["id", "name"], [["1", "a"], ["1", "x"], ["2", "c"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["id"]
        });

        // One duplicated row in the reference and one in the candidate.
        Assert.Equal(2, report.DuplicateKeyCount);
        var entry = Assert.Single(report.Entries);
        Assert.Equal(DiffKind.ValueChanged, entry.Kind);
        Assert.Equal("b", entry.ReferenceValue);
        Assert.Equal("x", entry.CandidateValue);
        Assert.Contains(report.Warnings, w => w.Contains("occurs 2 times in the reference"));
    }

    [Fact]
    public void AnUnresolvableKeyColumnThrows()
    {
        var reference = Table(["id", "name"], [["1", "Widget"]]);
        var candidate = Table(["id", "name"], [["1", "Widget"]], "candidate.csv");

        Assert.Throws<ArgumentException>(() =>
            _engine.Compare(reference, candidate, new DiffOptions
            {
                MatchMode = DiffMatchMode.Keyed,
                KeyColumns = ["missing"]
            }));
    }

    [Fact]
    public void KeyedComparisonWithoutHeadersUsesPositions()
    {
        var reference = TableNoHeader([["1", "Widget"], ["2", "Gadget"]]);
        var candidate = TableNoHeader([["2", "Gadget"], ["1", "Widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["1"]
        });

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void KeyColumnsAreNotComparedAsCells()
    {
        var reference = Table(["id", "name"], [[" 1 ", "Widget"]]);
        var candidate = Table(["id", "name"], [["1", "Widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["id"]
        });

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void KeyedComparisonPairsColumnsByNameWhenHeadersDifferInOrder()
    {
        var reference = Table(["id", "name", "price"], [["1", "Widget", "10"]]);
        var candidate = Table(["id", "price", "name"], [["1", "10", "Widget"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["id"]
        });

        Assert.Empty(report.Entries);
    }

    [Fact]
    public void KeyedComparisonReportsAColumnMissingFromTheCandidate()
    {
        var reference = Table(["id", "note"], [["1", "hello"]]);
        var candidate = Table(["id"], [["1"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["id"]
        });

        var entry = Assert.Single(report.Entries);
        Assert.Equal(DiffKind.ColumnOnlyInReference, entry.Kind);
        Assert.Equal("note", entry.Column);
    }

    [Fact]
    public void CompositeKeysMatchOnlyWhenAllPartsMatch()
    {
        var reference = Table(["a", "b", "value"], [["1", "x", "10"]]);
        var candidate = Table(["a", "b", "value"], [["1", "y", "10"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate, new DiffOptions
        {
            MatchMode = DiffMatchMode.Keyed,
            KeyColumns = ["a", "b"]
        });

        Assert.Equal(2, report.DifferenceCount);
        Assert.Equal(1, report.RowOnlyInReferenceCount);
        Assert.Equal(1, report.RowOnlyInCandidateCount);
    }

    // ── Report shape ──────────────────────────────────────────────────────────

    [Fact]
    public void TheReportCarriesBothTablesAndTheOptionsUsed()
    {
        var reference = Table(["id"], [["1"]]);
        var candidate = Table(["id"], [["2"]], "candidate.csv");
        var options = new DiffOptions { AbsoluteTolerance = 0.5d };

        var report = _engine.Compare(reference, candidate, options);

        Assert.Equal("reference.csv", report.ReferencePath);
        Assert.Equal("candidate.csv", report.CandidatePath);
        Assert.Equal(0.5d, report.Options.AbsoluteTolerance);
        Assert.Equal(1, report.ReferenceRowCount);
        Assert.Equal(1, report.CandidateRowCount);
        Assert.Equal(1, report.DifferenceCount);
    }

    [Fact]
    public void LoaderWarningsArePropagatedIntoTheReport()
    {
        var reference = new LoadedTable
        {
            SourcePath = "reference.csv",
            SheetName = null,
            HasHeader = true,
            Header = ["id"],
            Rows = [["1"]],
            Warnings = ["Encoding confidence was low."]
        };
        var candidate = Table(["id"], [["1"]], "candidate.csv");

        var report = _engine.Compare(reference, candidate);

        Assert.Contains("Encoding confidence was low.", report.Warnings);
    }
}
