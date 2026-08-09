# DataWizards

**CSV ↔ Excel conversion for .NET, with detection you can inspect and adjust.**

[![.NET](https://img.shields.io/badge/.NET-10.0-blue.svg)](https://dotnet.microsoft.com/)
[![Avalonia](https://img.shields.io/badge/UI-Avalonia%2012-purple.svg)](https://avaloniaui.net/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

DataWizards converts delimited text files to Excel workbooks and back. The
difficult part of that job is not writing the file — it is working out how to read
the one you were given: which encoding, which separator, whether the first row
names the columns, and whether `00123` is a number or an article code.

DataWizards makes every one of those decisions visible, explains the evidence
behind it, and lets you change the rules.

---

## Contents

- [Application](#application)
- [What the analysis panel tells you](#what-the-analysis-panel-tells-you)
- [Detection settings](#detection-settings)
- [Folder watcher](#folder-watcher)
- [Library use](#library-use)
- [Building](#building)
- [Project layout](#project-layout)
- [Troubleshooting](#troubleshooting)

---

## Application

Drop files onto the **Convert** tab, or drop a folder to queue everything
convertible inside it. The direction comes from the extension:

| Extension | Direction |
|---|---|
| `.csv` `.txt` `.log` `.tsv` `.dat` | text → Excel workbook |
| `.xlsx` `.xlsm` | workbook → delimited text |

The interface has a light and a dark theme, selectable in the header or left on
**System** to follow the operating system.

Settings live in `%APPDATA%\DataWizards\settings.json`. They are saved when you
press **Save settings** and again when the application closes. The file is plain
JSON — readable, editable, and safe to copy between machines; values outside a
sensible range are clamped on load.

---

## What the analysis panel tells you

Below the file list, one line reports how the selected file was read — separator,
column count, header, line count, encoding. That is usually all the confirmation
needed.

The full analysis lives in a panel on the right edge, opened with **Details** or
the edge strip and resizable by dragging its border. It **opens by itself only
when the detection was genuinely ambiguous**, and then states why at the top:
a warning was raised, a column holds values of different kinds with no rule to
settle it, or the header score landed within a hair of the threshold — a decision
that could as easily have gone the other way. An amber dot marks the same thing
without opening anything.

Inside, the verdict and the column types are shown immediately; the evidence
behind each decision sits in collapsed sections, present when a file misbehaves
and out of the way when it does not:

**Separator candidates** — every candidate character, scored by how consistently
it produces the same number of columns. The winner is the *most consistent*
candidate, not the most frequent. That distinction matters: a single description
column full of commas contains far more commas than the file has semicolons, and
frequency-based detection picks the wrong one every time.

**Header signals** — the four independent checks behind the header decision, each
with its score and its reasoning, so you can see which one disagreed.

**Columns** — the type detected from the values, the type actually used, and the
rule that overrode it if one did.

**Rows as parsed** — the first rows in aligned columns. A value that has slipped
into the wrong column is obvious at a glance.

Warnings appear here too: an unclosed quote, rows with differing field counts, a
separator that only explains part of the file, an encoding guess that was not
confident.

---

## Detection settings

Everything below is configurable on the **Detection** tab, and every change is
applied immediately to a sample shown beside it.

Selecting a delimited text file on the Convert tab loads its opening lines into
that sample automatically, so the rules can be tuned against real data rather than
an invented example. Typing into the box detaches it from the file, and a new
selection then leaves your text alone; **Use selected file** loads it deliberately
and **Built-in** restores the shipped example.

### Encoding

Detected from the file, or forced. A byte order mark wins by default. When
detection falls below the confidence threshold the configured fallback is used and
the analysis says so rather than silently producing mojibake.

### Separator

Candidate characters, an optional forced separator, the quote character, how many
rows to analyse, and the consistency threshold below which the file is flagged.

### Header detection

In **Auto** mode four signals are scored from 0 to 1, combined using configurable
weights, and compared against a threshold:

| Signal | What it measures | Default weight |
|---|---|---|
| Known field names | how many first-row values match the pattern list | 1.0 |
| Type divergence | columns holding text on row 1 and numbers or dates below | 1.5 |
| Unique values | column names are normally distinct | 0.5 |
| No empty cells | header rows rarely have gaps | 0.5 |

Type divergence carries the highest weight because it is the most reliable signal
for machine-generated exports, and it works even when the column names are ones
DataWizards has never seen.

Set the mode to **Always** or **Never** to skip scoring entirely. Use *skip
leading lines* for files that begin with a banner or comment block.

### Value types

The defaults are chosen to protect identifiers:

- **Leading zeros** — `00123` and `01067` stay text. Turning them into numbers
  drops the zeros, which is what mangles article numbers and postal codes.
- **Long digit runs** — anything past the digit limit (15 by default) stays text.
  Excel stores numbers as doubles, so beyond roughly fifteen digits the value
  silently changes. That is what rounds EAN codes and IBANs into nonsense.
- **Quoted values** — a value the source wrote in quotes is text, because the
  quotes are the author stating that the characters matter.
- **Dates** — parsed only against the configured format list, never against the
  machine's regional settings, so the same file reads identically everywhere.
- **Numbers** — Auto accepts `1234.56` first and then the configured culture, so
  `19.99` and `19,99` both work. Thousands separators are off by default because
  they make `1.234` ambiguous between one thousand and a decimal.

### Data type rules

A rule forces a column to a type regardless of what its values look like. Rules
match on the column name — *Contains*, *Exact*, *StartsWith*, *EndsWith* or
*Regex* — or on a column position for files without a header. The first match
wins, and a positional rule beats a name rule. An invalid regular expression never
matches rather than throwing, so one broken rule cannot stop a batch.

Pattern lists from a version before 0.2 can be brought across with **Import legacy
XML**, which reads the old `DataWizard.config.xml`.

---

## Folder watcher

Each rule watches a folder and converts what appears in it.

- Files are converted **one at a time**. They arrive in bursts, and converting a
  folder's worth in parallel only makes the disk thrash.
- A file is only converted once it has **stopped changing** for the settle delay.
  Large files arrive in pieces, and reading one mid-copy produces truncated data.
- Failures are **retried**, then optionally moved to an error folder.
- Sources can be kept, moved to a processed folder, or deleted after conversion.
- Results written by the watcher are **ignored for two minutes**, so a rule cannot
  feed on its own output. Watching both directions with the output in the watched
  folder is flagged as a warning — a separate output folder is safer.

---

## Library use

`DataWizard.Core` has no UI dependency and can be referenced on its own.

```csharp
using DataWizard.Core.Configuration;
using DataWizard.Core.Conversion;

var settings = DataWizardSettings.CreateDefault();
settings.Conversion.Overwrite = true;

var service = new ConversionService(settings);
service.Log += (_, entry) => Console.WriteLine(entry);

var result = service.Convert(@"C:\data\customers.csv");

if (result.Success)
    Console.WriteLine($"Wrote {result.OutputPaths[0]} ({result.RowCount} rows)");
else
    Console.WriteLine($"Failed: {result.ErrorMessage}");
```

### Inspecting a file without converting it

```csharp
using DataWizard.Core.Csv;

var analysis = new CsvAnalyzer(settings.Detection).Analyze(@"C:\data\customers.csv");

Console.WriteLine($"Separator: {analysis.Separator} ({analysis.SeparatorConfidence:P0})");
Console.WriteLine($"Header: {analysis.HeaderResult.Describe()}");

foreach (var column in analysis.Columns)
    Console.WriteLine($"  {column.Describe()}");

foreach (var warning in analysis.Warnings)
    Console.WriteLine($"  warning: {warning}");
```

### Forcing a column to a type

```csharp
settings.Detection.FieldRules.Insert(0, new FieldRule
{
    Pattern = "^article_no$",
    Match = MatchMode.Regex,
    DataType = FieldDataType.Text
});
```

### Watching a folder

```csharp
using DataWizard.Core.Watching;

var watcher = new FolderWatchService(service);
watcher.FileConverted += (_, r) => Console.WriteLine(r.Describe());

watcher.Start(new WatcherSettings
{
    Rules =
    [
        new WatchRule
        {
            Name = "Incoming orders",
            InputFolder = @"C:\incoming",
            OutputFolder = @"C:\processed",
            FilePatterns = "*.csv",
            SourceAction = SourceFileAction.Move,
            ProcessedFolder = @"C:\archive"
        }
    ]
});
```

---

## Building

Requires the **.NET 10 SDK**. No other prerequisites — everything else restores
from NuGet.

```bash
git clone https://github.com/Felteshobbies/DataWizards.git
cd DataWizards

dotnet build
dotnet test
dotnet run --project src/DataWizard.App
```

### Publishing a standalone executable

```bash
dotnet publish src/DataWizard.App -p:PublishProfile=win-x64
```

Produces a single `DataWizard.exe` in
`src/DataWizard.App/bin/publish/win-x64/` — roughly 71 MB, with the .NET runtime,
the Avalonia native libraries and all dependencies inside it. Nothing needs to be
installed on the target machine; copy the file and run it.

An ARM64 profile is available as `-p:PublishProfile=win-arm64`. Publishing with a
bare runtime identifier works too and applies the same settings:

```bash
dotnet publish src/DataWizard.App -r win-x64 -c Release
```

A plain `dotnet build` stays runtime-agnostic, so day-to-day builds and tests are
not slowed down by any of this.

**On the size.** Most of it is the .NET runtime plus Skia, which Avalonia renders
with. Trimming would cut it substantially but is deliberately off: Avalonia's view
locator, the MVVM toolkit and DocumentFormat.OpenXml all resolve types by
reflection, so a trimmed build fails at runtime rather than at build time — and
only on the code path that happens to need the removed type. If you want to try
it, set `PublishTrimmed` in the profile and test every tab.

Framework-dependent output is far smaller if the target machines already have the
.NET 10 desktop runtime:

```bash
dotnet publish src/DataWizard.App -r win-x64 --self-contained false
```

---

## Project layout

```
DataWizards/
├── src/
│   ├── DataWizard.Core/          Conversion engine, no UI dependency
│   │   ├── Configuration/        Settings, patterns, rules, persistence
│   │   ├── Csv/                  RFC 4180 reader, encoding, separator,
│   │   │                         header and value type detection
│   │   ├── Excel/                Streaming XLSX reader and writer
│   │   ├── Conversion/           Orchestration and logging
│   │   └── Watching/             Folder monitoring
│   └── DataWizard.App/           Avalonia desktop application
│       ├── Services/             Session, dialogs
│       ├── ViewModels/
│       ├── Views/
│       └── Styles/               Light and dark colour tokens
├── tests/
│   ├── DataWizard.Core.Tests/    Engine tests
│   └── DataWizard.App.Tests/     Headless UI smoke tests
└── samples/                      Files exercising the awkward cases
```

### Dependencies

| Package | Used for |
|---|---|
| DocumentFormat.OpenXml 3.5 | reading and writing XLSX |
| UTF.Unknown 2.6 | encoding detection |
| Avalonia 12.1 | desktop UI |
| CommunityToolkit.Mvvm 8.4 | view model plumbing |

---

## Troubleshooting

**Accented characters come out wrong.** Detection was not confident enough and the
fallback was used; the analysis panel says so. Set the encoding explicitly —
`windows-1252` for files from older Windows software.

**Everything lands in one column.** No candidate separator fitted. Check the
separator candidate table; if the file uses something unusual, add that character
to the candidate list or force it.

**The first row was treated as data, or the other way round.** Look at the header
signals to see which check disagreed, then adjust that signal's weight, move the
threshold, or set the mode to Always or Never.

**Leading zeros disappeared.** Check that *keep leading zeros as text* is on. If
the column mixes values with and without zeros, add a rule forcing it to Text.

**Excel opens the CSV as gibberish.** Excel needs a byte order mark to recognise a
UTF-8 CSV on double-click. Turn on *write a byte order mark* on the Output tab, or
switch the output encoding to `windows-1252`.

**Decimals were rounded.** The *max. decimals* setting rounds the value. Leave it
at 15 to keep everything a number can carry.

The **Log** tab records every decision per file. Turn on *Details* to see the full
analysis of each conversion.

---

## License

MIT — see [LICENSE](LICENSE).

## Contributing

Pull requests are welcome; see [CONTRIBUTING.md](CONTRIBUTING.md).

## Acknowledgements

- [DocumentFormat.OpenXml](https://github.com/OfficeDev/Open-XML-SDK)
- [UTF.Unknown](https://github.com/CharsetDetector/UTF-unknown)
- [Avalonia](https://github.com/AvaloniaUI/Avalonia)
