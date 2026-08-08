# Changelog

All notable changes to DataWizard will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2026-08-08

A rebuild rather than a revision. The conversion engine was rewritten, the
WinForms interface was replaced with an Avalonia one, and the whole application is
now in English.

> **Versioning note:** the previous entry was tagged 1.0.0 while the shipped
> assembly version was 0.1.1 and the interface was incomplete. Numbering restarts
> from 0.2.0 to match what actually exists.

### Breaking

- **Target framework** is now .NET 10. The .NET Framework 4.7.2 build is gone.
- **Projects renamed and moved.** `libDataWizard` → `src/DataWizard.Core`,
  `DataWizard.UI` → `src/DataWizard.App`. Namespaces are unified under
  `DataWizard.Core.*`; `CSV` and `XLS` no longer sit in separate namespaces.
- **API reshaped.** `CSV` and `XLS` are replaced by `CsvAnalyzer`, `ExcelReader`,
  `ExcelWriter` and `ConversionService`.
- **Settings are JSON** at `%APPDATA%\DataWizards\settings.json`. Pattern lists
  from the old `DataWizard.config.xml` can be brought over with *Import legacy
  XML* on the Detection tab.

### Fixed

- **The solution did not build from a clean clone.** `test/Program.cs` still
  called `XLS.ToCsv` with a `quoteAllText:` argument after the parameter had been
  renamed to `quoteMode`, and the UI project required a ClickOnce signing
  certificate that was not in the repository.
- **A byte order mark leaked into the first column name.** Analysis rewound its
  `StreamReader` with `BaseStream.Seek(0)`, which re-read the BOM as content.
  Each stage now reads from a fresh reader.
- **Header detection failed silently on ragged files.** The second row was indexed
  using the first row's field count; the resulting `IndexOutOfRangeException` was
  swallowed by a `catch` that returned "no header".
- **Header and column names were split with `string.Split`,** so a quoted value
  containing the separator shifted every following column.
- **Separator detection counted raw character frequency,** including characters
  inside quoted values. A description column full of commas outvoted the real
  separator. Candidates are now scored on how consistently they produce the same
  field count.
- **Reported line counts were the analysis sample size,** capped at 100, not the
  size of the file.
- **Date parsing used the ambient culture.** `01/02/2025` resolved differently on
  a German and an American machine. Parsing now uses only the configured formats.
- **CSV export rounded every number to two decimals** (`0.##`), losing data on
  every round trip, and hard-coded German number formatting regardless of the
  chosen separator or encoding.
- **The Convert button stayed disabled after one conversion,** requiring a restart.
- **"UTF-8" and "UTF-8 with BOM" produced identical output,** both with a BOM.
- **Leading zeros were lost** on unquoted values unless a config override existed.
- **Long digit runs lost precision.** Values beyond a double's range — EAN codes,
  IBANs, long article numbers — were silently rounded.
- **`CharsetDetector` returning no result caused a `NullReferenceException`** on
  empty or ambiguous files.
- **Text cells were written as `t="str"`,** which the format reserves for cached
  formula results. They are now inline strings.
- **Shared strings were resolved by walking the table per cell,** making large
  workbook exports quadratic.
- **Whole worksheets were assembled in memory** before being written.
- **Missing worksheet rows shifted the CSV** instead of producing blank lines.
- Removed per-line and per-field `Console.WriteLine` calls from the analysis path.
- Removed `MainForm.cs`, a second complete 974-line form that was compiled but
  never used, and the dead `--watch` command-line stub.

### Added

- **Avalonia 12 interface** with light and dark themes, following the system
  setting by default.
- **Analysis panel** showing the encoding, separator candidates with their scores,
  the header signals with their reasoning, per-column types, and the first rows
  laid out in aligned columns.
- **Extended detection settings**, all adjustable in the interface against a live
  sample: encoding confidence and fallback, separator candidates and consistency
  threshold, four individually weighted header signals with a score threshold,
  number culture and thousands handling, leading-zero preservation, an integer
  digit limit, a date format list with a plausible year range, boolean detection,
  and per-column data type rules matched by name or position.
- **Folder watcher**, previously an empty tab: multiple rules, file masks,
  optional subfolders, file settling before conversion, retries, an error folder,
  and keep/move/delete for the source. Results are ignored on the way back in so a
  rule cannot feed on its own output.
- **RFC 4180 record reader** handling quoted values that span several lines.
- **Streaming XLSX writing** via `OpenXmlWriter`, plus bold and frozen header
  rows, optional auto-filter, and estimated column widths.
- **Log tab** with severity filters, a text filter, follow-tail and copy.
- **Help tab** rendered natively instead of in the legacy `WebBrowser` control,
  which ran in an Internet Explorer compatibility mode and broke the stylesheet.
- **Settings persistence**, replacing an empty `saveSettings()` stub.
- **133 tests** covering the engine and headless smoke tests that realise every
  tab, so a XAML error in an unopened tab cannot reach a release.

### Changed

- All user-facing text, log output, code comments and documentation are in
  English. The previous build mixed German log messages and dialog filters with
  English labels.
- Default header patterns are anchored on word boundaries. The old `.*id.*` also
  matched "Bildname" and "Identity", so almost any text row looked like a header.

## [1.0.0] - 2025-01-30

### Added - Core Library

#### CSV Handling
- **CSV Parser** with automatic separator detection (`;`, `,`, `\t`, `|`)
- **Encoding Detection** using UTF.Unknown library
- **Header Recognition** with configurable pattern matching
- **Quote Handling** (RFC 4180 compliant)
- **Field Splitting** with proper quote escape handling
- **German Number Format** support (`123,45` alongside `123.45`)
- **SplitLine()** method returns `CsvField` objects with quote information

#### Excel Handling
- **XLSX Creation** with DocumentFormat.OpenXml
- **Data Type Detection** (Integer, Decimal, Date, Text)
- **StyleIndex System** for proper Excel formatting:
  - StyleIndex 0: Text/General
  - StyleIndex 1: Date (dd.MM.yyyy)
  - StyleIndex 2: Decimal (#,##0.00)
  - StyleIndex 3: Integer (0)
- **Multi-Sheet Support** for reading and writing
- **IDisposable Pattern** for proper resource management

#### Excel to CSV Export
- **ToCsv()** static method for Excel → CSV conversion
- **Multi-Sheet Export** with `exportAllSheets` parameter
- **Automatic Filename Generation** (appends sheet name)
- **Encoding Support** (UTF-8, Windows-1252, ISO-8859-15, custom)
- **Separator Options** (`;`, `,`, `\t`, `|`)
- **Quote Control** (`quoteAllText` parameter)
- **NumberFormat Detection** from Excel stylesheet
- **Date Format Recognition** (prevents false date detection)

#### Configuration System
- **XML-based Configuration** (`DataWizard.config.xml`)
- **Header Field Patterns** with regex support
- **Data Type Overrides** per field name
- **Field Data Types**: Auto, Text, Integer, Decimal, Date
- **Pattern Matching**: String contains or regex
- **Default Configuration** with 40+ common field patterns

### Features

#### Smart Type Detection
- **Quote Priority**: Fields in quotes → always Text
- **Config Override**: XML configuration takes precedence
- **Number Detection**: Both English and German formats
- **Date Validation**: Requires separators, prevents false positives
- **Heuristics**: Numbers < 100 never treated as dates

#### Robust Encoding
- **Auto-Detection**: UTF.Unknown library
- **Multi-Encoding Support**: UTF-8, UTF-16, Windows-1252, ISO-8859-x
- **BOM Handling**: Automatic byte order mark detection

#### Error Handling
- **File Validation**: Checks file existence before processing
- **Graceful Fallbacks**: Auto-detection falls back to defaults
- **Exception Messages**: Clear error descriptions
- **Resource Cleanup**: Proper disposal of streams and documents

### Fixed

#### Number Format Issues
- **Fixed**: `113,2` (German format) was parsed as `1132` (thousands separator)
  - **Solution**: Changed from `NumberStyles.Any` to `NumberStyles.Float | AllowLeadingSign`
  - **Result**: Correctly parses `113,2` as `113.2`

#### False Date Detection
- **Fixed**: Number `6` was interpreted as date (6. January)
  - **Solution**: `IsValidDate()` method with strict validation
  - **Result**: Only real dates (with separators) are recognized

#### DateTime.TryParse Too Aggressive
- **Fixed**: `DateTime.TryParse()` accepted simple numbers as dates
  - **Solution**: Check numbers FIRST, then dates
  - **Result**: Priority order prevents false date detection

#### OADate Confusion
- **Fixed**: Excel numbers incorrectly converted to dates in CSV export
  - **Solution**: Read actual NumberFormat from Excel stylesheet
  - **Result**: Only cells with date format are exported as dates

#### StyleIndex Reliability
- **Fixed**: Foreign Excel files with unknown StyleIndex caused issues
  - **Solution**: `IsDateFormatted()` reads actual NumberFormat ID
  - **Result**: Robust against Excel files from any source

### Technical Improvements

#### Code Quality
- **Class Rename**: `Field` → `CsvField` (avoids naming conflicts)
- **Helper Methods**: `FormatNumber()`, `IsValidDate()`, `IsDateFormatted()`
- **Code Comments**: Comprehensive XML documentation
- **Error Handling**: Try-catch blocks with meaningful exceptions

#### Architecture
- **Separation of Concerns**: CSV vs. Excel logic separated
- **Dependency Injection**: Config passed to XLS via setter
- **Static Methods**: ToCsv() doesn't require instance
- **Return Values**: ToCsv() returns list of created files

### Documentation

- **README.md**: Comprehensive GitHub documentation
- **CONFIG_DOCUMENTATION.md**: Configuration system guide
- **ENCODING_GUIDE.md**: Encoding comparison (ISO-8859-1 vs Windows-1252)
- **EXCEL_TO_CSV.md**: Excel export documentation
- **EXPORT_ALL_SHEETS.md**: Multi-sheet export guide
- **STYLEINDEX_DOCUMENTATION.md**: Excel formatting guide
- **GERMAN_NUMBER_FORMAT.md**: German number format handling
- **NUMBERSTYLES_FIX.md**: NumberStyles.Any problem explanation
- **OADATE_FIX.md**: OADate vs. numbers issue
- **DATETIME_TRYPARSE_FIX.md**: DateTime.TryParse problem
- **NUMBER_VS_DATE_HEURISTIC.md**: Heuristic explanation
- **QUOTES_AS_TEXT.md**: Quote handling documentation

### Dependencies

- **.NET Framework 4.7.2**
- **DocumentFormat.OpenXml 3.3.0**
- **UTF.Unknown 2.5.1**

### Known Limitations

- **Thousands Separators**: Not supported in input (by design)
  - Prevents ambiguity between decimal and thousand separators
  - Example: `1.234,56` vs `1,234.56`
  
- **Multi-line Cells**: Not fully tested
  - CSV fields with newlines inside quotes
  
- **Very Large Files**: Memory-bound
  - Entire file loaded into memory
  - Consider streaming for files > 100MB

### Breaking Changes

None - this is the initial release.

---

## [Unreleased]

### Planned Features

#### Windows UI Application
- Drag & drop file conversion
- Visual configuration editor
- Batch processing interface
- Live preview before conversion
- Progress indicators
- Error reporting GUI

#### File Watcher Service
- Monitor input folder
- Automatic conversion on file arrival
- Configurable conversion rules per folder
- Scheduled processing
- Error logging and retry logic
- Windows Service deployment

#### Additional Formats
- **ODS Support**: OpenDocument Spreadsheet
- **XLS Support**: Legacy Excel format
- **TSV Optimization**: Tab-separated values
- **Fixed-Width**: Fixed-width text files

#### Performance Improvements
- **Streaming Parser**: For large CSV files
- **Async Operations**: Non-blocking I/O
- **Parallel Processing**: Multi-threaded batch conversion
- **Memory Optimization**: Reduced memory footprint

#### Advanced Features
- **Data Validation**: Pre-conversion validation rules
- **Transformation Rules**: Custom field transformations
- **Merge Operations**: Combine multiple CSVs
- **Split Operations**: Split large files
- **Column Mapping**: Flexible column reordering

---

## Version History

- **1.0.0** (2025-01-30) - Initial release
  - Core CSV ↔ Excel conversion
  - Configuration system
  - Multi-sheet support
  - Comprehensive documentation

---

### Contributors

Thanks to everyone who contributed to this release!

- Initial development and architecture
- Bug fixes and testing
- Documentation and examples

---

### Support

For issues, questions, or contributions:
- **GitHub Issues**: [Report bugs or request features](https://github.com/yourusername/DataWizard/issues)
- **Pull Requests**: [Contribute code](https://github.com/yourusername/DataWizard/pulls)
- **Documentation**: Check the docs/ folder for detailed guides
