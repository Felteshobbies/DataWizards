# Changelog

All notable changes to DataWizard will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
