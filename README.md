# DataWizard

**Professional CSV ↔ Excel conversion library for .NET**

DataWizard is a powerful C# library for intelligent CSV and Excel file handling with automatic data type detection, configurable field recognition, and robust encoding support.

[![.NET](https://img.shields.io/badge/.NET-4.7.2-blue.svg)](https://dotnet.microsoft.com/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

## ✨ Features

### Core Functionality
- 📊 **Bi-directional Conversion**: CSV ↔ Excel (XLSX) with full data type preservation
- 🔍 **Smart Type Detection**: Automatic recognition of integers, decimals, dates, and text
- 🎯 **Header Recognition**: Configurable pattern matching with regex support
- 🌍 **Multi-Encoding**: UTF-8, Windows-1252, ISO-8859-15, and custom encodings
- 🔧 **Format Support**: Multiple separators (`;`, `,`, `\t`, `|`)

### Advanced Features
- ⚙️ **XML Configuration**: Centralized field type overrides and header patterns
- 🔢 **Number Format Support**: Both English (`123.45`) and German (`123,45`) formats
- 📝 **Quote Handling**: Automatic preservation of leading zeros and special characters
- 📅 **Date Intelligence**: Prevents false date detection for numeric values
- 📑 **Multi-Sheet Export**: Export all Excel worksheets to separate CSV files
- 🎨 **Excel Styling**: Proper number formats, date formats, and text alignment

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/DataWizard.git

# Build the solution
cd DataWizard
dotnet build
```

### Basic Usage

#### CSV to Excel

```csharp
using DataWizard;
using libDataWizard;

// Load CSV and convert to Excel
CSV csv = new CSV();
csv.Load(@"C:\data\input.csv");
csv.WriteXLSX(@"C:\data\output.xlsx", overwrite: true);
```

#### Excel to CSV

```csharp
// Convert Excel to CSV
XLS.ToCsv(@"C:\data\input.xlsx", @"C:\data\output.csv");

// Export all worksheets
List<string> files = XLS.ToCsv(@"C:\data\workbook.xlsx", 
    @"C:\data\export.csv", 
    exportAllSheets: true);
```

## 📖 Documentation

### CSV Analysis

DataWizard automatically detects:
- **Separator** (`;`, `,`, `\t`, `|`) with confidence probability
- **Encoding** (UTF-8, Windows-1252, etc.)
- **Header row** based on configurable patterns
- **Start line** (skips empty lines)
- **Field count** consistency

```csharp
CSV csv = new CSV();
csv.Load(@"data.csv");

Console.WriteLine($"Separator: {csv.Separator}");
Console.WriteLine($"Confidence: {csv.SeparatorProbability}%");
Console.WriteLine($"Encoding: {csv.DetectedEncoding}");
Console.WriteLine($"Header detected: {csv.DetectedHeaderLine}");
```

### Configuration System

Create a `DataWizard.config.xml` to customize behavior:

```xml
<?xml version="1.0" encoding="utf-8"?>
<DataWizardConfig>
  <HeaderFieldNames>
    <Pattern Name="^id$" IsRegex="true" />
    <Pattern Name="name" IsRegex="false" />
    <Pattern Name="preis" IsRegex="false" />
    <Pattern Name="datum" IsRegex="false" />
  </HeaderFieldNames>
  
  <DataTypeOverrides>
    <Override FieldNamePattern="^.*id$" IsRegex="true" DataType="Text" />
    <Override FieldNamePattern="plz" IsRegex="false" DataType="Text" />
    <Override FieldNamePattern="preis" IsRegex="false" DataType="Decimal" />
    <Override FieldNamePattern="menge" IsRegex="false" DataType="Integer" />
    <Override FieldNamePattern="datum" IsRegex="false" DataType="Date" />
  </DataTypeOverrides>
</DataWizardConfig>
```

#### Data Types

- `Auto`: Automatic detection (default)
- `Text`: String values (preserves leading zeros: `001`, `01234`)
- `Integer`: Whole numbers without decimals
- `Decimal`: Numbers with decimal places
- `Date`: Date values (ISO format: `yyyy-MM-dd`)

### Advanced Options

#### Custom Encoding

```csharp
// Excel to CSV with Windows-1252 encoding
XLS.ToCsv(@"input.xlsx", @"output.csv", 
    separator: ';',
    encoding: Encoding.GetEncoding(1252));

// CSV to Excel with custom separator detection
CSV csv = new CSV();
csv.Separators = new[] { ',', ';', '\t', '|' };
csv.Load(@"data.csv");
```

#### Quote Control

```csharp
// All text fields with quotes
XLS.ToCsv(@"input.xlsx", @"output.csv", quoteAllText: true);
// Result: "Name";"City";"PLZ"

// Minimal quotes (only when necessary)
XLS.ToCsv(@"input.xlsx", @"output.csv", quoteAllText: false);
// Result: Name;City;"01234"
```

#### Multi-Sheet Export

```csharp
// Export all worksheets to separate files
List<string> files = XLS.ToCsv(
    @"C:\data\workbook.xlsx", 
    @"C:\output\data.csv",
    exportAllSheets: true
);

// Results:
// - C:\output\data_Sheet1.csv
// - C:\output\data_Kunden.csv
// - C:\output\data_Produkte.csv

foreach (string file in files)
{
    Console.WriteLine($"Created: {file}");
}
```

## 🎯 Use Cases

### 1. Preserve Leading Zeros

```csv
ID;PLZ;Telefon
"001";"01234";"+49 176 12345678"
```

Using quotes or field type override ensures values keep their format.

### 2. German Number Format

```csv
Artikel;Preis;Menge
Widget;19,99;5
Gadget;123,45;10
```

Automatic detection of German format (`19,99`) and English format (`19.99`).

### 3. Date Handling

```csv
Datum;Beschreibung
2025-01-31;Bestellung
31.01.2025;Lieferung
```

Intelligent date detection prevents false positives (e.g., `45.66` is NOT a date).

### 4. Batch Processing

```csharp
// Convert all CSVs in a folder
foreach (string csvFile in Directory.GetFiles(@"C:\input", "*.csv"))
{
    CSV csv = new CSV();
    csv.Load(csvFile);
    
    string xlsxFile = Path.ChangeExtension(csvFile, ".xlsx");
    csv.WriteXLSX(xlsxFile, overwrite: true);
}
```

## 🏗️ Architecture

```
DataWizard/
├── DataWizard/              # Main library (CSV handling)
│   └── CSV.cs              # CSV parser and analyzer
├── libDataWizard/          # Helper library (Excel handling)
│   ├── XLS.cs              # Excel read/write operations
│   └── DataWizardConfig.cs # Configuration system
└── test/                   # Test project
    └── Program.cs          # Usage examples
```

## 🔧 Technical Details

### Dependencies

- **.NET Framework 4.7.2**
- **DocumentFormat.OpenXml 3.3.0** - Excel file handling
- **UTF.Unknown 2.5.1** - Encoding detection

### Supported Formats

#### CSV
- **Separators**: `;` `,` `\t` `|`
- **Encodings**: UTF-8, UTF-16, Windows-1252, ISO-8859-1, ISO-8859-15
- **Line Endings**: CRLF, LF
- **Quotes**: RFC 4180 compliant

#### Excel
- **Format**: XLSX (Office Open XML)
- **Versions**: Excel 2007 and later
- **Features**: Multiple worksheets, cell formatting, data types

### Data Type Detection

**Priority Order:**
1. **Quoted fields** (`"001"`) → Always treated as text
2. **Config overrides** → Forced type from XML configuration
3. **Number detection** → English (`123.45`) and German (`123,45`) formats
4. **Date detection** → Only with valid separators (`-`, `/`, `.`)
5. **Fallback** → Text

### Number Format Heuristics

To prevent false date detection:
- Numbers < 100 (e.g., `6`, `45.66`) → **Never** treated as dates
- Numbers ≥ 100 → Treated as numbers unless explicitly date-formatted
- Date format required: Must have separators and valid date structure

## 🛣️ Roadmap

### Upcoming Features

- ✅ **Core Library** (Current)
- 🔜 **Windows UI Application**
  - Drag & drop conversion
  - Visual configuration editor
  - Batch processing
  - Preview before conversion
  
- 🔜 **File Watcher Service**
  - Monitor input folder
  - Automatic conversion
  - Configurable rules per folder
  - Error logging and retry

## 📚 Examples

### Example 1: Import with Custom Config

```csharp
// Load configuration
CSV csv = new CSV(@"C:\config\DataWizard.config.xml");

// Analyze and convert
csv.Load(@"C:\data\customers.csv");
csv.WriteXLSX(@"C:\data\customers.xlsx", overwrite: true);
```

### Example 2: Export with Options

```csharp
// Export with specific settings
XLS.ToCsv(
    xlsxPath: @"C:\data\report.xlsx",
    csvPath: @"C:\data\report.csv",
    separator: ',',
    encoding: Encoding.UTF8,
    quoteAllText: false,
    worksheetIndex: 0,
    exportAllSheets: false
);
```

### Example 3: Roundtrip Conversion

```csharp
// CSV → Excel → CSV (data integrity test)
CSV csv1 = new CSV();
csv1.Load(@"original.csv");
csv1.WriteXLSX(@"temp.xlsx", true);

XLS.ToCsv(@"temp.xlsx", @"roundtrip.csv");

// Compare original.csv with roundtrip.csv
```

## 🐛 Troubleshooting

### Problem: Leading zeros removed

**Solution:** Use quotes in CSV or configure field as `Text` type
```xml
<Override FieldNamePattern="id" DataType="Text" />
```

### Problem: German numbers not recognized

**Solution:** Library automatically detects both formats. Ensure no thousands separators in input.

### Problem: False date detection

**Solution:** 
- Ensure StyleIndex is set when writing Excel
- Numbers < 100 are never treated as dates
- Use explicit date format (`yyyy-MM-dd`)

### Problem: Encoding issues (Umlaute)

**Solution:** Specify correct encoding
```csharp
XLS.ToCsv(@"input.xlsx", @"output.csv", 
    encoding: Encoding.GetEncoding(1252)); // Windows-1252
```

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📧 Contact

Project Link: [https://github.com/yourusername/DataWizard](https://github.com/yourusername/DataWizard)

## 🙏 Acknowledgments

- [DocumentFormat.OpenXml](https://github.com/OfficeDev/Open-XML-SDK) - Excel file manipulation
- [UTF.Unknown](https://github.com/CharsetDetector/UTF-unknown) - Encoding detection

---

**Note:** This is a library component. GUI application and file watcher service are planned for future releases.
