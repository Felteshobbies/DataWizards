# Excel StyleIndex und Field → CsvField Änderungen

## Übersicht der Änderungen

### 1. Field-Klasse umbenannt zu CsvField
**Grund:** Vermeidung von Namenskonflikten mit anderen Libraries

**Änderungen:**
- `Field` → `CsvField` in allen Dateien
- Betrifft: CSV.cs, XLS.cs, Program.cs

### 2. Excel StyleSheet mit korrekten Datentyp-Formaten

## Excel StyleIndex Mapping

Die XLS-Klasse erstellt jetzt ein Stylesheet mit 4 verschiedenen Formaten:

| StyleIndex | Datentyp | Excel-Format | Beispiel Anzeige |
|------------|----------|--------------|------------------|
| 0 | Text/Standard | General | "Hallo", "ABC123" |
| 1 | Datum | dd.MM.yyyy | 31.01.2025 |
| 2 | Decimal | #,##0.00 | 1.234,56 |
| 3 | Integer | 0 | 42 |

## Wie StyleIndex funktioniert

### In Excel werden Datentypen über StyleIndex gesteuert:

```csharp
// FALSCH (ohne StyleIndex):
cell.DataType = CellValues.Number;
cell.CellValue = new CellValue(dateValue.ToOADate());
// Excel zeigt: 45678.0 (Zahl statt Datum!)

// RICHTIG (mit StyleIndex für Datum):
cell.DataType = CellValues.Number;
cell.CellValue = new CellValue(dateValue.ToOADate());
cell.StyleIndex = 1; // Datum-Format
// Excel zeigt: 31.01.2025 ✅
```

## Stylesheet-Struktur

### NumberingFormats (Custom Formats)

```csharp
// Format-ID 164: Datum
NumberingFormat
{
    NumberFormatId = 164,
    FormatCode = "dd.mm.yyyy"  // Deutsches Datumsformat
}

// Format-ID 165: Dezimal mit Tausendertrennzeichen
NumberingFormat
{
    NumberFormatId = 165,
    FormatCode = "#,##0.00"  // z.B. 1.234,56
}
```

### CellFormats (Styles)

```csharp
// Index 0: Standard/Text
new CellFormat
{
    NumberFormatId = 0,  // General
    ApplyNumberFormat = false
}

// Index 1: Datum (dd.MM.yyyy)
new CellFormat
{
    NumberFormatId = 164,  // Unser Custom-Format
    ApplyNumberFormat = true
}

// Index 2: Decimal (#,##0.00)
new CellFormat
{
    NumberFormatId = 165,  // Unser Custom-Format
    ApplyNumberFormat = true
}

// Index 3: Integer (0)
new CellFormat
{
    NumberFormatId = 1,  // Built-in: "0"
    ApplyNumberFormat = true
}
```

## Automatische Datentyp-Erkennung mit StyleIndex

### CreateCellAutoDetect() - Verbesserte Version

```csharp
private Cell CreateCellAutoDetect(CsvField field)
{
    if (double.TryParse(field.Value, ...))
    {
        // Prüfe ob es eine Ganzzahl ist
        if (numericValue == Math.Floor(numericValue))
        {
            // Integer
            cell.StyleIndex = 3;  // Format: "0"
        }
        else
        {
            // Decimal
            cell.StyleIndex = 2;  // Format: "#,##0.00"
        }
    }
    else if (DateTime.TryParse(field.Value, ...))
    {
        // Datum
        cell.StyleIndex = 1;  // Format: "dd.MM.yyyy"
    }
    else
    {
        // Text
        cell.StyleIndex = 0;  // Format: "General"
    }
}
```

### Vorher vs. Nachher

**VORHER:**
```csv
Datum,Preis,Menge
2025-01-31,19.99,42
```
Excel zeigt:
- Datum: `45678` (Zahl!)
- Preis: `19.99` (ohne Tausendertrennzeichen)
- Menge: `42.0` (Dezimalstellen obwohl Integer)

**NACHHER:**
```csv
Datum,Preis,Menge
2025-01-31,19.99,42
```
Excel zeigt:
- Datum: `31.01.2025` ✅
- Preis: `19,99` ✅ (mit Komma statt Punkt)
- Menge: `42` ✅ (ohne Dezimalstellen)

## Config-basierte Datentyp-Überschreibung mit StyleIndex

### Text-Override (z.B. für PLZ)

```csharp
case FieldDataType.Text:
    cell.DataType = CellValues.String;
    cell.CellValue = new CellValue(field.Value);
    cell.StyleIndex = 0;  // Standard/Text
    break;
```

**Beispiel:**
```csv
PLZ
01234
```
Excel zeigt: `01234` (nicht `1234` !)

### Integer-Override (z.B. für Mengen)

```csharp
case FieldDataType.Integer:
    if (int.TryParse(field.Value, ...))
    {
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(intValue);
        cell.StyleIndex = 3;  // Integer-Format: "0"
    }
```

**Beispiel:**
```csv
Menge
42
```
Excel zeigt: `42` (nicht `42.0`)

### Decimal-Override (z.B. für Preise)

```csharp
case FieldDataType.Decimal:
    if (double.TryParse(field.Value, ...))
    {
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(decimalValue);
        cell.StyleIndex = 2;  // Decimal-Format: "#,##0.00"
    }
```

**Beispiel:**
```csv
Preis
1234.56
```
Excel zeigt: `1.234,56` (mit Tausendertrennzeichen und Komma)

### Date-Override

```csharp
case FieldDataType.Date:
    if (DateTime.TryParse(field.Value, ...))
    {
        double oaDate = dateValue.ToOADate();
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(oaDate);
        cell.StyleIndex = 1;  // Datum-Format: "dd.MM.yyyy"
    }
```

**Beispiel:**
```csv
Geburtsdatum
1990-05-15
```
Excel zeigt: `15.05.1990` (deutsches Format)

## Built-in Excel NumberFormatIds

Für Referenz - häufig verwendete Built-in Formate:

| ID | Format | Beispiel |
|----|--------|----------|
| 0 | General | Automatisch |
| 1 | 0 | 42 |
| 2 | 0.00 | 42.00 |
| 3 | #,##0 | 1,234 |
| 4 | #,##0.00 | 1,234.56 |
| 14 | mm-dd-yy | 01-31-25 |
| 22 | m/d/yy h:mm | 1/31/25 14:30 |

Wir verwenden Custom-Formate ab ID 164 für mehr Kontrolle.

## Migration

**Alte Klasse `Field`:**
```csharp
Field[] fields = csv.SplitLine(line, separator);
```

**Neue Klasse `CsvField`:**
```csharp
CsvField[] fields = csv.SplitLine(line, separator);
```

Einfach alle Vorkommen von `Field` durch `CsvField` ersetzen (außer in XLS.cs wo es um Excel-Felder geht).

## Zusammenfassung

✅ **CsvField statt Field** - Keine Namenskonflikte mehr  
✅ **Stylesheet erstellt** - Korrekte Excel-Formatierung  
✅ **StyleIndex 0-3** - Text, Datum, Decimal, Integer  
✅ **Automatische Erkennung** - Integer vs. Decimal unterschieden  
✅ **Config-Overrides** - Verwenden korrekte StyleIndex  
✅ **Deutsches Format** - dd.MM.yyyy, #,##0.00  

## Test

```csharp
CSV csv = new CSV(@"config.xml");
csv.Load(@"test.csv");
csv.WriteXLSX(@"output.xlsx", true);
```

Öffne `output.xlsx` in Excel:
- Datum-Spalten: Angezeigt als Datum (nicht als Zahl)
- Preis-Spalten: Mit 2 Dezimalstellen
- Mengen-Spalten: Ohne Dezimalstellen
- PLZ-Spalten: Als Text (führende Nullen erhalten)
