using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataWizard;
using System.Globalization;


namespace libDataWizard
{
    public class XLS : IDisposable
    {
        public Boolean overwrite { get; set; }
        public string filePath { get; set; }
        private SpreadsheetDocument document;
        private WorksheetPart worksheetPart;
        private SheetData sheetData;
        private WorkbookPart workbookPart;
        private bool disposed = false;

        // Für Datentyp-Überschreibungen
        private string[] headerFieldNames;
        private DataWizardConfig config;

        public XLS(string filePath, bool overwrite)
        {
            this.filePath = filePath;
            this.overwrite = overwrite;

            if (File.Exists(filePath) && !overwrite)
            {
                throw new Exception("File exists.");
            }

            // Wenn Datei existiert und overwrite=true, lösche sie
            if (File.Exists(filePath) && overwrite)
            {
                File.Delete(filePath);
            }

            // Erstelle neues Dokument
            document = SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);

            // Workbook und Worksheet erstellen
            workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Stylesheet für Formatierungen erstellen
            CreateStylesheet();

            worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new Sheet()
            {
                Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Tabelle1"
            };
            sheets.Append(sheet);

            // SheetData-Referenz speichern für spätere Verwendung
            sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        }

        /// <summary>
        /// Erstellt ein Stylesheet mit Formaten für verschiedene Datentypen
        /// </summary>
        private void CreateStylesheet()
        {
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();

            // Fonts
            stylesPart.Stylesheet.Fonts = new Fonts();
            stylesPart.Stylesheet.Fonts.Count = 1;
            stylesPart.Stylesheet.Fonts.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Font());

            // Fills
            stylesPart.Stylesheet.Fills = new Fills();
            stylesPart.Stylesheet.Fills.Count = 2;
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } });

            // Borders
            stylesPart.Stylesheet.Borders = new Borders();
            stylesPart.Stylesheet.Borders.Count = 1;
            stylesPart.Stylesheet.Borders.AppendChild(new Border());

            // NumberingFormats
            stylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
            stylesPart.Stylesheet.NumberingFormats.Count = 2;

            // Custom format für Datum: dd.MM.yyyy (Format-ID 164+)
            stylesPart.Stylesheet.NumberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = 164,
                FormatCode = "dd.mm.yyyy"
            });

            // Custom format für Dezimalzahlen: #,##0.00
            stylesPart.Stylesheet.NumberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = 165,
                FormatCode = "#,##0.00"
            });

            // CellFormats
            stylesPart.Stylesheet.CellFormats = new CellFormats();
            stylesPart.Stylesheet.CellFormats.Count = 4;

            // Index 0: Standard (Text/General)
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            });

            // Index 1: Datum (dd.MM.yyyy)
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = 164,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                ApplyNumberFormat = true
            });

            // Index 2: Dezimal (#,##0.00)
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = 165,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                ApplyNumberFormat = true
            });

            // Index 3: Integer (0 - keine Dezimalstellen)
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = 1, // Built-in: 0
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                ApplyNumberFormat = true
            });

            stylesPart.Stylesheet.Save();
        }

        /// <summary>
        /// Setzt die Header-Feldnamen für Datentyp-Überschreibungen
        /// </summary>
        public void SetHeaderFieldNames(string[] fieldNames, DataWizardConfig configuration)
        {
            this.headerFieldNames = fieldNames;
            this.config = configuration;
        }

        /// <summary>
        /// Fügt eine neue Zeile mit Feldern hinzu
        /// </summary>
        public void AddRow(CsvField[] fields, uint rowIndex)
        {
            Row row = new Row() { RowIndex = rowIndex };

            for (int i = 0; i < fields.Length; i++)
            {
                Cell cell = CreateCell(fields[i], i);
                row.Append(cell);
            }

            sheetData.Append(row);
        }

        /// <summary>
        /// Erstellt eine Zelle basierend auf dem CsvField-Objekt und Spalten-Index
        /// </summary>
        private Cell CreateCell(CsvField field, int columnIndex)
        {
            Cell cell = new Cell();

            // Wenn Feld in Anführungszeichen war -> immer als Text behandeln
            if (field.Quotes)
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(field.Value);
                cell.StyleIndex = 0; // Standard/Text
                return cell;
            }

            // Prüfe ob es eine Datentyp-Überschreibung gibt
            FieldDataType? overrideType = null;
            if (config != null && headerFieldNames != null && columnIndex < headerFieldNames.Length)
            {
                string fieldName = headerFieldNames[columnIndex];
                overrideType = config.GetDataTypeOverride(fieldName);
            }

            // Verwende Override-Typ falls vorhanden, sonst automatische Erkennung
            if (overrideType.HasValue && overrideType.Value != FieldDataType.Auto)
            {
                return CreateCellWithDataType(field, overrideType.Value);
            }
            else
            {
                return CreateCellAutoDetect(field);
            }
        }

        /// <summary>
        /// Erstellt eine Zelle mit automatischer Datentyp-Erkennung
        /// </summary>
        private Cell CreateCellAutoDetect(CsvField field)
        {
            Cell cell = new Cell();

            // WICHTIG: ERST Zahl prüfen, DANN Datum!
            // Grund: DateTime.TryParse ist zu aggressiv und erkennt "6" als Datum

            // NumberStyles ohne AllowThousands - sonst wird Komma als Tausendertrennzeichen interpretiert
            NumberStyles numberStyle = NumberStyles.Float | NumberStyles.AllowLeadingSign;

            // Versuche, den Wert als Zahl zu parsen (InvariantCulture = Punkt als Dezimaltrennzeichen)
            bool isInvariantNumber = double.TryParse(field.Value, numberStyle, CultureInfo.InvariantCulture, out double numericValue);

            // Fallback: Versuche deutsches Format (Komma als Dezimaltrennzeichen)
            if (!isInvariantNumber)
            {
                isInvariantNumber = double.TryParse(field.Value, numberStyle, CultureInfo.GetCultureInfo("de-DE"), out numericValue);
            }

            if (isInvariantNumber)
            {
                // Prüfe ob es eine Ganzzahl ist
                if (numericValue == Math.Floor(numericValue) && Math.Abs(numericValue) < int.MaxValue)
                {
                    // Integer
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(numericValue.ToString(CultureInfo.InvariantCulture));
                    cell.StyleIndex = 3; // Integer-Format
                }
                else
                {
                    // Decimal
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(numericValue.ToString(CultureInfo.InvariantCulture));
                    cell.StyleIndex = 2; // Decimal-Format
                }
            }
            else if (IsValidDate(field.Value, out DateTime dateValue))
            {
                // Datumswert - nur wenn es ein ECHTES Datum ist (nicht nur eine Zahl)
                double oaDate = dateValue.ToOADate();
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(oaDate);
                cell.StyleIndex = 1; // Datum-Format (dd.MM.yyyy)
            }
            else
            {
                // Textwert
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(field.Value);
                cell.StyleIndex = 0; // Standard/Text
            }

            return cell;
        }

        /// <summary>
        /// Prüft ob ein String ein ECHTES Datum ist (nicht nur eine Zahl die zufällig als Datum geparst werden kann)
        /// </summary>
        private static bool IsValidDate(string value, out DateTime dateValue)
        {
            dateValue = DateTime.MinValue;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            // Wenn es nur Ziffern sind (z.B. "6", "123"), ist es KEIN Datum
            if (value.All(c => char.IsDigit(c)))
                return false;

            // Versuche Datum zu parsen
            if (!DateTime.TryParse(value, out dateValue))
                return false;

            // Zusätzliche Validierung: Muss Trennzeichen enthalten (-, /, .)
            if (!value.Contains('-') && !value.Contains('/') && !value.Contains('.'))
                return false;

            // Datum muss in vernünftigem Bereich sein
            if (dateValue.Year < 1900 || dateValue.Year > 2100)
                return false;

            return true;
        }

        /// <summary>
        /// Erstellt eine Zelle mit festgelegtem Datentyp
        /// </summary>
        private Cell CreateCellWithDataType(CsvField field, FieldDataType dataType)
        {
            Cell cell = new Cell();

            // NumberStyles ohne AllowThousands - wichtig für deutsches Format!
            NumberStyles numberStyle = NumberStyles.Float | NumberStyles.AllowLeadingSign;

            switch (dataType)
            {
                case FieldDataType.Text:
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(field.Value);
                    cell.StyleIndex = 0; // Standard/Text
                    break;

                case FieldDataType.Integer:
                    // Versuche englisches Format
                    bool isIntParsed = int.TryParse(field.Value, numberStyle, CultureInfo.InvariantCulture, out int intValue);

                    // Fallback: deutsches Format
                    if (!isIntParsed)
                    {
                        isIntParsed = int.TryParse(field.Value, numberStyle, CultureInfo.GetCultureInfo("de-DE"), out intValue);
                    }

                    if (isIntParsed)
                    {
                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(intValue);
                        cell.StyleIndex = 3; // Integer-Format (0)
                    }
                    else
                    {
                        // Fallback zu Text wenn Parsing fehlschlägt
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(field.Value);
                        cell.StyleIndex = 0;
                    }
                    break;

                case FieldDataType.Decimal:
                    // Versuche englisches Format
                    bool isDecimalParsed = double.TryParse(field.Value, numberStyle, CultureInfo.InvariantCulture, out double decimalValue);

                    // Fallback: deutsches Format
                    if (!isDecimalParsed)
                    {
                        isDecimalParsed = double.TryParse(field.Value, numberStyle, CultureInfo.GetCultureInfo("de-DE"), out decimalValue);
                    }

                    if (isDecimalParsed)
                    {
                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(decimalValue.ToString(CultureInfo.InvariantCulture));
                        cell.StyleIndex = 2; // Decimal-Format (#,##0.00)
                    }
                    else
                    {
                        // Fallback zu Text wenn Parsing fehlschlägt
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(field.Value);
                        cell.StyleIndex = 0;
                    }
                    break;

                case FieldDataType.Date:
                    if (DateTime.TryParse(field.Value, out DateTime dateValue))
                    {
                        double oaDate = dateValue.ToOADate();
                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(oaDate);
                        cell.StyleIndex = 1; // Datum-Format (dd.MM.yyyy)
                    }
                    else
                    {
                        // Fallback zu Text wenn Parsing fehlschlägt
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(field.Value);
                        cell.StyleIndex = 0;
                    }
                    break;

                default:
                    // Auto oder unbekannt - verwende Auto-Detection
                    return CreateCellAutoDetect(field);
            }

            return cell;
        }

        /// <summary>
        /// Alternative Methode zum Erstellen einer Text-Zelle
        /// </summary>
        private Cell CreateTextCell(string cellValue)
        {
            return new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(cellValue)
            };
        }

        /// <summary>
        /// Speichert das Excel-Dokument (sollte vor Dispose aufgerufen werden)
        /// </summary>
        public void Save()
        {
            if (workbookPart != null && workbookPart.Workbook != null)
            {
                workbookPart.Workbook.Save();
            }
        }

        /// <summary>
        /// Dispose-Pattern Implementation
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Managed resources freigeben
                    if (document != null)
                    {
                        document.Dispose();
                        document = null;
                    }
                }
                disposed = true;
            }
        }

        /// <summary>
        /// Konvertiert die Excel-Datei zu CSV
        /// </summary>
        /// <param name="xlsxPath">Pfad zur XLSX-Datei</param>
        /// <param name="csvPath">Pfad zur CSV-Ausgabedatei (bei exportAllSheets wird Sheetname eingefügt)</param>
        /// <param name="separator">Trennzeichen (Standard: ;)</param>
        /// <param name="encoding">Encoding (Standard: UTF-8)</param>
        /// <param name="quoteAllText">Alle Text-Felder in Anführungszeichen setzen</param>
        /// <param name="worksheetIndex">Index des Worksheets (0-basiert, Standard: 0, ignoriert wenn exportAllSheets=true)</param>
        /// <param name="exportAllSheets">Alle Worksheets exportieren (Standard: false)</param>
        /// <returns>Liste der erstellten CSV-Dateien</returns>
        public static List<string> ToCsv(string xlsxPath, string csvPath, char separator = ';', Encoding encoding = null, bool quoteAllText = true, int worksheetIndex = 0, bool exportAllSheets = false)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;

            if (!File.Exists(xlsxPath))
                throw new FileNotFoundException($"Excel-Datei nicht gefunden: {xlsxPath}");

            List<string> createdFiles = new List<string>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(xlsxPath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                var sheets = workbookPart.Workbook.Descendants<Sheet>().ToList();

                if (sheets.Count == 0)
                    throw new ArgumentException("Excel-Datei enthält keine Worksheets.");

                if (exportAllSheets)
                {
                    // Exportiere alle Worksheets
                    for (int i = 0; i < sheets.Count; i++)
                    {
                        Sheet sheet = sheets[i];
                        string sheetName = SanitizeFileName(sheet.Name);

                        // Erstelle Dateinamen mit Sheetnamen
                        string outputPath = InsertSheetNameIntoPath(csvPath, sheetName);

                        // Exportiere Sheet
                        ExportSheetToCsv(sheet, workbookPart, outputPath, separator, encoding, quoteAllText);
                        createdFiles.Add(outputPath);
                    }
                }
                else
                {
                    // Exportiere nur ein Worksheet
                    Sheet sheet = sheets.ElementAtOrDefault(worksheetIndex);
                    if (sheet == null)
                        throw new ArgumentException($"Worksheet mit Index {worksheetIndex} nicht gefunden.");

                    ExportSheetToCsv(sheet, workbookPart, csvPath, separator, encoding, quoteAllText);
                    createdFiles.Add(csvPath);
                }
            }

            return createdFiles;
        }

        /// <summary>
        /// Fügt den Sheetnamen in den Dateipfad ein
        /// </summary>
        private static string InsertSheetNameIntoPath(string csvPath, string sheetName)
        {
            string directory = Path.GetDirectoryName(csvPath);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(csvPath);
            string extension = Path.GetExtension(csvPath);

            // Dateiname_Sheetname.csv
            string newFileName = $"{fileNameWithoutExt}_{sheetName}{extension}";

            if (string.IsNullOrEmpty(directory))
                return newFileName;

            return Path.Combine(directory, newFileName);
        }

        /// <summary>
        /// Bereinigt Sheetnamen für Dateinamen (entfernt ungültige Zeichen)
        /// </summary>
        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "Sheet";

            // Entferne ungültige Zeichen für Dateinamen
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                name = name.Replace(c, '_');
            }

            // Entferne auch weitere problematische Zeichen
            name = name.Replace(" ", "_");  // Leerzeichen zu Unterstrich
            name = name.Replace(":", "_");
            name = name.Replace("/", "_");
            name = name.Replace("\\", "_");

            return name;
        }

        /// <summary>
        /// Exportiert ein einzelnes Sheet zu CSV
        /// </summary>
        private static void ExportSheetToCsv(Sheet sheet, WorkbookPart workbookPart, string csvPath, char separator, Encoding encoding, bool quoteAllText)
        {
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            // SharedStringTable für Text-Werte
            SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;

            using (StreamWriter writer = new StreamWriter(csvPath, false, encoding))
            {
                foreach (Row row in sheetData.Elements<Row>())
                {
                    List<string> cellValues = new List<string>();

                    // Hole alle Zellen der Zeile
                    var cells = row.Elements<Cell>().ToList();

                    if (cells.Count == 0)
                        continue; // Leere Zeile überspringen

                    // Bestimme die maximale Spaltenanzahl
                    // WICHTIG: CellReference kann null sein, also filtern!
                    int maxColumn = 0;

                    foreach (Cell cell in cells)
                    {
                        if (!string.IsNullOrEmpty(cell.CellReference))
                        {
                            int colIndex = GetColumnIndex(cell.CellReference);
                            if (colIndex > maxColumn)
                                maxColumn = colIndex;
                        }
                    }

                    // Fallback: Wenn alle CellReferences null sind, verwende einfach die Anzahl der Zellen
                    if (maxColumn == 0 && cells.Count > 0)
                    {
                        maxColumn = cells.Count - 1;
                    }

                    // Iteriere durch alle Spalten (inklusive leere)
                    for (int colIndex = 0; colIndex <= maxColumn; colIndex++)
                    {
                        string cellValue = "";

                        // Finde Zelle für diese Spalte
                        Cell cell = null;

                        // Wenn CellReference vorhanden, verwende es
                        cell = cells.FirstOrDefault(c => !string.IsNullOrEmpty(c.CellReference) && GetColumnIndex(c.CellReference) == colIndex);

                        // Fallback: Wenn keine CellReference, verwende Index
                        if (cell == null && colIndex < cells.Count)
                        {
                            cell = cells[colIndex];
                        }

                        if (cell != null && cell.CellValue != null)
                        {
                            cellValue = GetCellValue(cell, stringTablePart, workbookPart);
                        }

                        // Formatiere Zellwert für CSV
                        cellValues.Add(FormatCsvField(cellValue, separator, quoteAllText, cell));
                    }

                    // Schreibe Zeile
                    writer.WriteLine(string.Join(separator.ToString(), cellValues));
                }
            }
        }

        /// <summary>
        /// Konvertiert CellReference (z.B. "A1", "B2", "AA1") zu Spalten-Index (0-basiert)
        /// </summary>
        private static int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
                return 0;

            // Extrahiere Buchstaben (Spalte) aus CellReference
            string columnName = new string(cellReference.Where(c => char.IsLetter(c)).ToArray());

            if (string.IsNullOrEmpty(columnName))
                return 0;

            int columnIndex = 0;

            // Konvertiere Buchstaben zu Zahl (A=1, B=2, ..., Z=26, AA=27, AB=28, ...)
            for (int i = 0; i < columnName.Length; i++)
            {
                columnIndex = columnIndex * 26 + (columnName[i] - 'A' + 1);
            }

            return columnIndex - 1; // 0-basiert: A=0, B=1, etc.
        }

        /// <summary>
        /// Extrahiert den Zellwert aus einer Excel-Zelle
        /// </summary>
        private static string GetCellValue(Cell cell, SharedStringTablePart stringTablePart, WorkbookPart workbookPart)
        {
            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                // Text aus SharedStringTable
                if (stringTablePart != null)
                {
                    return stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
                }
            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
            {
                // Boolean
                return value == "1" ? "TRUE" : "FALSE";
            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.String)
            {
                // Inline String
                return value;
            }

            // Prüfe ob die Zelle ein Datum-Format hat (durch Stylesheet)
            bool isDateFormat = IsDateFormatted(cell, workbookPart);

            if (isDateFormat)
            {
                // Behandle als Datum
                if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double oaDate))
                {
                    try
                    {
                        DateTime date = DateTime.FromOADate(oaDate);
                        return date.ToString("yyyy-MM-dd"); // ISO-Format
                    }
                    catch
                    {
                        // Ungültige OADate - gib Zahl zurück
                        return FormatNumber(value);
                    }
                }
            }

            // Nicht als Datum formatiert - behandle als Zahl
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
            {
                return FormatNumber(value);
            }

            // Fallback: Rohwert
            return value;
        }

        /// <summary>
        /// Prüft ob eine Zelle mit Datum-Format formatiert ist
        /// </summary>
        private static bool IsDateFormatted(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.StyleIndex == null)
                return false;

            try
            {
                WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart;
                if (stylesPart == null || stylesPart.Stylesheet == null)
                    return false;

                Stylesheet stylesheet = stylesPart.Stylesheet;
                CellFormats cellFormats = stylesheet.CellFormats;

                if (cellFormats == null)
                    return false;

                uint styleIndex = cell.StyleIndex.Value;
                if (styleIndex >= cellFormats.Count)
                    return false;

                CellFormat cellFormat = (CellFormat)cellFormats.ElementAt((int)styleIndex);
                if (cellFormat.NumberFormatId == null)
                    return false;

                uint numFmtId = cellFormat.NumberFormatId.Value;

                // Excel built-in Datumsformate: 14-22, 45-47
                // 14: m/d/yyyy
                // 15: d-mmm-yy
                // 16: d-mmm
                // 17: mmm-yy
                // 18: h:mm AM/PM
                // 19: h:mm:ss AM/PM
                // 20: h:mm
                // 21: h:mm:ss
                // 22: m/d/yy h:mm
                // 45-47: weitere Datumsformate

                // Unser eigenes Schema: 164 (Datum)
                bool isBuiltInDateFormat = (numFmtId >= 14 && numFmtId <= 22) || (numFmtId >= 45 && numFmtId <= 47);
                bool isCustomDateFormat = (numFmtId == 164); // Unser eigenes Format

                return isBuiltInDateFormat || isCustomDateFormat;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Formatiert eine Zahl im deutschen Format (Komma als Dezimaltrennzeichen)
        /// </summary>
        private static string FormatNumber(string value)
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double numValue))
            {
                // Prüfe ob es eine Ganzzahl ist
                if (numValue == Math.Floor(numValue) && Math.Abs(numValue) < int.MaxValue)
                {
                    // Integer - keine Dezimalstellen
                    return numValue.ToString("0", CultureInfo.GetCultureInfo("de-DE"));
                }
                else
                {
                    // Decimal - mit Dezimalstellen (aber keine unnötigen Nullen)
                    return numValue.ToString("0.##", CultureInfo.GetCultureInfo("de-DE"));
                }
            }

            return value;
        }

        /// <summary>
        /// Formatiert ein Feld für CSV (mit Quotes wenn nötig)
        /// </summary>
        private static string FormatCsvField(string value, char separator, bool quoteAllText, Cell cell)
        {
            bool needsQuotes = false;

            // Prüfe ob Quotes nötig sind
            if (value.Contains(separator) || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
            {
                needsQuotes = true;
            }

            // Wenn quoteAllText und es ist ein Text-Feld
            if (quoteAllText && cell != null)
            {
                // Text-Felder (DataType String oder StyleIndex 0)
                if ((cell.DataType != null && cell.DataType.Value == CellValues.String) ||
                    (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) ||
                    (cell.StyleIndex != null && cell.StyleIndex.Value == 0))
                {
                    needsQuotes = true;
                }
            }

            if (needsQuotes)
            {
                // Escape Anführungszeichen (verdoppeln)
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }

            return value;
        }

        ///// <summary>
        ///// Konvertiert CellReference (z.B. "A1", "B2") zu Spalten-Index (0-basiert)
        ///// </summary>
        //private static int GetColumnIndex(string cellReference)
        //{
        //    if (string.IsNullOrEmpty(cellReference))
        //        return 0;

        //    // Extrahiere Buchstaben (Spalte) aus CellReference
        //    string columnName = new string(cellReference.Where(c => char.IsLetter(c)).ToArray());

        //    int columnIndex = 0;
        //    int factor = 1;

        //    // Konvertiere Buchstaben zu Zahl (A=0, B=1, ..., Z=25, AA=26, ...)
        //    for (int i = columnName.Length - 1; i >= 0; i--)
        //    {
        //        columnIndex += (columnName[i] - 'A' + 1) * factor;
        //        factor *= 26;
        //    }

        //    return columnIndex - 1; // 0-basiert
        //}
        /// <summary>
        /// Konvertiert CellReference (z.B. "A1", "B2", "AA1") zu Spalten-Index (0-basiert)
        /// </summary>
        //private static int GetColumnIndex(string cellReference)
        //{
        //    if (string.IsNullOrEmpty(cellReference))
        //        return 0;

        //    // Extrahiere Buchstaben (Spalte) aus CellReference
        //    string columnName = new string(cellReference.Where(c => char.IsLetter(c)).ToArray());

        //    if (string.IsNullOrEmpty(columnName))
        //        return 0;

        //    int columnIndex = 0;

        //    // Konvertiere Buchstaben zu Zahl (A=1, B=2, ..., Z=26, AA=27, AB=28, ...)
        //    for (int i = 0; i < columnName.Length; i++)
        //    {
        //        columnIndex = columnIndex * 26 + (columnName[i] - 'A' + 1);
        //    }

        //    return columnIndex - 1; // 0-basiert: A=0, B=1, etc.
        //}
        //~XLS()
        //{
        //    Dispose(false);
        //}
    }
}
