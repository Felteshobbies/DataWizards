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
    /// <summary>
    /// Datentyp einer Excel-Zelle für CSV-Export
    /// </summary>
    internal enum CellType
    {
        Empty,      // Leere Zelle
        Text,       // Text/String
        Number,     // Zahl (Integer oder Decimal)
        Date,       // Datum
        Boolean     // Boolean
    }

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

            if (File.Exists(filePath) && overwrite)
            {
                File.Delete(filePath);
            }

            document = SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);

            workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

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

            sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        }

        private void CreateStylesheet()
        {
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();

            stylesPart.Stylesheet.Fonts = new Fonts();
            stylesPart.Stylesheet.Fonts.Count = 1;
            stylesPart.Stylesheet.Fonts.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Font());

            stylesPart.Stylesheet.Fills = new Fills();
            stylesPart.Stylesheet.Fills.Count = 2;
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } });

            stylesPart.Stylesheet.Borders = new Borders();
            stylesPart.Stylesheet.Borders.Count = 1;
            stylesPart.Stylesheet.Borders.AppendChild(new Border());

            stylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
            stylesPart.Stylesheet.NumberingFormats.Count = 2;
            stylesPart.Stylesheet.NumberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = 164,
                FormatCode = "dd.mm.yyyy"
            });
            stylesPart.Stylesheet.NumberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = 165,
                FormatCode = "#,##0.00"
            });

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
            // Index 3: Integer
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = 1,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                ApplyNumberFormat = true
            });

            stylesPart.Stylesheet.Save();
        }

        public void SetHeaderFieldNames(string[] fieldNames, DataWizardConfig configuration)
        {
            this.headerFieldNames = fieldNames;
            this.config = configuration;
        }

        /// <summary>
        /// Konvertiert 0-basierten Spalten-Index zu Excel-Spaltenname (A, B, ..., Z, AA, AB, ...)
        /// </summary>
        private static string GetColumnName(int columnIndex)
        {
            string columnName = "";
            int index = columnIndex;

            while (index >= 0)
            {
                columnName = (char)('A' + index % 26) + columnName;
                index = index / 26 - 1;
            }

            return columnName;
        }

        /// <summary>
        /// Fügt eine neue Zeile mit Feldern hinzu
        /// </summary>
        public void AddRow(CsvField[] fields, uint rowIndex)
        {
            Row row = new Row() { RowIndex = rowIndex };

            for (int i = 0; i < fields.Length; i++)
            {
                // ✅ CellReference korrekt setzen: "A1", "B1", "C2" etc.
                string cellReference = GetColumnName(i) + rowIndex.ToString();
                Cell cell = CreateCell(fields[i], i, cellReference);
                row.Append(cell);
            }

            sheetData.Append(row);
        }

        /// <summary>
        /// Erstellt eine Zelle mit CellReference
        /// </summary>
        private Cell CreateCell(CsvField field, int columnIndex, string cellReference)
        {
            Cell cell = new Cell();
            cell.CellReference = cellReference; // ✅ Immer setzen!

            // Wenn Feld in Anführungszeichen war -> immer als Text behandeln
            if (field.Quotes)
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(field.Value);
                cell.StyleIndex = 0;
                return cell;
            }

            // Prüfe ob es eine Datentyp-Überschreibung gibt
            FieldDataType? overrideType = null;
            if (config != null && headerFieldNames != null && columnIndex < headerFieldNames.Length)
            {
                string fieldName = headerFieldNames[columnIndex];
                overrideType = config.GetDataTypeOverride(fieldName);
            }

            if (overrideType.HasValue && overrideType.Value != FieldDataType.Auto)
            {
                return CreateCellWithDataType(field, overrideType.Value, cellReference);
            }
            else
            {
                return CreateCellAutoDetect(field, cellReference);
            }
        }

        /// <summary>
        /// Erstellt eine Zelle mit automatischer Datentyp-Erkennung
        /// </summary>
        private Cell CreateCellAutoDetect(CsvField field, string cellReference)
        {
            Cell cell = new Cell();
            cell.CellReference = cellReference; // ✅ Immer setzen!

            NumberStyles numberStyle = NumberStyles.Float | NumberStyles.AllowLeadingSign;

            bool isNumber = double.TryParse(field.Value, numberStyle, CultureInfo.InvariantCulture, out double numericValue);

            if (!isNumber)
                isNumber = double.TryParse(field.Value, numberStyle, CultureInfo.GetCultureInfo("de-DE"), out numericValue);

            if (isNumber)
            {
                if (numericValue == Math.Floor(numericValue) && Math.Abs(numericValue) < int.MaxValue)
                {
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(numericValue.ToString(CultureInfo.InvariantCulture));
                    cell.StyleIndex = 3; // Integer
                }
                else
                {
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(numericValue.ToString(CultureInfo.InvariantCulture));
                    cell.StyleIndex = 2; // Decimal
                }
            }
            else if (IsValidDate(field.Value, out DateTime dateValue))
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(dateValue.ToOADate());
                cell.StyleIndex = 1; // Datum
            }
            else
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(field.Value);
                cell.StyleIndex = 0; // Text
            }

            return cell;
        }

        private static bool IsValidDate(string value, out DateTime dateValue)
        {
            dateValue = DateTime.MinValue;

            if (string.IsNullOrWhiteSpace(value)) return false;
            if (value.All(c => char.IsDigit(c))) return false;
            if (!DateTime.TryParse(value, out dateValue)) return false;
            if (!value.Contains('-') && !value.Contains('/') && !value.Contains('.')) return false;
            if (dateValue.Year < 1900 || dateValue.Year > 2100) return false;

            return true;
        }

        private Cell CreateCellWithDataType(CsvField field, FieldDataType dataType, string cellReference)
        {
            Cell cell = new Cell();
            cell.CellReference = cellReference; // ✅ Immer setzen!

            NumberStyles numberStyle = NumberStyles.Float | NumberStyles.AllowLeadingSign;

            switch (dataType)
            {
                case FieldDataType.Text:
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(field.Value);
                    cell.StyleIndex = 0;
                    break;

                case FieldDataType.Integer:
                    bool isIntParsed = int.TryParse(field.Value, numberStyle, CultureInfo.InvariantCulture, out int intValue);
                    if (!isIntParsed)
                        isIntParsed = int.TryParse(field.Value, numberStyle, CultureInfo.GetCultureInfo("de-DE"), out intValue);

                    cell.DataType = isIntParsed ? CellValues.Number : CellValues.String;
                    cell.CellValue = isIntParsed ? new CellValue(intValue) : new CellValue(field.Value);
                    cell.StyleIndex = isIntParsed ? 3u : 0u;
                    break;

                case FieldDataType.Decimal:
                    bool isDecParsed = double.TryParse(field.Value, numberStyle, CultureInfo.InvariantCulture, out double decValue);
                    if (!isDecParsed)
                        isDecParsed = double.TryParse(field.Value, numberStyle, CultureInfo.GetCultureInfo("de-DE"), out decValue);

                    cell.DataType = isDecParsed ? CellValues.Number : CellValues.String;
                    cell.CellValue = isDecParsed ? new CellValue(decValue.ToString(CultureInfo.InvariantCulture)) : new CellValue(field.Value);
                    cell.StyleIndex = isDecParsed ? 2u : 0u;
                    break;

                case FieldDataType.Date:
                    if (DateTime.TryParse(field.Value, out DateTime dateValue))
                    {
                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(dateValue.ToOADate());
                        cell.StyleIndex = 1;
                    }
                    else
                    {
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(field.Value);
                        cell.StyleIndex = 0;
                    }
                    break;

                default:
                    return CreateCellAutoDetect(field, cellReference);
            }

            return cell;
        }

        public void Save()
        {
            if (workbookPart != null && workbookPart.Workbook != null)
            {
                workbookPart.Workbook.Save();
            }
        }

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
                    if (document != null)
                    {
                        document.Dispose();
                        document = null;
                    }
                }
                disposed = true;
            }
        }

        // ═══════════════════════════════════════════════════════════
        // XLSX → CSV
        // ═══════════════════════════════════════════════════════════

        public static List<string> ToCsv(string xlsxPath, string csvPath, char separator = ';', Encoding encoding = null, bool forceQuoteAll = false, int worksheetIndex = 0, bool exportAllSheets = false)
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
                    for (int i = 0; i < sheets.Count; i++)
                    {
                        string sheetName = SanitizeFileName(sheets[i].Name);
                        string outputPath = InsertSheetNameIntoPath(csvPath, sheetName);
                        ExportSheetToCsv(sheets[i], workbookPart, outputPath, separator, encoding, forceQuoteAll);
                        createdFiles.Add(outputPath);
                    }
                }
                else
                {
                    Sheet sheet = sheets.ElementAtOrDefault(worksheetIndex);
                    if (sheet == null)
                        throw new ArgumentException($"Worksheet mit Index {worksheetIndex} nicht gefunden.");

                    ExportSheetToCsv(sheet, workbookPart, csvPath, separator, encoding, forceQuoteAll);
                    createdFiles.Add(csvPath);
                }
            }

            return createdFiles;
        }

        private static string InsertSheetNameIntoPath(string csvPath, string sheetName)
        {
            string directory = Path.GetDirectoryName(csvPath);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(csvPath);
            string extension = Path.GetExtension(csvPath);
            string newFileName = $"{fileNameWithoutExt}_{sheetName}{extension}";
            return string.IsNullOrEmpty(directory) ? newFileName : Path.Combine(directory, newFileName);
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Sheet";
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            name = name.Replace(" ", "_").Replace(":", "_").Replace("/", "_").Replace("\\", "_");
            return name;
        }

        private static void ExportSheetToCsv(Sheet sheet, WorkbookPart workbookPart, string csvPath, char separator, Encoding encoding, bool forceQuoteAll)
        {
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;

            using (StreamWriter writer = new StreamWriter(csvPath, false, encoding))
            {
                foreach (Row row in sheetData.Elements<Row>())
                {
                    var cells = row.Elements<Cell>().ToList();
                    if (cells.Count == 0) continue;

                    // Maximale Spalte anhand CellReference bestimmen
                    int maxColumn = 0;
                    foreach (Cell cell in cells)
                    {
                        if (!string.IsNullOrEmpty(cell.CellReference))
                        {
                            int colIndex = GetColumnIndex(cell.CellReference);
                            if (colIndex > maxColumn) maxColumn = colIndex;
                        }
                    }

                    // Fallback wenn keine CellReferences vorhanden
                    if (maxColumn == 0 && cells.Count > 0)
                        maxColumn = cells.Count - 1;

                    List<string> cellValues = new List<string>();

                    for (int colIndex = 0; colIndex <= maxColumn; colIndex++)
                    {
                        string cellValue = "";
                        CellType cellType = CellType.Empty;

                        // Zelle anhand CellReference suchen - kein Fallback auf Index!
                        Cell cell = cells.FirstOrDefault(c => !string.IsNullOrEmpty(c.CellReference)
                                                           && GetColumnIndex(c.CellReference) == colIndex);

                        if (cell != null && cell.CellValue != null)
                            cellValue = GetCellValueWithType(cell, stringTablePart, workbookPart, out cellType);

                        cellValues.Add(FormatCsvField(cellValue, separator, forceQuoteAll, cellType));
                    }

                    writer.WriteLine(string.Join(separator.ToString(), cellValues));
                }
            }
        }

        private static int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return 0;

            string columnName = new string(cellReference.Where(c => char.IsLetter(c)).ToArray());
            if (string.IsNullOrEmpty(columnName)) return 0;

            int columnIndex = 0;
            for (int i = 0; i < columnName.Length; i++)
                columnIndex = columnIndex * 26 + (columnName[i] - 'A' + 1);

            return columnIndex - 1;
        }

        private static string GetCellValueWithType(Cell cell, SharedStringTablePart stringTablePart, WorkbookPart workbookPart, out CellType cellType)
        {
            cellType = CellType.Text;
            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                cellType = CellType.Text;
                if (stringTablePart != null)
                    return stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
                return value;
            }

            if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
            {
                cellType = CellType.Boolean;
                return value == "1" ? "TRUE" : "FALSE";
            }

            if (cell.DataType != null && cell.DataType.Value == CellValues.String)
            {
                cellType = CellType.Text;
                return value;
            }

            if (IsDateFormatted(cell, workbookPart))
            {
                cellType = CellType.Date;
                if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double oaDate))
                {
                    try { return DateTime.FromOADate(oaDate).ToString("yyyy-MM-dd"); }
                    catch { cellType = CellType.Number; return FormatNumber(value); }
                }
            }

            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
            {
                cellType = CellType.Number;
                return FormatNumber(value);
            }

            cellType = CellType.Text;
            return value;
        }

        private static bool IsDateFormatted(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.StyleIndex == null) return false;

            try
            {
                WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart;
                if (stylesPart?.Stylesheet?.CellFormats == null) return false;

                CellFormats cellFormats = stylesPart.Stylesheet.CellFormats;
                uint styleIndex = cell.StyleIndex.Value;
                if (styleIndex >= cellFormats.Count) return false;

                CellFormat cellFormat = (CellFormat)cellFormats.ElementAt((int)styleIndex);
                if (cellFormat.NumberFormatId == null) return false;

                uint numFmtId = cellFormat.NumberFormatId.Value;
                return (numFmtId >= 14 && numFmtId <= 22) || (numFmtId >= 45 && numFmtId <= 47) || numFmtId == 164;
            }
            catch { return false; }
        }

        private static string FormatNumber(string value)
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double numValue))
            {
                if (numValue == Math.Floor(numValue) && Math.Abs(numValue) < int.MaxValue)
                    return numValue.ToString("0", CultureInfo.GetCultureInfo("de-DE"));
                else
                    return numValue.ToString("0.##", CultureInfo.GetCultureInfo("de-DE"));
            }
            return value;
        }

        private static string FormatCsvField(string value, char separator, bool forceQuoteAll, CellType cellType)
        {
            if (string.IsNullOrEmpty(value))
                return "\"\"";

            bool needsQuotes = false;

            if (forceQuoteAll)
                needsQuotes = true;
            else if (value.Contains(separator) || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
                needsQuotes = true;
            else if (cellType == CellType.Text || cellType == CellType.Empty)
                needsQuotes = true;

            if (needsQuotes)
            {
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }

            return value;
        }
    }
}
