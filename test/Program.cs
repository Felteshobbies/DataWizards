using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using DataWizard;
using libDataWizard;


namespace test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== DataWizard Test ===\n");

            try
            {
                // Test 0: Konfiguration erstellen
                Console.WriteLine("--- Test 0: Konfiguration ---");
                string configPath = @"C:\pa-temp\DataWizard.config.xml";
                
                // Erstelle Default-Config falls nicht vorhanden
                if (!System.IO.File.Exists(configPath))
                {
                    var defaultConfig = DataWizardConfig.CreateDefault();
                    defaultConfig.Save(configPath);
                    Console.WriteLine($"Standard-Konfiguration erstellt: {configPath}");
                }
                else
                {
                    Console.WriteLine($"Konfiguration gefunden: {configPath}");
                }

                // Test 1: CSV laden und analysieren mit Config
                Console.WriteLine("\n--- Test 1: CSV-Analyse mit Config ---");
                CSV csv = new CSV(configPath); // Lädt Config aus Datei
                csv.Load(@"C:\pa-temp\test3.txt");

                Console.WriteLine($"Datei geladen: C:\\temp\\test3.txt");
                Console.WriteLine($"Erkanntes Encoding: {csv.DetectedEncoding.EncodingName}");
                Console.WriteLine($"Separator: '{csv.Separator}'");
                Console.WriteLine($"Separator-Wahrscheinlichkeit: {csv.SeparatorProbability:F2}%");
                Console.WriteLine($"Start-Zeile: {csv.StartLine}");
                Console.WriteLine($"End-Zeile: {csv.EndLine}");
                Console.WriteLine($"Gesamt-Zeilen: {csv.Lines}");
                Console.WriteLine($"Anzahl Felder: {csv.FieldCount}");
                Console.WriteLine($"Felder konsistent: {csv.IsFieldCountEqual}");
                Console.WriteLine($"Header erkannt: {csv.DetectedHeaderLine}");

                Console.WriteLine("\n--- Test 2: CSV zu Excel konvertieren (mit Config) ---");
                
                // Test 2: CSV nach Excel schreiben mit Datentyp-Überschreibungen
                string xlsxPath = @"C:\pa-temp\output2.xlsx";
                csv.WriteXLSX(xlsxPath, true);
                Console.WriteLine($"Excel-Datei erstellt: {xlsxPath}");
                Console.WriteLine("Hinweis: Datentypen wurden gemäß Config überschrieben");

                Console.WriteLine("\n--- Test 3: Direkte Excel-Erstellung ---");
                
                // Test 3: Direkt Excel erstellen (alter Test)
                using (XLS xls = new XLS(@"C:\pa-temp\test_direct.xlsx", true))
                {
                    // Beispieldaten hinzufügen
                    CsvField[] headerRow = new CsvField[]
                    {
                        new CsvField { Value = "Name", Quotes = false },
                        new CsvField { Value = "Alter", Quotes = false },
                        new CsvField { Value = "Stadt", Quotes = false }
                    };
                    xls.AddRow(headerRow, 1);

                    CsvField[] dataRow1 = new CsvField[]
                    {
                        new CsvField { Value = "Max Mustermann", Quotes = true },
                        new CsvField { Value = "25", Quotes = false },
                        new CsvField { Value = "Berlin", Quotes = false }
                    };
                    xls.AddRow(dataRow1, 2);

                    CsvField[] dataRow2 = new CsvField[]
                    {
                        new CsvField { Value = "Erika Musterfrau", Quotes = true },
                        new CsvField { Value = "30", Quotes = false },
                        new CsvField { Value = "München", Quotes = false }
                    };
                    xls.AddRow(dataRow2, 3);

                    xls.Save();
                } // Automatisches Dispose hier
                
                Console.WriteLine("Direkte Excel-Datei erstellt: C:\\temp\\test_direct.xlsx");

                Console.WriteLine("\n--- Test 4: Config-Funktionen testen ---");
                var testConfig = DataWizardConfig.Load(configPath);
                
                // Teste Header-Erkennung
                Console.WriteLine("\nHeader-Feldnamen-Tests:");
                string[] testFields = { "id", "name", "preis", "datum", "xyz123" };
                foreach (var field in testFields)
                {
                    bool isHeader = testConfig.IsHeaderFieldName(field);
                    Console.WriteLine($"  '{field}' -> Header: {isHeader}");
                }
                
                // Teste Datentyp-Überschreibungen
                Console.WriteLine("\nDatentyp-Überschreibungen:");
                string[] testDataTypeFields = { "kundennummer", "preis", "menge", "plz", "name" };
                foreach (var field in testDataTypeFields)
                {
                    var overrideType = testConfig.GetDataTypeOverride(field);
                    Console.WriteLine($"  '{field}' -> {(overrideType.HasValue ? overrideType.Value.ToString() : "Auto")}");
                }

                Console.WriteLine("\n--- Test 5: Excel zu CSV Konvertierung ---");
                
                // Erstelle zuerst eine Excel-Datei mit mehreren Worksheets für den Test
                string multiSheetPath = @"C:\pa-temp\multi_sheet.xlsx";
               
                Console.WriteLine($"Test-Excel mit mehreren Sheets erstellt: {multiSheetPath}");
                
                // Konvertiere das zuvor erspa-tellte Excel zurück zu CSV
                string csvOutputPath = @"C:\pa-temp\from_excel.csv";
                
                // Standard: UTF-8, Semikolon, Text in Quotes
                XLS.ToCsv(@"C:\pa-temp\output.xlsx", csvOutputPath);
                Console.WriteLine($"CSV erstellt (UTF-8, Semikolon): {csvOutputPath}");
                
                // Mit verschiedenen Optionen
                string csvOutputPath2 = @"C:\pa-temp\from_excel_comma.csv";
                XLS.ToCsv(@"C:\pa-temp\output.xlsx", csvOutputPath2, separator: ',', encoding: Encoding.UTF8, quoteAllText: false);
                Console.WriteLine($"CSV erstellt (UTF-8, Komma, minimal quotes): {csvOutputPath2}");
                
                // Windows-1252 Encoding
                string csvOutputPath3 = @"C:\pa-temp\from_excel_ansi.csv";
                XLS.ToCsv(@"C:\pa-temp\output.xlsx", csvOutputPath3, separator: ';', encoding: Encoding.GetEncoding(1252), quoteAllText: true);
                Console.WriteLine($"CSV erstellt (Windows-1252, Semikolon): {csvOutputPath3}");

                // ALLE Worksheets exportieren
                Console.WriteLine("\n--- Test 6: Alle Worksheets exportieren ---");
                string csvMultiBase = @"C:\pa-temp\multi_output.csv";
                List<string> createdFiles = XLS.ToCsv(multiSheetPath, csvMultiBase, exportAllSheets: true);
                
                Console.WriteLine($"Alle Worksheets exportiert ({createdFiles.Count} Dateien):");
                foreach (string file in createdFiles)
                {
                    Console.WriteLine($"  - {Path.GetFileName(file)}");
                }

                Console.WriteLine("\n=== Alle Tests erfolgreich! ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nFEHLER: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }

            Console.WriteLine("\nDrücke eine Taste zum Beenden...");
            Console.ReadKey();
        }

    
    }
}
