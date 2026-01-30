using libDataWizard;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using UtfUnknown;


namespace DataWizard
{
    public class CSV
    {
        public string filePath { get; set; }

        public event EventHandler Update;
        public char[] Separators = new[] { ',', ';', '\t', '|' };
        public char Separator;
        public double SeparatorProbability;
        private int[] _separatorCount;

        public int FieldCount;
        public int StartLine;
        public int EndLine;
        public int Lines;
        public int MaxLinesAnalyze = 100;
        public bool IsFieldCountEqual;
        public bool DetectedHeaderLine;
        public Encoding DetectedEncoding = Encoding.UTF8;

        // Konfiguration für Feldnamen und Datentypen
        public DataWizardConfig Config { get; set; }
        private string[] headerFieldNames; // Erkannte Header-Feldnamen

        private string fieldNames = "no nr article part partno part-no price name id date plz ort street strasse email e-mail";

        protected virtual void OnUpdateEvent(UpdateEventArgs e)
        {
            EventHandler newUpdate = Update;
            if (newUpdate != null)
            {
                newUpdate(this, e);
            }
        }

        public CSV()
        {
            Clear();
            _separatorCount = new int[Separators.Length];
            
            // Lade Standard-Konfiguration
            LoadDefaultConfig();
        }

        public CSV(string configPath)
        {
            Clear();
            _separatorCount = new int[Separators.Length];
            
            // Lade Konfiguration aus Datei
            LoadConfig(configPath);
        }

        /// <summary>
        /// Lädt die Standard-Konfiguration
        /// </summary>
        private void LoadDefaultConfig()
        {
            Config = DataWizardConfig.CreateDefault();
        }

        /// <summary>
        /// Lädt Konfiguration aus einer XML-Datei
        /// </summary>
        public void LoadConfig(string configPath)
        {
            try
            {
                Config = DataWizardConfig.Load(configPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fehler beim Laden der Konfiguration: {ex.Message}");
                LoadDefaultConfig();
            }
        }

        public void Load(String filePath)
        {
            this.filePath = filePath;

            // charset detection
            DetectionResult detectionResult = CharsetDetector.DetectFromFile(this.filePath);
            DetectedEncoding = detectionResult.Detected.Encoding;

            using (StreamReader reader = new StreamReader(this.filePath, DetectedEncoding))
            {
                Analyze(reader);
            }
        }

        public void Analyze(StreamReader reader)
        {
            string firstLine = "";
            string secondline = "";

            var separatorCount = new Dictionary<char, int>();
            var fieldCountList = new List<int>(); // Speichert die Anzahl der Felder pro Zeile

            foreach (char separator in Separators)
            {
                separatorCount[separator] = 0;
            }

            bool isContent = false;
            string line;
            int lineCounter = 0; // Separater Zähler für die erste Schleife

            // Erste Schleife: Separator-Erkennung
            while ((line = reader.ReadLine()) != null && lineCounter < MaxLinesAnalyze)
            {
                lineCounter++;
                
                // Empty lines?
                if (!isContent && !String.IsNullOrWhiteSpace(line))
                {
                    StartLine = lineCounter;
                    isContent = true;
                    firstLine = line; // Erste Zeile mit Inhalt speichern
                }
                else if (isContent && String.IsNullOrWhiteSpace(firstLine) == false && String.IsNullOrWhiteSpace(secondline))
                {
                    secondline = line; // Zweite Zeile mit Inhalt speichern
                }

                if (isContent)
                {
                    // count separators
                    foreach (char separator in Separators)
                    {
                        separatorCount[separator] += line.Count(c => c == separator);
                    }
                }
            }

            Lines = lineCounter;
            EndLine = lineCounter;

            // Besten Separator ermitteln
            char bestSeparator = Separator;
            int highestCount = 0;
            int totalCount = 0;

            foreach (var kvp in separatorCount)
            {
                totalCount += kvp.Value;
                if (kvp.Value > highestCount)
                {
                    highestCount = kvp.Value;
                    bestSeparator = kvp.Key;
                }
            }

            // calculate probability percent
            double probability = totalCount > 0 ? (highestCount / (double)totalCount) * 100 : 0;

            Separator = bestSeparator;
            SeparatorProbability = probability;

            // detect for headerline
            DetectedHeaderLine = IsHeader(firstLine, secondline, Separator);

            // reset stream
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            reader.DiscardBufferedData();

            // Zweite Schleife: Field-Count-Analyse und detaillierte Auswertung
            lineCounter = 0;
            while ((line = reader.ReadLine()) != null && lineCounter < MaxLinesAnalyze)
            {
                lineCounter++;
                
                if (lineCounter >= StartLine) // Nur ab StartLine analysieren
                {
                    CsvField[] fields = SplitLine(line, Separator);
                    fieldCountList.Add(fields.Length);

                    // Debug-Ausgabe (optional, kann später entfernt werden)
                    Console.WriteLine($"Zeile {lineCounter}: {line}");
                    foreach (CsvField field in fields)
                    {
                        Console.WriteLine($"  Quoted: {field.Quotes}, Wert: '{field.Value}'");
                    }
                }
            }

            // Field-Count auswerten
            if (fieldCountList.Count > 0)
            {
                FieldCount = fieldCountList[0]; // Erste Zeile als Referenz
                IsFieldCountEqual = fieldCountList.All(count => count == FieldCount);
            }

            // Speichere Header-Feldnamen falls vorhanden
            if (DetectedHeaderLine && !String.IsNullOrWhiteSpace(firstLine))
            {
                headerFieldNames = firstLine.Split(Separator)
                    .Select(f => f.Trim().Trim('"'))
                    .ToArray();
            }
        }

        private bool IsHeader(string firstLine, string secondLine, char separator)
        {
            bool isHeader = true;
            try
            {
                // Wenn eine der Zeilen leer ist, können wir nicht entscheiden
                if (String.IsNullOrWhiteSpace(firstLine) || String.IsNullOrWhiteSpace(secondLine))
                {
                    return false;
                }

                string[] firstFields = firstLine.Split(separator);
                string[] secondFields = secondLine.Split(separator);

                // Wenn unterschiedliche Anzahl Felder, unsicher
                if (firstFields.Length != secondFields.Length)
                {
                   // return false;
                }

                int knownHeaderFieldsCount = 0;

                for (int i = 0; i < firstFields.Length; i++)
                {
                    string field1 = firstFields[i].Trim().Trim('"');
                    string field2 = secondFields[i].Trim().Trim('"');

                    // Prüfe ob Feld 1 einem bekannten Header-Feldnamen entspricht
                    if (Config != null && Config.IsHeaderFieldName(field1))
                    {
                        knownHeaderFieldsCount++;
                    }

                    // no header, if numbers exists in both first and second line
                    if (double.TryParse(field1, NumberStyles.Any, CultureInfo.InvariantCulture, out _) && 
                        double.TryParse(field2, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                    {
                        isHeader = false;
                    }
                }

                // Wenn mindestens 50% der Felder bekannte Header-Namen sind, ist es wahrscheinlich ein Header
                if (knownHeaderFieldsCount > 0)
                {
                    isHeader = true;
                }
            }
            catch (Exception)
            {
                return false;
            }
            return isHeader;
        }

        public void Clear()
        {
            Separator = ';';
            SeparatorProbability = 100;
            FieldCount = 0;
            StartLine = 0;
            EndLine = 0;
            Lines = 0;
            IsFieldCountEqual = true;
            DetectedHeaderLine = false;
        }

        public void WriteXLSX(string xlsxPath, bool overwrite)
        {
            using (XLS xls = new XLS(xlsxPath, overwrite))
            {
                // Übergebe Header-Feldnamen und Config an XLS
                if (DetectedHeaderLine && headerFieldNames != null)
                {
                    xls.SetHeaderFieldNames(headerFieldNames, Config);
                }

                using (StreamReader reader = new StreamReader(this.filePath, DetectedEncoding))
                {
                    string line;
                    int lineNumber = 0;
                    int currentRow = 1; // Excel-Zeilen beginnen bei 1

                    while ((line = reader.ReadLine()) != null)
                    {
                        lineNumber++;

                        // Überspringe Zeilen vor StartLine
                        if (lineNumber < StartLine)
                            continue;

                        CsvField[] fields = SplitLine(line, Separator);

                        // Füge Zeile zum Excel hinzu
                        xls.AddRow(fields, (uint)currentRow);
                        currentRow++;
                    }

                    // Speichere das Excel-Dokument vor dem Dispose
                    xls.Save();
                }
            } // Hier wird automatisch xls.Dispose() aufgerufen
        }

        public class UpdateEventArgs : EventArgs
        {
            public string Text;

            public UpdateEventArgs(string text)
            {
                this.Text = text;
            }
        }

        public CsvField[] SplitLine(string line, char separator)
        {
            List<CsvField> result = new List<CsvField>();
            int index = 0;

            while (index < line.Length)
            {
                if (line[index] == '"') // Falls das Feld mit Hochkommata beginnt
                {
                    int closingQuote = line.IndexOf('"', index + 1);
                    while (closingQuote != -1 && closingQuote + 1 < line.Length && line[closingQuote + 1] == '"')
                    {
                        // Doppelte Hochkommata (escaped quotes) überspringen
                        closingQuote = line.IndexOf('"', closingQuote + 2);
                    }

                    if (closingQuote == -1) closingQuote = line.Length; // Kein schließendes Hochkomma gefunden                   
                    result.Add(new CsvField { Value = line.Substring(index + 1, closingQuote - index - 1), Quotes = true });
                    index = closingQuote + 1;

                    // Überspringe das Trennzeichen nach dem geschlossenen Hochkomma
                    if (index < line.Length && line[index] == separator)
                    {
                        index++;
                    }
                }
                else
                {
                    // Suche nach dem nächsten Trennzeichen
                    int nextSeparator = line.IndexOf(separator, index);
                    if (nextSeparator == -1) nextSeparator = line.Length;

                    // Füge den aktuellen Wert hinzu (auch leere Werte)
                    string value = line.Substring(index, nextSeparator - index);
                    result.Add(new CsvField { Value = value, Quotes = false });
                    index = nextSeparator + 1; // Überspringe das Trennzeichen
                }
            }

            // Berücksichtige abschließende leere Felder
            if (line.Length > 0 && line.EndsWith(separator.ToString()))
            {
                result.Add(new CsvField { Value = "", Quotes = false });
            }

            return result.ToArray();
        }
    }

    public class CsvField
    {
        public String Value { get; set; }
        public Boolean Quotes { get; set; }

        public CsvField()
        {
            this.Quotes = false;
        }
    }
}
