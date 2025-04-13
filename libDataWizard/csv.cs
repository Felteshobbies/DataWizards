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

                //string line;
                //while ((line = reader.ReadLine()) != null)
                //{
                //    var col = line.Split(Separator);
                //    Console.WriteLine(col[0]);
                //}
            }
        }

        public void Analyze(StreamReader reader)
        {
            string firstLine = "";
            string secondline = "";


            var separatorCount = new Dictionary<char, int>();
            // var fieldCount = new Dictionary<char, List<int>>();
            foreach (char separator in Separators)
            {
                separatorCount[separator] = 0;
            }

            bool isContent = false;
            string line;
            while ((line = reader.ReadLine()) != null && Lines < MaxLinesAnalyze)
            {
                // Empty lines?
                Lines++;
                if (!isContent && !String.IsNullOrWhiteSpace(line))
                {
                    StartLine = Lines;
                    isContent = true;
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

            while ((line = reader.ReadLine()) != null && Lines < MaxLinesAnalyze)
            {
                Console.WriteLine(line);
                string[] res = SplitLine(line, Separator);
                foreach (string s in res)
                {
                    Console.WriteLine(s);
                }


            }
        }

        private bool IsHeader(string firstLine, string secondLine, char separator)
        {
            bool isHeader = true;
            try
            {
                string[] firstFields = firstLine.Split(separator);
                string[] secondFields = secondLine.Split(separator);

                for (int i = 0; i < firstFields.Length; i++)
                {
                    //no header, if numbers exists
                    if (double.TryParse(firstFields[i], out _) && double.TryParse(secondFields[i], out _)) isHeader = false;
                }


            }
            catch (Exception)
            {
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

        public void WriteXLSX(bool overwrite)
        {
            XLS xls = new XLS(this.filePath, overwrite);

            using (StreamReader reader = new StreamReader(this.filePath, DetectedEncoding))
            {


                //string line;
                //while ((line = reader.ReadLine()) != null)
                //{
                //    var col = line.Split(Separator);
                //    Console.WriteLine(col[0]);
                //}
            }
        }

        public class UpdateEventArgs : EventArgs
        {
            public string Text;

            public UpdateEventArgs(string text)
            {
                this.Text = text;
            }

        }



        public static string[] SplitLine(string line, char separator)
        {
            List<string> result = new List<string>();
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
                    result.Add(line.Substring(index + 1, closingQuote - index - 1));
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
                    result.Add(line.Substring(index, nextSeparator - index));
                    index = nextSeparator + 1; // Überspringe das Trennzeichen
                }
            }

            // Berücksichtige abschließende leere Felder
            if (line.EndsWith(separator.ToString()))
            {
                result.Add("");
            }

            return result.ToArray();
        }











    }

    public class Field
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public string Value { get; set; }
        public Type Type { get; set; }

        public Field() {
            this.Type = typeof(String);
        }
    }
}



public class CsvAnalyzer
{
    public class CsvProperties
    {
        public bool IsCSV { get; set; }
        public char? Separator { get; set; }
        public string Charset { get; set; }
        public int FieldCount { get; set; }
        public bool IsHeader { get; set; }
        public List<string> HeaderFields { get; set; } = new List<string>();
        public bool SkipFirst { get; set; }
        public bool SkipLast { get; set; }
        public Dictionary<int, string> FieldFormats { get; set; } = new Dictionary<int, string>();
    }

    public CsvProperties AnalyzeFile(string filePath)
    {
        var properties = new CsvProperties();

        try
        {
            // Bestimme Charset
            properties.Charset = DetectCharset(filePath);

            // Lese Datei
            var lines = File.ReadAllLines(filePath, Encoding.GetEncoding(properties.Charset));

            if (lines.Length == 0)
            {
                properties.IsCSV = false;
                return properties;
            }

            // Bestimme Trennzeichen
            var possibleSeparators = new[] { ',', ';', '\t', '|' };
            properties.Separator = DetectSeparator(lines, possibleSeparators);

            // Prüfe CSV-Eigenschaften
            properties.IsCSV = properties.Separator.HasValue;

            if (!properties.IsCSV) return properties;

            // Überprüfe Felder und Header
            var firstLine = lines[0];
            properties.FieldCount = firstLine.Split(properties.Separator.Value).Length;

            // Header-Analyse
            properties.IsHeader = IsValidHeader(firstLine, properties.Separator.Value);
            if (properties.IsHeader)
            {
                properties.HeaderFields = firstLine.Split(properties.Separator.Value).ToList();
            }

            // Überspringe ungültige Zeilen
            properties.SkipFirst = !properties.IsHeader;
            properties.SkipLast = lines.Any(line => string.IsNullOrWhiteSpace(line));

            // Bestimme Feldformate
            properties.FieldFormats = DetectFieldFormats(lines.Skip(properties.SkipFirst ? 1 : 0).ToArray(), properties.Separator.Value);

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler bei der Analyse: {ex.Message}");
        }

        return properties;
    }

    private string DetectCharset(string filePath)
    {
        // Beispiel für die Charset-Erkennung (Vereinfacht)
        return "UTF-8"; // Annahme: UTF-8, kann erweitert werden
    }

    private char? DetectSeparator(string[] lines, char[] possibleSeparators)
    {
        foreach (var separator in possibleSeparators)
        {
            if (lines.All(line => line.Contains(separator)))
            {
                return separator;
            }
        }
        return null;
    }

    private bool IsValidHeader(string line, char separator)
    {
        return line.Split(separator).All(field => !string.IsNullOrWhiteSpace(field));
    }

    private Dictionary<int, string> DetectFieldFormats(string[] lines, char separator)
    {
        var fieldFormats = new Dictionary<int, string>();
        foreach (var line in lines)
        {
            var fields = line.Split(separator);
            for (int i = 0; i < fields.Length; i++)
            {
                if (!fieldFormats.ContainsKey(i))
                {
                    fieldFormats[i] = fields[i].StartsWith("\"") && fields[i].EndsWith("\"") ? "String" : "Unknown";
                }
            }
        }
        return fieldFormats;
    }
}

