using System;
using System.Collections.Generic;
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

        public void Load(String path)
        {
            // charset detection
            DetectionResult detectionResult = CharsetDetector.DetectFromFile(path);
            DetectedEncoding = detectionResult.Detected.Encoding;

            using (StreamReader reader = new StreamReader(path, DetectedEncoding ))
            {
                Analyze(reader);

                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var col = line.Split(Separator);
                    Console.WriteLine(col[0]);
                }
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

        public class UpdateEventArgs : EventArgs
        {
            public string Text;

            public UpdateEventArgs(string text)
            {
                this.Text = text;
            }

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

