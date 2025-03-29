using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace libDataWizard
{
    public class DataWizard
    {
        public event EventHandler   Update;

        protected virtual void OnUpdateEvent(UpdateEventArgs e)
        {
            EventHandler newUpdate = Update;
            if (newUpdate != null)
            {
                newUpdate(this, e);
            }
        }

        public void Load(String path)
        {
           
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

