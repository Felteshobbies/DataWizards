using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace libDataWizard
{
    /// <summary>
    /// Konfiguration für Feldnamen-Erkennung und Datentyp-Überschreibung
    /// </summary>
    [XmlRoot("DataWizardConfig")]
    public class DataWizardConfig
    {
        /// <summary>
        /// Liste von Feldnamen-Patterns für Header-Erkennung
        /// </summary>
        [XmlArray("HeaderFieldNames")]
        [XmlArrayItem("Field")]
        public List<HeaderFieldPattern> HeaderFieldNames { get; set; }

        /// <summary>
        /// Liste von Datentyp-Überschreibungen basierend auf Feldnamen
        /// </summary>
        [XmlArray("DataTypeOverrides")]
        [XmlArrayItem("Override")]
        public List<DataTypeOverride> DataTypeOverrides { get; set; }

        public DataWizardConfig()
        {
            HeaderFieldNames = new List<HeaderFieldPattern>();
            DataTypeOverrides = new List<DataTypeOverride>();
        }

        /// <summary>
        /// Lädt die Konfiguration aus einer XML-Datei
        /// </summary>
        public static DataWizardConfig Load(string configPath)
        {
            if (!File.Exists(configPath))
            {
                // Erstelle Default-Konfiguration wenn nicht vorhanden
                var defaultConfig = CreateDefault();
                defaultConfig.Save(configPath);
                return defaultConfig;
            }

            using (StreamReader reader = new StreamReader(configPath))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(DataWizardConfig));
                return (DataWizardConfig)serializer.Deserialize(reader);
            }
        }

        /// <summary>
        /// Speichert die Konfiguration in eine XML-Datei
        /// </summary>
        public void Save(string configPath)
        {
            using (StreamWriter writer = new StreamWriter(configPath))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(DataWizardConfig));
                serializer.Serialize(writer, this);
            }
        }

        /// <summary>
        /// Erstellt eine Default-Konfiguration mit typischen Feldnamen
        /// </summary>
        public static DataWizardConfig CreateDefault()
        {
            var config = new DataWizardConfig();

            // Typische Header-Feldnamen (case-insensitive)
            config.HeaderFieldNames.AddRange(new[]
            {
                // IDs und Nummern
                new HeaderFieldPattern { Pattern = ".*id.*", IsRegex = true },
                new HeaderFieldPattern { Pattern = ".*nr.*", IsRegex = true },
                new HeaderFieldPattern { Pattern = ".*no.*", IsRegex = true },
                new HeaderFieldPattern { Pattern = ".*nummer.*", IsRegex = true },
                new HeaderFieldPattern { Pattern = ".*number.*", IsRegex = true },
                
                
                // Artikel/Produkt
                new HeaderFieldPattern { Pattern = "artikel", IsRegex = false },
                new HeaderFieldPattern { Pattern = "article", IsRegex = false },
                new HeaderFieldPattern { Pattern = "produkt", IsRegex = false },
                new HeaderFieldPattern { Pattern = "product", IsRegex = false },
                new HeaderFieldPattern { Pattern = "part", IsRegex = false },
                new HeaderFieldPattern { Pattern = "teilnummer", IsRegex = false },
                new HeaderFieldPattern { Pattern = "partno", IsRegex = false },
                new HeaderFieldPattern { Pattern = "part-no", IsRegex = false },
                new HeaderFieldPattern { Pattern = "sku", IsRegex = false },
                
                // Namen
                new HeaderFieldPattern { Pattern = "name", IsRegex = false },
                new HeaderFieldPattern { Pattern = "bezeichnung", IsRegex = false },
                new HeaderFieldPattern { Pattern = "description", IsRegex = false },
                new HeaderFieldPattern { Pattern = "titel", IsRegex = false },
                new HeaderFieldPattern { Pattern = "title", IsRegex = false },
                
                // Preis/Finanz
                new HeaderFieldPattern { Pattern = "preis", IsRegex = false },
                new HeaderFieldPattern { Pattern = "price", IsRegex = false },
                new HeaderFieldPattern { Pattern = "betrag", IsRegex = false },
                new HeaderFieldPattern { Pattern = "amount", IsRegex = false },
                new HeaderFieldPattern { Pattern = "kosten", IsRegex = false },
                new HeaderFieldPattern { Pattern = "cost", IsRegex = false },
                
                // Datum
                new HeaderFieldPattern { Pattern = "datum", IsRegex = false },
                new HeaderFieldPattern { Pattern = "date", IsRegex = false },
                new HeaderFieldPattern { Pattern = "^von$", IsRegex = true },
                new HeaderFieldPattern { Pattern = "^bis$", IsRegex = true },
                new HeaderFieldPattern { Pattern = "^from$", IsRegex = true },
                new HeaderFieldPattern { Pattern = "^to$", IsRegex = true },
                
                // Adressen
                new HeaderFieldPattern { Pattern = "strasse", IsRegex = false },
                new HeaderFieldPattern { Pattern = "street", IsRegex = false },
                new HeaderFieldPattern { Pattern = "ort", IsRegex = false },
                new HeaderFieldPattern { Pattern = "city", IsRegex = false },
                new HeaderFieldPattern { Pattern = "plz", IsRegex = false },
                new HeaderFieldPattern { Pattern = "postleitzahl", IsRegex = false },
                new HeaderFieldPattern { Pattern = "zip", IsRegex = false },
                new HeaderFieldPattern { Pattern = "postal", IsRegex = false },
                
                // Kontakt
                new HeaderFieldPattern { Pattern = "email", IsRegex = false },
                new HeaderFieldPattern { Pattern = "e-mail", IsRegex = false },
                new HeaderFieldPattern { Pattern = "mail", IsRegex = false },
                new HeaderFieldPattern { Pattern = "telefon", IsRegex = false },
                new HeaderFieldPattern { Pattern = "phone", IsRegex = false },
                new HeaderFieldPattern { Pattern = "tel", IsRegex = false },
                
                // Mengen
                new HeaderFieldPattern { Pattern = "menge", IsRegex = false },
                new HeaderFieldPattern { Pattern = "quantity", IsRegex = false },
                new HeaderFieldPattern { Pattern = "anzahl", IsRegex = false },
                new HeaderFieldPattern { Pattern = "count", IsRegex = false },
                new HeaderFieldPattern { Pattern = "qty", IsRegex = false },
            });

            // Datentyp-Überschreibungen
            config.DataTypeOverrides.AddRange(new[]
            {
                // IDs immer als Text (führende Nullen beibehalten)
                new DataTypeOverride { FieldNamePattern = "^id$", IsRegex = true, DataType = FieldDataType.Text },
                new DataTypeOverride { FieldNamePattern = ".*id$", IsRegex = true, DataType = FieldDataType.Text },
                new DataTypeOverride { FieldNamePattern = "kundennummer", IsRegex = false, DataType = FieldDataType.Text },
                new DataTypeOverride { FieldNamePattern = "customerno", IsRegex = false, DataType = FieldDataType.Text },
                new DataTypeOverride { FieldNamePattern = "matnr", IsRegex = false, DataType = FieldDataType.Text },
                new DataTypeOverride { FieldNamePattern = "part", IsRegex = false, DataType = FieldDataType.Text },
                new DataTypeOverride { FieldNamePattern = "artikel", IsRegex = false, DataType = FieldDataType.Text },
                
                // Preise als Decimal
                new DataTypeOverride { FieldNamePattern = "preis", IsRegex = false, DataType = FieldDataType.Decimal },
                new DataTypeOverride { FieldNamePattern = "price", IsRegex = false, DataType = FieldDataType.Decimal },
                new DataTypeOverride { FieldNamePattern = "betrag", IsRegex = false, DataType = FieldDataType.Decimal },
                new DataTypeOverride { FieldNamePattern = "amount", IsRegex = false, DataType = FieldDataType.Decimal },
                
                // Mengen als Integer
                new DataTypeOverride { FieldNamePattern = "menge", IsRegex = false, DataType = FieldDataType.Integer },
                new DataTypeOverride { FieldNamePattern = "quantity", IsRegex = false, DataType = FieldDataType.Integer },
                new DataTypeOverride { FieldNamePattern = "anzahl", IsRegex = false, DataType = FieldDataType.Integer },
                new DataTypeOverride { FieldNamePattern = "qty", IsRegex = false, DataType = FieldDataType.Integer },
                
                // Datum-Felder
                new DataTypeOverride { FieldNamePattern = "datum", IsRegex = false, DataType = FieldDataType.Date },
                new DataTypeOverride { FieldNamePattern = "date", IsRegex = false, DataType = FieldDataType.Date },
                new DataTypeOverride { FieldNamePattern = "geburtsdatum", IsRegex = false, DataType = FieldDataType.Date },
                new DataTypeOverride { FieldNamePattern = "birthdate", IsRegex = false, DataType = FieldDataType.Date },
                
                // PLZ immer als Text
                new DataTypeOverride { FieldNamePattern = "plz", IsRegex = false, DataType = FieldDataType.Text },
                new DataTypeOverride { FieldNamePattern = "postleitzahl", IsRegex = false, DataType = FieldDataType.Text },
                new DataTypeOverride { FieldNamePattern = "zip", IsRegex = false, DataType = FieldDataType.Text },
            });

            return config;
        }

        /// <summary>
        /// Prüft ob ein Feldname einem Header-Pattern entspricht
        /// </summary>
        public bool IsHeaderFieldName(string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
                return false;

            fieldName = fieldName.Trim().ToLower();

            foreach (var pattern in HeaderFieldNames)
            {
                if (pattern.IsRegex)
                {
                    try
                    {
                        if (Regex.IsMatch(fieldName, pattern.Pattern, RegexOptions.IgnoreCase))
                            return true;
                    }
                    catch
                    {
                        // Ungültiger Regex - ignorieren
                    }
                }
                else
                {
                    if (fieldName.Contains(pattern.Pattern.ToLower()))
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Gibt den Override-Datentyp für ein Feld zurück (falls vorhanden)
        /// </summary>
        public FieldDataType? GetDataTypeOverride(string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
                return null;

            fieldName = fieldName.Trim().ToLower();

            foreach (var dataTypeOverride in DataTypeOverrides)
            {
                if (dataTypeOverride.IsRegex)
                {
                    try
                    {
                        if (Regex.IsMatch(fieldName, dataTypeOverride.FieldNamePattern, RegexOptions.IgnoreCase))
                            return dataTypeOverride.DataType;
                    }
                    catch
                    {
                        // Ungültiger Regex - ignorieren
                    }
                }
                else
                {
                    if (fieldName.Contains(dataTypeOverride.FieldNamePattern.ToLower()))
                        return dataTypeOverride.DataType;
                }
            }

            return null;
        }
    }

    /// <summary>
    /// Pattern für Header-Feldnamen
    /// </summary>
    public class HeaderFieldPattern
    {
        /// <summary>
        /// Pattern (entweder String-Contain oder Regex)
        /// </summary>
        [XmlAttribute]
        public string Pattern { get; set; }

        /// <summary>
        /// Ist Pattern ein regulärer Ausdruck?
        /// </summary>
        [XmlAttribute]
        public bool IsRegex { get; set; }

        public HeaderFieldPattern()
        {
            Pattern = "";
            IsRegex = false;
        }
    }

    /// <summary>
    /// Datentyp-Überschreibung für bestimmte Feldnamen
    /// </summary>
    public class DataTypeOverride
    {
        /// <summary>
        /// Feldname-Pattern (entweder String-Contain oder Regex)
        /// </summary>
        [XmlAttribute]
        public string FieldNamePattern { get; set; }

        /// <summary>
        /// Ist Pattern ein regulärer Ausdruck?
        /// </summary>
        [XmlAttribute]
        public bool IsRegex { get; set; }

        /// <summary>
        /// Zu verwendender Datentyp
        /// </summary>
        [XmlAttribute]
        public FieldDataType DataType { get; set; }

        public DataTypeOverride()
        {
            FieldNamePattern = "";
            IsRegex = false;
            DataType = FieldDataType.Text;
        }
    }

    /// <summary>
    /// Datentypen für Excel-Felder
    /// </summary>
    public enum FieldDataType
    {
        /// <summary>
        /// Automatische Erkennung (Standard)
        /// </summary>
        Auto,

        /// <summary>
        /// Text/String
        /// </summary>
        Text,

        /// <summary>
        /// Ganzzahl
        /// </summary>
        Integer,

        /// <summary>
        /// Dezimalzahl
        /// </summary>
        Decimal,

        /// <summary>
        /// Datum
        /// </summary>
        Date
    }
}
