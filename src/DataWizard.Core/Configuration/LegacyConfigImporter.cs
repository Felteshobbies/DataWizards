using System.Xml.Linq;

namespace DataWizard.Core.Configuration;

/// <summary>
/// Reads the XML configuration used before version 0.2 so existing pattern lists
/// are not lost on upgrade.
/// </summary>
/// <remarks>
/// The old format stored a single <c>IsRegex</c> flag rather than a match mode,
/// and the shipped file and the README disagreed on the element names
/// (<c>&lt;Field Pattern="..."/&gt;</c> versus <c>&lt;Pattern Name="..."/&gt;</c>).
/// Both spellings are accepted here.
/// </remarks>
public static class LegacyConfigImporter
{
    /// <summary>The result of an import attempt.</summary>
    /// <param name="HeaderPatterns">Header field names found in the file.</param>
    /// <param name="FieldRules">Data type overrides found in the file.</param>
    /// <param name="Warnings">Entries that could not be converted.</param>
    public readonly record struct Result(
        List<NamePattern> HeaderPatterns,
        List<FieldRule> FieldRules,
        List<string> Warnings);

    /// <summary>
    /// Parses a legacy <c>DataWizard.config.xml</c>.
    /// </summary>
    /// <exception cref="FileNotFoundException">The file does not exist.</exception>
    /// <exception cref="System.Xml.XmlException">The file is not well-formed XML.</exception>
    public static Result ImportFile(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("Legacy configuration file not found.", path);

        return Import(XDocument.Load(path));
    }

    /// <summary>Parses legacy configuration from an already-loaded document.</summary>
    public static Result Import(XDocument document)
    {
        var headerPatterns = new List<NamePattern>();
        var fieldRules = new List<FieldRule>();
        var warnings = new List<string>();

        var root = document.Root;
        if (root is null)
        {
            warnings.Add("The document has no root element.");
            return new Result(headerPatterns, fieldRules, warnings);
        }

        foreach (var element in root.Elements()
                     .Where(e => e.Name.LocalName.Equals("HeaderFieldNames", StringComparison.OrdinalIgnoreCase))
                     .Elements())
        {
            var pattern = Attribute(element, "Pattern") ?? Attribute(element, "Name");

            if (string.IsNullOrWhiteSpace(pattern))
            {
                warnings.Add($"Skipped a <{element.Name.LocalName}> entry without a pattern.");
                continue;
            }

            headerPatterns.Add(new NamePattern
            {
                Pattern = pattern,
                Match = ReadBool(element, "IsRegex") ? MatchMode.Regex : MatchMode.Contains,
                Comment = "Imported from legacy configuration"
            });
        }

        foreach (var element in root.Elements()
                     .Where(e => e.Name.LocalName.Equals("DataTypeOverrides", StringComparison.OrdinalIgnoreCase))
                     .Elements())
        {
            var pattern = Attribute(element, "FieldNamePattern")
                          ?? Attribute(element, "Pattern")
                          ?? Attribute(element, "Name");

            if (string.IsNullOrWhiteSpace(pattern))
            {
                warnings.Add($"Skipped a <{element.Name.LocalName}> entry without a pattern.");
                continue;
            }

            var rawType = Attribute(element, "DataType");
            if (!Enum.TryParse<FieldDataType>(rawType, ignoreCase: true, out var dataType))
            {
                warnings.Add($"Unknown data type '{rawType}' for pattern '{pattern}'; imported as Text.");
                dataType = FieldDataType.Text;
            }

            fieldRules.Add(new FieldRule
            {
                Pattern = pattern,
                Match = ReadBool(element, "IsRegex") ? MatchMode.Regex : MatchMode.Contains,
                DataType = dataType,
                Comment = "Imported from legacy configuration"
            });
        }

        if (headerPatterns.Count == 0 && fieldRules.Count == 0)
            warnings.Add("No header patterns or data type overrides were found in the file.");

        return new Result(headerPatterns, fieldRules, warnings);
    }

    /// <summary>
    /// Imports a legacy file into <paramref name="settings"/>, replacing the
    /// current pattern lists.
    /// </summary>
    public static Result ImportInto(DataWizardSettings settings, string path)
    {
        var result = ImportFile(path);

        if (result.HeaderPatterns.Count > 0)
            settings.Detection.KnownFieldNames = result.HeaderPatterns;

        if (result.FieldRules.Count > 0)
            settings.Detection.FieldRules = result.FieldRules;

        return result;
    }

    private static string? Attribute(XElement element, string name) =>
        element.Attributes()
            .FirstOrDefault(a => a.Name.LocalName.Equals(name, StringComparison.OrdinalIgnoreCase))
            ?.Value;

    private static bool ReadBool(XElement element, string name) =>
        bool.TryParse(Attribute(element, name), out var value) && value;
}
