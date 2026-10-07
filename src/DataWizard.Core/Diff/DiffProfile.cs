using System.Text.Json;
using System.Text.Json.Serialization;

namespace DataWizard.Core.Diff;

/// <summary>
/// A saved comparison: the two files, the options used and a name. Stored as a
/// separate JSON file per comparison, so a set of regular comparisons is a set
/// of files that can live next to the data.
/// </summary>
public sealed class DiffProfile
{
    /// <summary>Schema version, so future releases can migrate old files.</summary>
    public int Version { get; set; } = 1;

    /// <summary>A user-chosen name for the comparison.</summary>
    public string? Name { get; set; }

    /// <summary>The reference file the comparison was last run against.</summary>
    public string? ReferencePath { get; set; }

    /// <summary>The candidate file the comparison was last run against.</summary>
    public string? CandidatePath { get; set; }

    /// <summary>The comparison options.</summary>
    public DiffOptions Options { get; set; } = new();

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>Loads a profile from disk.</summary>
    /// <exception cref="FileNotFoundException">The file does not exist.</exception>
    /// <exception cref="InvalidDataException">The file is not a valid profile.</exception>
    public static DiffProfile Load(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("Diff profile not found.", path);

        string json;

        try
        {
            json = File.ReadAllText(path);
        }
        catch (IOException ex)
        {
            throw new InvalidDataException($"Diff profile '{path}' could not be read: {ex.Message}", ex);
        }

        DiffProfile? profile;

        try
        {
            profile = JsonSerializer.Deserialize<DiffProfile>(json, JsonOptions);
        }
        catch (JsonException ex)
        {
            throw new InvalidDataException($"Diff profile '{path}' is not valid JSON: {ex.Message}", ex);
        }

        if (profile is null)
            throw new InvalidDataException($"Diff profile '{path}' is empty.", null);

        profile.Options ??= new DiffOptions();
        profile.Normalize();
        return profile;
    }

    /// <summary>
    /// Saves the profile to disk. The file is written to a temporary path first
    /// and then moved into place, so an interrupted save cannot leave a
    /// half-written profile behind.
    /// </summary>
    public void Save(string path)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);

        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        Normalize();

        var json = JsonSerializer.Serialize(this, JsonOptions);
        var temp = path + ".tmp";

        File.WriteAllText(temp, json);
        File.Move(temp, path, overwrite: true);
    }

    /// <summary>Clamps the options into a usable range.</summary>
    public void Normalize()
    {
        Options ??= new DiffOptions();
        Options.Normalize();
    }

    /// <summary>Creates a copy, so the UI can edit a profile without touching the saved one.</summary>
    public DiffProfile Clone() => new()
    {
        Version = Version,
        Name = Name,
        ReferencePath = ReferencePath,
        CandidatePath = CandidatePath,
        Options = Options.Clone()
    };
}
