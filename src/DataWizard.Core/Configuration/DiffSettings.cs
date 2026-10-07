namespace DataWizard.Core.Configuration;

/// <summary>
/// Bookkeeping for the Diff tab: which profile was used last and which profiles
/// were used recently. The profiles themselves are separate files.
/// </summary>
public sealed class DiffSettings
{
    /// <summary>How many recent profiles are remembered.</summary>
    public const int MaxRecentProfiles = 10;

    /// <summary>The profile file the Diff tab last worked with.</summary>
    public string? LastProfilePath { get; set; }

    /// <summary>Recently used profile files, newest first.</summary>
    public List<string> RecentProfilePaths { get; set; } = [];

    /// <summary>
    /// Remembers a profile: it becomes the last one and moves to the front of
    /// the recent list.
    /// </summary>
    public void RememberProfile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;

        LastProfilePath = path;

        var recent = new List<string> { path };

        foreach (var existing in RecentProfilePaths)
        {
            if (!string.Equals(existing, path, StringComparison.OrdinalIgnoreCase))
                recent.Add(existing);
        }

        RecentProfilePaths = recent.Take(MaxRecentProfiles).ToList();
    }

    /// <summary>Clamps the section into a usable range.</summary>
    public void Normalize()
    {
        RecentProfilePaths ??= [];
        RecentProfilePaths = RecentProfilePaths
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Take(MaxRecentProfiles)
            .ToList();
    }

    /// <summary>Creates a copy, so the settings UI can cancel out of edits.</summary>
    public DiffSettings Clone() => new()
    {
        LastProfilePath = LastProfilePath,
        RecentProfilePaths = [.. RecentProfilePaths]
    };
}
