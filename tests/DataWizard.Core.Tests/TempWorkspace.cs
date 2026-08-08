using System.Text;

namespace DataWizard.Core.Tests;

/// <summary>
/// A scratch directory that deletes itself, so tests can write real files without
/// leaving anything behind.
/// </summary>
public sealed class TempWorkspace : IDisposable
{
    public TempWorkspace()
    {
        Root = Path.Combine(Path.GetTempPath(), "DataWizardTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Root);
    }

    /// <summary>The directory tests may write into.</summary>
    public string Root { get; }

    /// <summary>Builds a path inside the workspace without creating anything.</summary>
    public string PathTo(string relative) => System.IO.Path.Combine(Root, relative);

    /// <summary>Writes a text file and returns its path.</summary>
    public string WriteFile(string name, string content, Encoding? encoding = null)
    {
        var path = PathTo(name);
        var directory = System.IO.Path.GetDirectoryName(path);

        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        File.WriteAllText(path, content, encoding ?? new UTF8Encoding(false));
        return path;
    }

    /// <summary>Writes a text file from lines joined with CRLF.</summary>
    public string WriteLines(string name, params string[] lines) =>
        WriteFile(name, string.Join("\r\n", lines));

    /// <summary>Creates a subdirectory and returns its path.</summary>
    public string CreateDirectory(string name)
    {
        var path = PathTo(name);
        Directory.CreateDirectory(path);
        return path;
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(Root))
                Directory.Delete(Root, recursive: true);
        }
        catch (IOException)
        {
            // A file still held open by the OS is not worth failing a test over.
        }
    }
}
