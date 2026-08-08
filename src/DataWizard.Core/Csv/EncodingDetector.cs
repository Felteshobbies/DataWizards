using System.Text;
using DataWizard.Core.Configuration;
using UtfUnknown;

namespace DataWizard.Core.Csv;

/// <summary>The outcome of encoding detection for one file.</summary>
/// <param name="Encoding">The encoding to read the file with.</param>
/// <param name="Confidence">Detector confidence between 0 and 1.</param>
/// <param name="Source">How the encoding was arrived at.</param>
/// <param name="HasByteOrderMark">Whether the file starts with a byte order mark.</param>
public readonly record struct EncodingDetectionResult(
    Encoding Encoding,
    double Confidence,
    EncodingSource Source,
    bool HasByteOrderMark)
{
    /// <summary>A short description for the log.</summary>
    public string Describe() => Source switch
    {
        EncodingSource.Forced => $"{Encoding.WebName} (forced)",
        EncodingSource.ByteOrderMark => $"{Encoding.WebName} (byte order mark)",
        EncodingSource.Detected => $"{Encoding.WebName} ({Confidence:P0} confidence)",
        _ => $"{Encoding.WebName} (fallback)"
    };
}

/// <summary>How an encoding was decided on.</summary>
public enum EncodingSource
{
    /// <summary>Taken from the settings without inspecting the file.</summary>
    Forced,
    /// <summary>Read from the file's byte order mark.</summary>
    ByteOrderMark,
    /// <summary>Guessed from the byte statistics.</summary>
    Detected,
    /// <summary>Detection failed or was not confident enough.</summary>
    Fallback
}

/// <summary>
/// Works out which encoding a text file uses.
/// </summary>
public static class EncodingDetector
{
    private static bool _codePagesRegistered;

    /// <summary>
    /// Makes the legacy single-byte code pages (windows-1252, ISO-8859-15 and
    /// friends) available. .NET Core only ships a handful of encodings by
    /// default, so without this <c>Encoding.GetEncoding(1252)</c> throws.
    /// </summary>
    public static void EnsureCodePagesRegistered()
    {
        if (_codePagesRegistered)
            return;

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _codePagesRegistered = true;
    }

    /// <summary>
    /// Detects the encoding of <paramref name="path"/> honouring the rules in
    /// <paramref name="settings"/>.
    /// </summary>
    public static EncodingDetectionResult Detect(string path, DetectionSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        EnsureCodePagesRegistered();

        var bom = ReadByteOrderMark(path);

        if (!string.IsNullOrWhiteSpace(settings.ForcedEncoding))
        {
            var forced = TryGetEncoding(settings.ForcedEncoding);
            if (forced is not null)
                return new EncodingDetectionResult(forced, 1d, EncodingSource.Forced, bom is not null);
        }

        if (settings.TrustByteOrderMark && bom is not null)
            return new EncodingDetectionResult(bom, 1d, EncodingSource.ByteOrderMark, true);

        var fallback = TryGetEncoding(settings.FallbackEncoding) ?? new UTF8Encoding(false);

        try
        {
            var detection = CharsetDetector.DetectFromFile(path);
            var detected = detection?.Detected;

            // An empty or very short file gives the detector nothing to work with,
            // in which case Detected is null - the old code dereferenced it blindly.
            if (detected?.Encoding is not null && detected.Confidence >= settings.MinEncodingConfidence)
                return new EncodingDetectionResult(detected.Encoding, detected.Confidence, EncodingSource.Detected, bom is not null);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Fall through to the fallback encoding.
        }

        return new EncodingDetectionResult(fallback, 0d, EncodingSource.Fallback, bom is not null);
    }

    /// <summary>
    /// Returns the encoding described by a byte order mark, or <c>null</c> when the
    /// file does not start with one.
    /// </summary>
    public static Encoding? ReadByteOrderMark(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            Span<byte> buffer = stackalloc byte[4];
            var read = stream.ReadAtLeast(buffer, 4, throwOnEndOfStream: false);
            return ReadByteOrderMark(buffer[..read]);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return null;
        }
    }

    /// <summary>Returns the encoding described by the leading bytes of a buffer.</summary>
    public static Encoding? ReadByteOrderMark(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            return new UTF8Encoding(true);

        if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0x00 && bytes[3] == 0x00)
            return new UTF32Encoding(bigEndian: false, byteOrderMark: true);

        if (bytes.Length >= 4 && bytes[0] == 0x00 && bytes[1] == 0x00 && bytes[2] == 0xFE && bytes[3] == 0xFF)
            return new UTF32Encoding(bigEndian: true, byteOrderMark: true);

        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            return new UnicodeEncoding(bigEndian: false, byteOrderMark: true);

        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            return new UnicodeEncoding(bigEndian: true, byteOrderMark: true);

        return null;
    }

    /// <summary>
    /// Resolves an encoding by web name, code page number or display name.
    /// Returns <c>null</c> when nothing matches.
    /// </summary>
    public static Encoding? TryGetEncoding(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return null;

        EnsureCodePagesRegistered();
        name = name.Trim();

        if (int.TryParse(name, out var codePage))
        {
            try
            {
                return Encoding.GetEncoding(codePage);
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
            {
                return null;
            }
        }

        try
        {
            return Encoding.GetEncoding(name);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
        {
            return null;
        }
    }

    /// <summary>
    /// The encodings offered in the UI. Only those actually available on this
    /// machine are returned.
    /// </summary>
    public static IReadOnlyList<string> CommonEncodings()
    {
        EnsureCodePagesRegistered();

        string[] candidates =
        [
            "utf-8",
            "utf-16",
            "utf-16BE",
            "windows-1252",
            "iso-8859-1",
            "iso-8859-15",
            "ibm850",
            "us-ascii"
        ];

        return candidates.Where(c => TryGetEncoding(c) is not null).ToArray();
    }
}
