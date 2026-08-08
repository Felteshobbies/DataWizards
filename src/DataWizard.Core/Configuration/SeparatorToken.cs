namespace DataWizard.Core.Configuration;

/// <summary>
/// Converts between separator characters and the readable tokens used in
/// settings files and drop-down lists. Storing a raw <see cref="char"/> would
/// serialise tab as an invisible character, which is unpleasant to edit by hand.
/// </summary>
public static class SeparatorToken
{
    public const string Auto = "Auto";

    private static readonly (string Token, char Value)[] Known =
    [
        ("Semicolon", ';'),
        ("Comma", ','),
        ("Tab", '\t'),
        ("Pipe", '|'),
        ("Colon", ':'),
        ("Space", ' ')
    ];

    /// <summary>Tokens offered in the UI, excluding <see cref="Auto"/>.</summary>
    public static IReadOnlyList<string> KnownTokens { get; } = Known.Select(k => k.Token).ToArray();

    /// <summary>
    /// Parses a token such as <c>Semicolon</c>, <c>Tab</c>, <c>\t</c> or a bare
    /// character. Returns <c>null</c> for <see cref="Auto"/>, empty input, or
    /// anything unrecognised.
    /// </summary>
    public static char? Parse(string? token)
    {
        if (string.IsNullOrWhiteSpace(token))
            return null;

        token = token.Trim();

        if (token.Equals(Auto, StringComparison.OrdinalIgnoreCase))
            return null;

        foreach (var (name, value) in Known)
        {
            if (token.Equals(name, StringComparison.OrdinalIgnoreCase))
                return value;
        }

        // Escape sequences, so a hand-edited settings file can say "\t".
        switch (token)
        {
            case "\\t": return '\t';
            case "\\s": return ' ';
        }

        return token.Length == 1 ? token[0] : null;
    }

    /// <summary>Formats a separator as a readable token, or <see cref="Auto"/> for <c>null</c>.</summary>
    public static string Format(char? separator)
    {
        if (separator is null)
            return Auto;

        foreach (var (name, value) in Known)
        {
            if (value == separator.Value)
                return name;
        }

        return separator.Value.ToString();
    }

    /// <summary>
    /// A label suited to a drop-down, for example <c>Semicolon (;)</c>.
    /// </summary>
    public static string Describe(char? separator)
    {
        if (separator is null)
            return Auto;

        var token = Format(separator);
        return separator.Value == '\t' ? "Tab (\\t)" : $"{token} ({separator.Value})";
    }

    /// <summary>
    /// Expands a compact candidate string (for example <c>";,\t|"</c>) into
    /// distinct characters, honouring <c>\t</c> escape sequences.
    /// </summary>
    public static char[] ParseCandidates(string? candidates)
    {
        if (string.IsNullOrEmpty(candidates))
            return [';', ',', '\t', '|'];

        var result = new List<char>();

        for (var i = 0; i < candidates.Length; i++)
        {
            var c = candidates[i];

            if (c == '\\' && i + 1 < candidates.Length)
            {
                var next = candidates[i + 1];
                if (next is 't' or 's')
                {
                    result.Add(next == 't' ? '\t' : ' ');
                    i++;
                    continue;
                }
            }

            result.Add(c);
        }

        return result.Distinct().ToArray();
    }
}
