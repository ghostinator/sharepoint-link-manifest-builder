using System.Text;

namespace SharePointLinkManifestBuilder.Core.Security;

/// <summary>
/// Escapes untrusted text for Markdown output, so a file name cannot restructure the document
/// or forge a link.
/// </summary>
public static class MarkdownEscaper
{
    private static readonly char[] SpecialCharacters =
        ['\\', '`', '*', '_', '{', '}', '[', ']', '(', ')', '#', '+', '-', '.', '!', '|', '<', '>'];

    /// <summary>Escapes Markdown syntax characters and flattens line breaks.</summary>
    public static string Escape(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length + 8);
        foreach (var c in value)
        {
            if (c is '\r' or '\n')
            {
                builder.Append(' ');
                continue;
            }

            if (Array.IndexOf(SpecialCharacters, c) >= 0)
            {
                builder.Append('\\');
            }

            builder.Append(c);
        }

        return builder.ToString();
    }

    /// <summary>
    /// Escapes a value for use inside a Markdown table cell.
    /// <para>
    /// The pipe character is already in <see cref="SpecialCharacters"/>, so this simply defers
    /// to <see cref="Escape"/>. Escaping the pipe a second time here would emit a literal
    /// backslash followed by an unescaped pipe, breaking the very table it aims to protect.
    /// </para>
    /// </summary>
    public static string EscapeTableCell(string? value) => Escape(value);

    /// <summary>
    /// Renders a Markdown link with an escaped label. URLs are angle-bracketed so parentheses
    /// and spaces inside a SharePoint URL cannot terminate the link target early.
    /// </summary>
    public static string Link(string? label, string? url)
    {
        var text = Escape(label);

        if (string.IsNullOrWhiteSpace(url))
        {
            return text;
        }

        return $"[{text}](<{url.Replace(">", "%3E", StringComparison.Ordinal)}>)";
    }
}
