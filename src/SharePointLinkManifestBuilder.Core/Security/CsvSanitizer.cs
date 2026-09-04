using System.Globalization;
using System.Text;

namespace SharePointLinkManifestBuilder.Core.Security;

/// <summary>
/// Writes CSV that is safe to open in a spreadsheet.
/// <para>
/// File names come from SharePoint and are attacker-influenced. A file literally named
/// <c>=cmd|'/c calc'!A1.docx</c> would, in a naive CSV, become a live formula the moment a
/// user opens the export. Every field is therefore neutralized before quoting.
/// </para>
/// </summary>
public static class CsvSanitizer
{
    /// <summary>
    /// Characters that make a spreadsheet treat a cell as a formula. Tab and carriage return
    /// are included because they can be used to shift a payload past naive checks.
    /// </summary>
    public static readonly char[] FormulaTriggers = ['=', '+', '-', '@', '\t', '\r'];

    /// <summary>
    /// Neutralizes a single field: prefixes a formula trigger with an apostrophe, then quotes
    /// and escapes per RFC 4180.
    /// </summary>
    /// <param name="value">The raw field value.</param>
    public static string SanitizeField(string? value)
    {
        var text = value ?? string.Empty;

        if (text.Length > 0 && Array.IndexOf(FormulaTriggers, text[0]) >= 0)
        {
            text = "'" + text;
        }

        var needsQuoting = text.Any(c => c is ',' or '"' or '\n' or '\r');
        if (!needsQuoting)
        {
            return text;
        }

        return "\"" + text.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
    }

    /// <summary>Builds one CSV row from field values, sanitizing each.</summary>
    public static string BuildRow(params string?[] fields) =>
        string.Join(',', fields.Select(SanitizeField));

    /// <summary>Builds one CSV row from a sequence of field values.</summary>
    public static string BuildRow(IEnumerable<string?> fields) =>
        string.Join(',', fields.Select(SanitizeField));

    /// <summary>Formats a timestamp for CSV using a round-trippable invariant representation.</summary>
    public static string FormatTimestamp(DateTimeOffset value) =>
        value.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);

    /// <summary>Writes a complete CSV document with a header row and CRLF line endings.</summary>
    public static string BuildDocument(IEnumerable<string> headers, IEnumerable<IEnumerable<string?>> rows)
    {
        var builder = new StringBuilder();
        builder.Append(BuildRow(headers.Cast<string?>())).Append("\r\n");

        foreach (var row in rows)
        {
            builder.Append(BuildRow(row)).Append("\r\n");
        }

        return builder.ToString();
    }
}
