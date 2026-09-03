using System.Text;

namespace SharePointLinkManifestBuilder.Core.Urls;

/// <summary>
/// Encodes a sharing URL into the token Microsoft Graph's <c>/shares/{token}</c> endpoint
/// expects: <c>u!</c> followed by an unpadded base64url encoding of the UTF-8 URL.
/// </summary>
public static class GraphShareTokenEncoder
{
    /// <summary>Encodes a sharing URL into a Graph share token.</summary>
    /// <param name="sharingUrl">The absolute sharing URL.</param>
    /// <exception cref="ArgumentException">The URL is null or whitespace.</exception>
    public static string Encode(string sharingUrl)
    {
        if (string.IsNullOrWhiteSpace(sharingUrl))
        {
            throw new ArgumentException("A sharing URL is required.", nameof(sharingUrl));
        }

        var base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(sharingUrl.Trim()));

        // base64url: strip padding, then substitute the two URL-unsafe characters.
        var token = base64
            .TrimEnd('=')
            .Replace('/', '_')
            .Replace('+', '-');

        return "u!" + token;
    }
}
