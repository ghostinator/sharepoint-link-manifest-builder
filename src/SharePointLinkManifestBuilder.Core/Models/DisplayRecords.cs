namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>
/// One local storage location, shown on the privacy and setup pages.
/// <para>
/// A named tuple would read the same in C#, but tuple element names are compiler metadata
/// rather than runtime members, so a compiled XAML binding cannot see them. These small records
/// exist so the UI can bind to real, named properties.
/// </para>
/// </summary>
/// <param name="Description">What is stored there, in plain language.</param>
/// <param name="Path">The absolute path on this machine.</param>
public sealed record StorageLocationInfo(string Description, string Path);

/// <summary>A help entry: a topic and its explanation.</summary>
/// <param name="Topic">The heading.</param>
/// <param name="Explanation">The body text.</param>
public sealed record HelpTopic(string Topic, string Explanation);

/// <summary>A third-party dependency listed on the About page.</summary>
/// <param name="Component">The component name.</param>
/// <param name="Licence">Its licence.</param>
/// <param name="Purpose">Why this application uses it.</param>
public sealed record ThirdPartyComponent(string Component, string Licence, string Purpose);
