namespace Avig.OpenXml;

/// <summary>
/// Part names and relationship targets. Part names here have no leading slash
/// (<c>word/document.xml</c>), as zip entries do; the package's own relationships belong to
/// the empty part name.
/// </summary>
public static class OpcPath
{
    /// <summary>
    /// True when a relationship type ends in <paramref name="kind"/>, for example "styles".
    /// Types are compared on their last segment, because Strict OOXML (ISO 29500 Strict) has
    /// other namespaces (purl.oclc.org) but the same name at the end.
    /// </summary>
    public static bool IsRelationship(string type, string kind)
    {
        return type.EndsWith("/" + kind, StringComparison.Ordinal);
    }

    /// <summary>The path of a part's .rels file, for example word/_rels/document.xml.rels; _rels/.rels for the package itself.</summary>
    public static string RelationshipsPath(string partPath)
    {
        if (partPath.Length == 0)
        {
            return "_rels/.rels";
        }

        int slash = partPath.LastIndexOf('/');

        if (slash < 0)
        {
            return "_rels/" + partPath + ".rels";
        }

        return partPath[..(slash + 1)] + "_rels/" + partPath[(slash + 1)..] + ".rels";
    }

    /// <summary>
    /// Resolves a relationship target against the part that points to it. A target starting
    /// with / is from the package root; otherwise relative to the part's folder, with .. allowed
    /// (but never above the root).
    /// </summary>
    public static string ResolveTarget(string sourcePartPath, string target)
    {
        string cleaned = target.Replace('\\', '/');

        if (cleaned.StartsWith('/') == true)
        {
            cleaned = cleaned[1..];
            sourcePartPath = string.Empty;
        }

        int slash = sourcePartPath.LastIndexOf('/');
        List<string> segments = [];

        if (slash >= 0)
        {
            segments.AddRange(sourcePartPath[..slash].Split('/', StringSplitOptions.RemoveEmptyEntries));
        }

        foreach (string segment in cleaned.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == "..")
            {
                if (segments.Count > 0)
                {
                    segments.RemoveAt(segments.Count - 1);
                }

                continue;
            }

            if (segment == ".")
            {
                continue;
            }

            segments.Add(segment);
        }

        return string.Join('/', segments);
    }

    /// <summary>
    /// The target to write in a relationship from <paramref name="sourcePartPath"/> to
    /// <paramref name="targetPartPath"/>: relative to the source's folder, as Office writes it.
    /// </summary>
    public static string RelativeTarget(string sourcePartPath, string targetPartPath)
    {
        string[] from = Folder(sourcePartPath);
        string[] to = targetPartPath.Split('/', StringSplitOptions.RemoveEmptyEntries);
        int common = 0;

        while (common < from.Length && common < to.Length - 1 && string.Equals(from[common], to[common], StringComparison.OrdinalIgnoreCase) == true)
        {
            common++;
        }

        List<string> segments = [];

        for (int index = common; index < from.Length; index++)
        {
            segments.Add("..");
        }

        for (int index = common; index < to.Length; index++)
        {
            segments.Add(to[index]);
        }

        return string.Join('/', segments);
    }

    /// <summary>
    /// A part name as zip entries hold it: / instead of \, no leading slash. Null for a name no
    /// package may have: empty, climbing out of the package with .., rooted at a drive, or with
    /// control characters.
    /// </summary>
    public static bool TryNormalize(string name, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? normalized)
    {
        normalized = null;
        string cleaned = name.Replace('\\', '/');

        if (cleaned.StartsWith('/') == true)
        {
            cleaned = cleaned[1..];
        }

        if (cleaned.Length == 0 || cleaned.Split('/').Contains("..") == true || cleaned.StartsWith('/') == true
            || (cleaned.Length >= 2 && cleaned[1] == ':') || cleaned.Any(char.IsControl) == true)
        {
            return false;
        }

        normalized = cleaned;

        return true;
    }

    private static string[] Folder(string partPath)
    {
        int slash = partPath.LastIndexOf('/');

        if (slash < 0)
        {
            return [];
        }

        return partPath[..slash].Split('/', StringSplitOptions.RemoveEmptyEntries);
    }
}
