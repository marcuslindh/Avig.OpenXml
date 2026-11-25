using System.Text;

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
    /// A part name as zip entries hold it: / instead of \, no leading slash. False for a name no
    /// package may have (see <see cref="IsSafe"/>).
    /// </summary>
    public static bool TryNormalize(string name, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? normalized)
    {
        normalized = null;
        string cleaned = name.Replace('\\', '/');

        if (cleaned.StartsWith('/') == true)
        {
            cleaned = cleaned[1..];
        }

        if (IsSafe(cleaned) == false)
        {
            return false;
        }

        normalized = cleaned;

        return true;
    }

    /// <summary>
    /// Whether a part name, with / for \ and no leading slash, is one a package may have. Not:
    /// empty; rooted (/, //server, C:); with a segment that is empty, . or .., or ends with a dot
    /// or a space (Windows drops those, so two names would be one file); with : (an NTFS stream)
    /// or control characters; with a percent-encoded dot, slash, backslash or NUL; with a
    /// Windows device name (CON, NUL, COM1 …) as a segment; or one of these once Unicode
    /// compatibility forms are folded (a fullwidth ．．／ is ../ to some unpackers). OPC forbids
    /// most of it (ECMA-376 part 2, 6.2.2), and no office program writes the rest; each is a way
    /// for a file to put something where it does not belong when it is unpacked or saved again.
    /// </summary>
    public static bool IsSafe(string name)
    {
        if (IsSafeAsWritten(name) == false)
        {
            return false;
        }

        string folded = name.Normalize(NormalizationForm.FormKC).Replace('\\', '/');

        return string.Equals(folded, name, StringComparison.Ordinal) == true || IsSafeAsWritten(folded) == true;
    }

    /// <summary>The key two part names share when they are the same file to a reader or an unpacker: case and Unicode compatibility forms folded.</summary>
    public static string ComparisonKey(string name)
    {
        return name.Replace('\\', '/').Normalize(NormalizationForm.FormKC).ToUpperInvariant();
    }

    private static readonly string[] DeviceNames = ["CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"];

    private static readonly string[] EncodedSeparators = ["%2e", "%2f", "%5c", "%00"];

    private static bool IsSafeAsWritten(string name)
    {
        if (name.Length == 0 || name.StartsWith('/') == true || name.Contains(':', StringComparison.Ordinal) == true || name.Any(char.IsControl) == true)
        {
            return false;
        }

        if (EncodedSeparators.Any(encoded => name.Contains(encoded, StringComparison.OrdinalIgnoreCase) == true) == true)
        {
            return false;
        }

        foreach (string segment in name.Split('/'))
        {
            if (segment.Length == 0 || segment.EndsWith('.') == true || segment.EndsWith(' ') == true)
            {
                return false;
            }

            string baseName = segment.Split('.')[0].TrimEnd();

            if (DeviceNames.Contains(baseName, StringComparer.OrdinalIgnoreCase) == true)
            {
                return false;
            }
        }

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
