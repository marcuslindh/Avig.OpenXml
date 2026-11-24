using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Xml;

namespace Avig.OpenXml;

/// <summary>
/// Checks and reads on a package opened as a <see cref="ZipArchive"/>. Every part that is read
/// goes through <see cref="OpenEntryAsync"/>, so that no part can unpack to more than its zip
/// header says, and the directory is checked with <see cref="CheckSize"/> and
/// <see cref="CheckEntriesAsync"/> before anything is read.
/// </summary>
public static class OpcZip
{
    /// <summary>
    /// The entry for a part name. Part names in OPC are case-insensitive, and some programs
    /// write \ instead of /.
    /// </summary>
    public static bool TryFindEntry(ZipArchive archive, string path, [NotNullWhen(true)] out ZipArchiveEntry? entry)
    {
        entry = archive.GetEntry(path);

        if (entry is not null)
        {
            return true;
        }

        foreach (ZipArchiveEntry candidate in archive.Entries)
        {
            if (string.Equals(candidate.FullName.Replace('\\', '/'), path, StringComparison.OrdinalIgnoreCase) == true)
            {
                entry = candidate;

                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Invalid for a zip directory no office program writes: two entries with the same name
    /// (case ignored, \ read as /), a name that climbs out of the package (..), starts at a root
    /// or a drive, or holds control characters, and an entry that cannot be unpacked (a
    /// compression method .NET does not have, or encryption). Duplicates let two readers see two
    /// different files; such names would be passed on when the file is saved, to whoever unpacks
    /// it next; and an entry that cannot be opened would otherwise fail only when it is saved.
    /// </summary>
    public static async Task<Result> CheckEntriesAsync(ZipArchive archive, CancellationToken cancellationToken)
    {
        HashSet<string> names = new(StringComparer.OrdinalIgnoreCase);

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            string name = entry.FullName.Replace('\\', '/');

            if (names.Add(name) == false)
            {
                return Result.Invalid($"{OpcErrors.DamagedPrefix}: the part '{entry.FullName}' is in it twice.");
            }

            // An empty name opened fine and then failed the save, where the entry is written again.
            if (name.Length == 0)
            {
                return Result.Invalid($"{OpcErrors.DamagedPrefix}: it has a part without a name.");
            }

            bool climbs = name.Split('/').Contains("..");
            bool rooted = name.StartsWith('/') == true || (name.Length >= 2 && name[1] == ':');

            if (climbs == true || rooted == true || name.Any(char.IsControl) == true)
            {
                return Result.Invalid($"{OpcErrors.DamagedPrefix}: the part name '{entry.FullName}' points outside the package.");
            }

            try
            {
                await using Stream stream = await entry.OpenAsync(cancellationToken);
            }
            catch (Exception exception) when (exception is NotSupportedException or InvalidDataException)
            {
                return Result.Invalid($"{OpcErrors.DamagedPrefix}: the part '{entry.FullName}' cannot be unpacked ({exception.Message}).");
            }
        }

        return Result.Success();
    }

    /// <summary>
    /// Opens a part of a file being read, held to the size its zip header gives (see
    /// <see cref="LimitedStream"/>).
    /// </summary>
    public static async Task<Stream> OpenEntryAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        Stream stream = await entry.OpenAsync(cancellationToken);

        return new LimitedStream(stream, entry.Length, entry.FullName);
    }

    /// <summary>
    /// A stream that refuses to give more than <paramref name="limit"/> bytes, with an
    /// InvalidDataException that says the part named <paramref name="name"/> is damaged. For a
    /// part whose size is known from its header.
    /// </summary>
    public static Stream LimitStream(Stream inner, long limit, string name)
    {
        return new LimitedStream(inner, limit, name);
    }

    /// <summary>
    /// Invalid when the parts of the file add up to more than <paramref name="maxUncompressedBytes"/>
    /// unpacked. Checked from the zip directory before anything is read.
    /// <paramref name="limitName"/> names the setting in the message, for example
    /// "KalkylLimits.MaxUncompressedBytes", so that the caller's error kinds can recognize it.
    /// </summary>
    public static Result CheckSize(ZipArchive archive, long maxUncompressedBytes, string limitName)
    {
        long total = 0;

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            total += entry.Length;

            if (total > maxUncompressedBytes || total < 0)
            {
                return Result.Invalid(OpcErrors.TooLarge(maxUncompressedBytes, limitName));
            }
        }

        return Result.Success();
    }

    /// <summary>
    /// The internal relationships in a .rels file: id to (type, target). External ones
    /// (hyperlinks) are left out; <see cref="ReadAllRelationshipsAsync"/> has them too.
    /// </summary>
    public static async Task<IReadOnlyDictionary<string, (string Type, string Target)>> ReadRelationshipsAsync(ZipArchive archive, string relsPath, CancellationToken cancellationToken)
    {
        Dictionary<string, (string Type, string Target)> relationships = new(StringComparer.Ordinal);

        foreach (OpcRelationship relationship in await ReadAllRelationshipsAsync(archive, relsPath, cancellationToken))
        {
            if (relationship.IsExternal == false)
            {
                relationships[relationship.Id] = (relationship.Type, relationship.Target);
            }
        }

        return relationships;
    }

    /// <summary>Every relationship in a .rels file, external ones included, in file order. Empty when the file is not there.</summary>
    public static async Task<IReadOnlyList<OpcRelationship>> ReadAllRelationshipsAsync(ZipArchive archive, string relsPath, CancellationToken cancellationToken)
    {
        if (TryFindEntry(archive, relsPath, out ZipArchiveEntry? entry) == false)
        {
            return [];
        }

        await using Stream stream = await OpenEntryAsync(entry, cancellationToken);

        return await OpcRelationship.ReadAsync(stream, cancellationToken);
    }

    internal static async Task<IReadOnlyList<OpcRelationship>> ReadRelationshipsFromAsync(XmlReader reader)
    {
        List<OpcRelationship> relationships = [];

        while (await reader.ReadAsync() == true)
        {
            if (OpcXml.IsElement(reader, "Relationship") == true)
            {
                string? id = reader.GetAttribute("Id");
                string? type = reader.GetAttribute("Type");
                string? target = reader.GetAttribute("Target");
                bool external = string.Equals(reader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);

                if (id is not null && type is not null && target is not null)
                {
                    relationships.Add(new OpcRelationship(id, type, target, external));
                }
            }
        }

        return relationships;
    }
}
