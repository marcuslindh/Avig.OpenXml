using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Avig.OpenXml;

/// <summary>
/// Writes an OPC package, as a zip or as Flat OPC. Parts are written one at a time; the
/// relationships added with <see cref="AddRelationship(string, string, string)"/> become .rels parts, and
/// [Content_Types].xml is made from the content types the parts were written with, when the
/// package is completed. A zip part goes straight to the file, so a large part needs no more
/// memory than its writer does; Flat OPC is one XML file and is held in memory until
/// <see cref="CompleteAsync"/>.
/// </summary>
/// <example>
/// <code>
/// await using OpcPackageWriter package = await OpcPackageWriter.CreateAsync(output);
/// await package.WritePartAsync("word/document.xml", DocumentContentType, async stream => await WriteBodyAsync(stream));
/// package.AddRelationship(string.Empty, OpcNamespaces.OfficeDocumentRelationship, "word/document.xml");
/// await package.CompleteAsync();
/// </code>
/// </example>
public sealed class OpcPackageWriter : IAsyncDisposable
{
    private const string FlatPrefix = "pkg";

    private readonly Stream output;
    private readonly bool leaveOpen;
    private readonly ZipArchive? archive;
    private readonly List<(string Name, string ContentType, byte[]? Data)> written = [];
    private readonly HashSet<string> names = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, List<OpcRelationship>> relationships = new(StringComparer.OrdinalIgnoreCase);
    private bool completed;

    private OpcPackageWriter(OpcPackageKind kind, Stream output, bool leaveOpen, ZipArchive? archive)
    {
        Kind = kind;
        this.output = output;
        this.leaveOpen = leaveOpen;
        this.archive = archive;
    }

    /// <summary>Whether the package is written as a zip or as Flat OPC.</summary>
    public OpcPackageKind Kind { get; }

    /// <summary>
    /// The program a Flat OPC file says it is for, in its mso-application processing
    /// instruction, such as Word.Document; Windows opens the .xml file in that program. None by default.
    /// </summary>
    public string? FlatOpcProgId { get; set; }

    /// <summary>Starts a package on a stream.</summary>
    /// <param name="output">Where the file goes.</param>
    /// <param name="kind">A zip file (the default) or Flat OPC.</param>
    /// <param name="leaveOpen">Leave the stream open when the writer is disposed.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The writer.</returns>
    public static async Task<OpcPackageWriter> CreateAsync(Stream output, OpcPackageKind kind = OpcPackageKind.Zip, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(output);
        ZipArchive? archive = null;

        if (kind == OpcPackageKind.Zip)
        {
            archive = await ZipArchive.CreateAsync(output, ZipArchiveMode.Create, true, null, cancellationToken);
        }

        return new OpcPackageWriter(kind, output, leaveOpen, archive);
    }

    /// <summary>True when a part with this name has been written.</summary>
    public bool HasPart(string name)
    {
        return names.Contains(Clean(name));
    }

    /// <summary>
    /// Writes a part. <paramref name="write"/> gets a stream to write the part's bytes to and
    /// must be done with it when its task completes.
    /// </summary>
    /// <param name="name">The part name, such as word/document.xml.</param>
    /// <param name="contentType">Its content type, for [Content_Types].xml.</param>
    /// <param name="write">Writes the bytes.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>A task that completes when the part is written.</returns>
    /// <exception cref="InvalidOperationException">The package is completed, or the part is already written.</exception>
    public async Task WritePartAsync(string name, string contentType, Func<Stream, Task> write, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(write);
        ArgumentException.ThrowIfNullOrWhiteSpace(contentType);
        string part = Claim(name);

        if (archive is not null)
        {
            ZipArchiveEntry entry = archive.CreateEntry(part, CompressionLevel.Optimal);

            await using (Stream stream = await entry.OpenAsync(cancellationToken))
            {
                await write(stream);
            }

            written.Add((part, contentType, null));

            return;
        }

        using MemoryStream buffer = new();
        await write(buffer);
        written.Add((part, contentType, buffer.ToArray()));
    }

    /// <summary>Writes a part from bytes held in memory.</summary>
    /// <param name="name">The part name.</param>
    /// <param name="contentType">Its content type.</param>
    /// <param name="data">The bytes.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>A task that completes when the part is written.</returns>
    public Task WritePartAsync(string name, string contentType, ReadOnlyMemory<byte> data, CancellationToken cancellationToken = default)
    {
        return WritePartAsync(name, contentType, async stream => await stream.WriteAsync(data, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Copies a part from an opened package as it is, with its content type and (unless
    /// <paramref name="withRelationships"/> is false) its relationships.
    /// </summary>
    /// <returns>A task that completes when the part is copied.</returns>
    public async Task CopyPartAsync(OpcPackage source, string name, bool withRelationships = true, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(source);
        string part = Clean(name);
        if (source.TryGetContentType(part, out string? contentType) == false)
        {
            contentType = "application/octet-stream";
        }

        await WritePartAsync(part, contentType, async stream =>
        {
            Result<Stream> opened = await source.OpenPartAsync(part, cancellationToken);

            if (opened.IsSuccess(out Stream? from) == true)
            {
                await using (from)
                {
                    await from.CopyToAsync(stream, cancellationToken);
                }
            }
        }, cancellationToken);

        if (withRelationships == true)
        {
            foreach (OpcRelationship relationship in await source.GetRelationshipsAsync(part, cancellationToken))
            {
                AddRelationship(part, relationship);
            }
        }
    }

    /// <summary>
    /// Adds a relationship from <paramref name="sourcePart"/> (the empty name for the package
    /// itself) to a part, and gives back its id. The target is written relative to the source.
    /// </summary>
    /// <returns>The new relationship's id, rId1, rId2 … in the order they were added for the source.</returns>
    public string AddRelationship(string sourcePart, string type, string targetPart)
    {
        string source = sourcePart.TrimStart('/');
        string id = NextId(source);
        string target = Clean(targetPart);

        if (source.Length > 0)
        {
            target = OpcPath.RelativeTarget(source, target);
        }

        AddRelationship(source, new OpcRelationship(id, type, target, false));

        return id;
    }

    /// <summary>Adds a relationship to a URL outside the package (TargetMode="External"), such as a hyperlink.</summary>
    /// <returns>The new relationship's id.</returns>
    public string AddExternalRelationship(string sourcePart, string type, string url)
    {
        string source = sourcePart.TrimStart('/');
        string id = NextId(source);
        AddRelationship(source, new OpcRelationship(id, type, url, true));

        return id;
    }

    /// <summary>Adds a relationship as it is, with its own id and target, as when a part is copied.</summary>
    /// <exception cref="InvalidOperationException">The source already has a relationship with that id.</exception>
    public void AddRelationship(string sourcePart, OpcRelationship relationship)
    {
        ArgumentNullException.ThrowIfNull(relationship);
        ThrowIfCompleted();
        string source = sourcePart.TrimStart('/');

        if (relationships.TryGetValue(source, out List<OpcRelationship>? list) == false)
        {
            list = [];
            relationships[source] = list;
        }

        if (list.Any(existing => string.Equals(existing.Id, relationship.Id, StringComparison.Ordinal) == true) == true)
        {
            throw new InvalidOperationException($"The part '{source}' already has a relationship with the id '{relationship.Id}'.");
        }

        list.Add(relationship);
    }

    /// <summary>
    /// Writes the .rels parts and [Content_Types].xml and finishes the file. Nothing can be
    /// written afterwards. Disposing without completing leaves an unusable file.
    /// </summary>
    /// <returns>A task that completes when the file is written.</returns>
    public async Task CompleteAsync(CancellationToken cancellationToken = default)
    {
        ThrowIfCompleted();

        foreach ((string source, List<OpcRelationship> list) in relationships)
        {
            byte[] rels = RelationshipsXml(list);
            await WritePartAsync(OpcPath.RelationshipsPath(source), OpcNamespaces.RelationshipsContentType, rels, cancellationToken);
        }

        completed = true;

        if (archive is not null)
        {
            byte[] types = ContentTypesXml();
            ZipArchiveEntry entry = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);

            await using (Stream stream = await entry.OpenAsync(cancellationToken))
            {
                await stream.WriteAsync(types, cancellationToken);
            }

            await archive.DisposeAsync();

            return;
        }

        await WriteFlatAsync(cancellationToken);
    }

    /// <summary>Closes the file, completed or not.</summary>
    /// <returns>A task that completes when the file is closed.</returns>
    public async ValueTask DisposeAsync()
    {
        if (completed == false && archive is not null)
        {
            await archive.DisposeAsync();
        }

        completed = true;

        if (leaveOpen == false)
        {
            await output.DisposeAsync();
        }
    }

    private static string Clean(string name)
    {
        if (OpcPath.TryNormalize(name, out string? normalized) == false)
        {
            throw new ArgumentException($"'{name}' is not a part name a package may have.", nameof(name));
        }

        return normalized;
    }

    private static byte[] RelationshipsXml(List<OpcRelationship> list)
    {
        using MemoryStream buffer = new();

        using (XmlWriter writer = XmlWriter.Create(buffer, new XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
        {
            writer.WriteStartDocument(true);
            writer.WriteStartElement("Relationships", OpcNamespaces.PackageRelationships);

            foreach (OpcRelationship relationship in list)
            {
                writer.WriteStartElement("Relationship", OpcNamespaces.PackageRelationships);
                writer.WriteAttributeString("Id", relationship.Id);
                writer.WriteAttributeString("Type", relationship.Type);
                writer.WriteAttributeString("Target", relationship.Target);

                if (relationship.IsExternal == true)
                {
                    writer.WriteAttributeString("TargetMode", "External");
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        return buffer.ToArray();
    }

    private static bool IsXml(string contentType)
    {
        return contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) == true || contentType.EndsWith("/xml", StringComparison.OrdinalIgnoreCase) == true;
    }

    private string Claim(string name)
    {
        ThrowIfCompleted();
        string part = Clean(name);

        if (string.Equals(part, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase) == true)
        {
            throw new InvalidOperationException("[Content_Types].xml is written by CompleteAsync, from the content types the parts were written with.");
        }

        if (names.Add(part) == false)
        {
            throw new InvalidOperationException($"The part '{part}' is already written.");
        }

        return part;
    }

    private string NextId(string source)
    {
        int count = 1;

        if (relationships.TryGetValue(source, out List<OpcRelationship>? list) == true)
        {
            count = list.Count + 1;
        }

        string id = $"rId{count}";

        while (list is not null && list.Any(existing => string.Equals(existing.Id, id, StringComparison.Ordinal) == true) == true)
        {
            count++;
            id = $"rId{count}";
        }

        return id;
    }

    private void ThrowIfCompleted()
    {
        if (completed == true)
        {
            throw new InvalidOperationException("The package is completed; nothing more can be written to it.");
        }
    }

    /// <summary>
    /// A Default per extension where every part with it has the same content type (images,
    /// .rels), and an Override for every other part, as Office writes it.
    /// </summary>
    private byte[] ContentTypesXml()
    {
        Dictionary<string, string?> byExtension = new(StringComparer.OrdinalIgnoreCase)
        {
            ["rels"] = OpcNamespaces.RelationshipsContentType,
            ["xml"] = "application/xml",
        };

        foreach ((string name, string contentType, _) in written)
        {
            string extension = Extension(name);

            if (extension.Length == 0 || string.Equals(extension, "xml", StringComparison.OrdinalIgnoreCase) == true)
            {
                continue;
            }

            if (byExtension.TryGetValue(extension, out string? known) == false)
            {
                byExtension[extension] = contentType;
            }
            else if (string.Equals(known, contentType, StringComparison.OrdinalIgnoreCase) == false)
            {
                byExtension[extension] = null;
            }
        }

        using MemoryStream buffer = new();

        using (XmlWriter writer = XmlWriter.Create(buffer, new XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
        {
            writer.WriteStartDocument(true);
            writer.WriteStartElement("Types", OpcNamespaces.ContentTypes);

            foreach ((string extension, string? contentType) in byExtension)
            {
                if (contentType is not null)
                {
                    writer.WriteStartElement("Default", OpcNamespaces.ContentTypes);
                    writer.WriteAttributeString("Extension", extension);
                    writer.WriteAttributeString("ContentType", contentType);
                    writer.WriteEndElement();
                }
            }

            foreach ((string name, string contentType, _) in written)
            {
                if (byExtension.TryGetValue(Extension(name), out string? byDefault) == true && string.Equals(byDefault, contentType, StringComparison.OrdinalIgnoreCase) == true)
                {
                    continue;
                }

                writer.WriteStartElement("Override", OpcNamespaces.ContentTypes);
                writer.WriteAttributeString("PartName", "/" + name);
                writer.WriteAttributeString("ContentType", contentType);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        return buffer.ToArray();
    }

    private static string Extension(string name)
    {
        int dot = name.LastIndexOf('.');
        int slash = name.LastIndexOf('/');

        if (dot < 0 || dot < slash)
        {
            return string.Empty;
        }

        return name[(dot + 1)..];
    }

    /// <summary>
    /// The whole package as pkg:package. XML parts go in pkg:xmlData, copied node by node so
    /// that the part keeps its own namespace declarations; everything else in pkg:binaryData as
    /// base64, wrapped at 76 characters as Word writes it.
    /// </summary>
    private async Task WriteFlatAsync(CancellationToken cancellationToken)
    {
        XmlWriterSettings settings = new() { Async = true, Encoding = new UTF8Encoding(false), CloseOutput = false };

        await using XmlWriter writer = XmlWriter.Create(output, settings);
        await writer.WriteStartDocumentAsync(true);

        if (FlatOpcProgId is not null)
        {
            await writer.WriteProcessingInstructionAsync("mso-application", $"progid=\"{FlatOpcProgId}\"");
        }

        await writer.WriteStartElementAsync(FlatPrefix, "package", OpcNamespaces.FlatOpc);

        foreach ((string name, string contentType, byte[]? data) in written)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await writer.WriteStartElementAsync(FlatPrefix, "part", OpcNamespaces.FlatOpc);
            await writer.WriteAttributeStringAsync(FlatPrefix, "name", OpcNamespaces.FlatOpc, "/" + name);
            await writer.WriteAttributeStringAsync(FlatPrefix, "contentType", OpcNamespaces.FlatOpc, contentType);
            byte[] bytes = data ?? [];

            if (IsXml(contentType) == true && bytes.Length > 0)
            {
                await writer.WriteStartElementAsync(FlatPrefix, "xmlData", OpcNamespaces.FlatOpc);

                using (MemoryStream part = new(bytes, false))
                using (XmlReader reader = OpcXml.CreateReader(part))
                {
                    await reader.MoveToContentAsync();
                    await writer.WriteNodeAsync(reader, true);
                }

                await writer.WriteEndElementAsync();
            }
            else
            {
                await writer.WriteAttributeStringAsync(FlatPrefix, "compression", OpcNamespaces.FlatOpc, "store");
                await writer.WriteStartElementAsync(FlatPrefix, "binaryData", OpcNamespaces.FlatOpc);
                await writer.WriteStringAsync(Convert.ToBase64String(bytes, Base64FormattingOptions.InsertLineBreaks));
                await writer.WriteEndElementAsync();
            }

            await writer.WriteEndElementAsync();
        }

        await writer.WriteEndElementAsync();
        await writer.WriteEndDocumentAsync();
        await writer.FlushAsync();
    }
}
