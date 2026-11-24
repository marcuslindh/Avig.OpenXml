using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Avig.OpenXml;

/// <summary>Which of the two forms of an OPC package a file has.</summary>
public enum OpcPackageKind
{
    /// <summary>A zip file, as .docx, .xlsx and .pptx files are.</summary>
    Zip,

    /// <summary>
    /// Flat OPC: the whole package as one XML file (pkg:package), as Word's "Word XML Document"
    /// and Office's clipboard write it.
    /// </summary>
    FlatOpc,
}

/// <summary>
/// An OPC package opened for reading: its parts, their content types and the relationships
/// between them, whether the file is a zip or Flat OPC. Opening checks the whole directory
/// against the limits before anything is read, and every part is read through a stream that
/// holds it to its size.
/// </summary>
/// <example>
/// <code>
/// Result&lt;OpcPackage&gt; opened = await OpcPackage.OpenAsync("report.docx");
///
/// if (opened.IsSuccess(out OpcPackage? package) == true)
/// {
///     await using (package)
///     {
///         string? main = await package.FindMainPartAsync();
///     }
/// }
/// </code>
/// </example>
public sealed class OpcPackage : IAsyncDisposable
{
    private static readonly byte[] ZipSignature = [0x50, 0x4B, 0x03, 0x04];
    private static readonly byte[] EmptyZipSignature = [0x50, 0x4B, 0x05, 0x06];
    private static readonly byte[] CompoundFileSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    private readonly Stream? stream;
    private readonly bool leaveOpen;
    private readonly ZipArchive? archive;
    private readonly Dictionary<string, Part> parts;
    private readonly Dictionary<string, IReadOnlyList<OpcRelationship>> relationshipCache = new(StringComparer.OrdinalIgnoreCase);

    private OpcPackage(OpcPackageKind kind, OpcOpenOptions options, Stream? stream, bool leaveOpen, ZipArchive? archive, Dictionary<string, Part> parts)
    {
        Kind = kind;
        Options = options;
        this.stream = stream;
        this.leaveOpen = leaveOpen;
        this.archive = archive;
        this.parts = parts;
        PartNames = [.. parts.Values.Select(part => part.Name)];
    }

    /// <summary>Whether the file was a zip or Flat OPC.</summary>
    public OpcPackageKind Kind { get; }

    /// <summary>The options the package was opened with.</summary>
    public OpcOpenOptions Options { get; }

    /// <summary>Every part's name, without a leading slash, in file order. The .rels parts and [Content_Types].xml are among them.</summary>
    public IReadOnlyList<string> PartNames { get; }

    /// <summary>
    /// True when the package holds a VBA project, which only macro-enabled files (.docm,
    /// .xlsm, .dotm) have.
    /// </summary>
    public bool HasMacros => parts.Values.Any(part => string.Equals(part.ContentType, OpcNamespaces.VbaProjectContentType, StringComparison.OrdinalIgnoreCase) == true
        || part.Name.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase) == true);

    /// <summary>Opens a file.</summary>
    /// <param name="path">The file.</param>
    /// <param name="options">Limits and message wording; <see cref="OpcOpenOptions.Default"/> when null.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>
    /// The package; NotFound when the file does not exist, Forbidden when it may not be read,
    /// Conflict when another program holds it locked, and Invalid when it is not a package, is
    /// damaged, or goes past a limit.
    /// </returns>
    public static async Task<Result<OpcPackage>> OpenAsync(string path, OpcOpenOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        FileStream file;

        try
        {
            file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 65_536, FileOptions.Asynchronous);
        }
        catch (FileNotFoundException)
        {
            return Result<OpcPackage>.NotFound($"'{path}' does not exist.");
        }
        catch (DirectoryNotFoundException)
        {
            return Result<OpcPackage>.NotFound($"The folder for '{path}' does not exist.");
        }
        catch (UnauthorizedAccessException)
        {
            return Result<OpcPackage>.Forbidden($"'{path}' may not be read.");
        }
        catch (IOException exception)
        {
            return Result<OpcPackage>.Conflict($"'{path}' could not be opened: {exception.Message}");
        }

        Result<OpcPackage> opened = await OpenAsync(file, options, false, cancellationToken);

        if (opened.IsFailure(out _, out _) == true)
        {
            await file.DisposeAsync();
        }

        return opened;
    }

    /// <summary>
    /// Opens a package from a stream. A stream that cannot seek is copied to memory first,
    /// because the zip directory sits at the end of the file.
    /// </summary>
    /// <param name="input">The stream.</param>
    /// <param name="options">Limits and message wording; <see cref="OpcOpenOptions.Default"/> when null.</param>
    /// <param name="leaveOpen">Leave the stream open when the package is disposed.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The package, or Invalid when the content is not a package, is damaged, or goes past a limit.</returns>
    public static async Task<Result<OpcPackage>> OpenAsync(Stream input, OpcOpenOptions? options = null, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        options ??= OpcOpenOptions.Default;
        Stream source = input;
        bool ownsSource = false;

        if (input.CanSeek == false)
        {
            MemoryStream copy = new();

            if (await CopyLimitedAsync(input, copy, options.MaxUncompressedBytes, cancellationToken) == false)
            {
                await copy.DisposeAsync();

                return Result<OpcPackage>.Invalid($"The stream is longer than {options.MaxUncompressedBytes:N0} bytes, {OpcErrors.LimitMarker}{options.LimitsName}.MaxUncompressedBytes.");
            }

            copy.Position = 0;
            source = copy;
            ownsSource = true;

            if (leaveOpen == false)
            {
                await input.DisposeAsync();
            }
        }

        Result<OpcPackage> opened;

        try
        {
            opened = await OpenSeekableAsync(source, options, ownsSource == false && leaveOpen == true, cancellationToken);
        }
        catch (XmlException exception)
        {
            opened = Result<OpcPackage>.Invalid($"{OpcErrors.DamagedPrefix}: a part has broken XML ({exception.Message}).");
        }
        catch (InvalidDataException exception)
        {
            opened = Result<OpcPackage>.Invalid(Describe(exception));
        }

        if (opened.IsFailure(out _, out _) == true && ownsSource == true)
        {
            await source.DisposeAsync();
        }

        return opened;
    }

    /// <summary>The message for an <see cref="InvalidDataException"/> met while reading: this package's own as it is, one from .NET marked as a damaged file.</summary>
    public static string Describe(InvalidDataException exception)
    {
        ArgumentNullException.ThrowIfNull(exception);

        if (exception.Message.StartsWith(OpcErrors.DamagedPrefix, StringComparison.Ordinal) == true
            || exception.Message.Contains(OpcErrors.LimitMarker, StringComparison.Ordinal) == true)
        {
            return exception.Message;
        }

        return $"{OpcErrors.DamagedPrefix}: {exception.Message}";
    }

    /// <summary>True when the package has a part with this name (case ignored, leading slash allowed).</summary>
    public bool Contains(string partName)
    {
        return TryGetPart(partName, out _);
    }

    /// <summary>The part's content type, from [Content_Types].xml or Flat OPC's pkg:contentType; false when there is no such part or the package does not say.</summary>
    public bool TryGetContentType(string partName, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? contentType)
    {
        contentType = null;

        if (TryGetPart(partName, out Part? part) == true)
        {
            contentType = part.ContentType;
        }

        return contentType is not null;
    }

    /// <summary>The part's size in bytes unpacked, as its zip header gives it.</summary>
    public long GetLength(string partName)
    {
        if (TryGetPart(partName, out Part? part) == true)
        {
            return part.Length;
        }

        return 0;
    }

    /// <summary>
    /// Opens a part for reading; NotFound when the package has no such part. A zip part's
    /// stream refuses to give more bytes than its header said, with an InvalidDataException.
    /// </summary>
    public async Task<Result<Stream>> OpenPartAsync(string partName, CancellationToken cancellationToken = default)
    {
        if (TryGetPart(partName, out Part? part) == false)
        {
            return Result<Stream>.NotFound($"The package has no part '{partName}'.");
        }

        if (part.Data is not null)
        {
            return new MemoryStream(part.Data, false);
        }

        return await OpcZip.OpenEntryAsync(part.Entry!, cancellationToken);
    }

    /// <summary>A part's bytes; NotFound when the package has no such part.</summary>
    public async Task<Result<byte[]>> ReadPartAsync(string partName, CancellationToken cancellationToken = default)
    {
        Result<Stream> opened = await OpenPartAsync(partName, cancellationToken);

        if (opened.IsFailure(out string? error, out ResultStatus status, out Stream? part) == true)
        {
            return Result<byte[]>.Failure(error, status);
        }

        await using (part)
        {
            using MemoryStream copy = new();
            await part.CopyToAsync(copy, cancellationToken);

            return copy.ToArray();
        }
    }

    /// <summary>A reader for an XML part, with the package's depth limit and no DTD (see <see cref="OpcXml"/>).</summary>
    public XmlReader CreateXmlReader(Stream part)
    {
        return OpcXml.CreateReader(part, Options.MaxXmlDepth);
    }

    /// <summary>
    /// The relationships of a part, or of the package itself for the empty name, in file order.
    /// Empty when it has none. Read once and kept.
    /// </summary>
    public async Task<IReadOnlyList<OpcRelationship>> GetRelationshipsAsync(string partName, CancellationToken cancellationToken = default)
    {
        string name = partName.TrimStart('/');

        if (relationshipCache.TryGetValue(name, out IReadOnlyList<OpcRelationship>? cached) == true)
        {
            return cached;
        }

        IReadOnlyList<OpcRelationship> relationships = [];
        Result<Stream> opened = await OpenPartAsync(OpcPath.RelationshipsPath(name), cancellationToken);

        if (opened.IsSuccess(out Stream? rels) == true)
        {
            await using (rels)
            {
                using XmlReader reader = CreateXmlReader(rels);
                relationships = await OpcZip.ReadRelationshipsFromAsync(reader);
            }
        }

        relationshipCache[name] = relationships;

        return relationships;
    }

    /// <summary>
    /// The part the first relationship of a type points to from <paramref name="sourcePart"/>,
    /// resolved to a part name; NotFound when there is none (external ones do not count).
    /// </summary>
    public async Task<Result<string>> FindRelatedPartAsync(string sourcePart, string relationshipKind, CancellationToken cancellationToken = default)
    {
        foreach (OpcRelationship relationship in await GetRelationshipsAsync(sourcePart, cancellationToken))
        {
            if (relationship.IsExternal == false && OpcPath.IsRelationship(relationship.Type, relationshipKind) == true)
            {
                return OpcPath.ResolveTarget(sourcePart.TrimStart('/'), relationship.Target);
            }
        }

        return Result<string>.NotFound($"The part '{sourcePart}' has no {relationshipKind} relationship.");
    }

    /// <summary>The main part (word/document.xml, xl/workbook.xml …) that the package's officeDocument relationship points to; NotFound when there is none.</summary>
    public Task<Result<string>> FindMainPartAsync(CancellationToken cancellationToken = default)
    {
        return FindRelatedPartAsync(string.Empty, "officeDocument", cancellationToken);
    }

    /// <summary>Closes the file.</summary>
    /// <returns>A task that completes when the file is closed.</returns>
    public async ValueTask DisposeAsync()
    {
        if (archive is not null)
        {
            await archive.DisposeAsync();
        }

        if (stream is not null && leaveOpen == false)
        {
            await stream.DisposeAsync();
        }
    }

    private static async Task<bool> CopyLimitedAsync(Stream source, Stream target, long limit, CancellationToken cancellationToken)
    {
        byte[] buffer = new byte[81_920];
        long total = 0;
        int read = await source.ReadAsync(buffer, cancellationToken);

        while (read > 0)
        {
            total += read;

            if (total > limit)
            {
                return false;
            }

            await target.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
            read = await source.ReadAsync(buffer, cancellationToken);
        }

        return true;
    }

    private static async Task<Result<OpcPackage>> OpenSeekableAsync(Stream source, OpcOpenOptions options, bool leaveOpen, CancellationToken cancellationToken)
    {
        byte[] header = new byte[8];
        long start = source.Position;
        int read = await source.ReadAtLeastAsync(header, header.Length, false, cancellationToken);
        source.Position = start;

        if (read >= 8 && header.AsSpan().SequenceEqual(CompoundFileSignature) == true)
        {
            return Result<OpcPackage>.Invalid($"{options.NotFormatPrefix}. {options.CompoundFileHint}");
        }

        bool zip = read >= 4 && (header.AsSpan(0, 4).SequenceEqual(ZipSignature) == true || header.AsSpan(0, 4).SequenceEqual(EmptyZipSignature) == true);

        if (zip == true)
        {
            return await OpenZipAsync(source, options, leaveOpen, cancellationToken);
        }

        if (options.AllowFlatOpc == true && LooksLikeXml(header.AsSpan(0, read)) == true)
        {
            return await OpenFlatAsync(source, options, leaveOpen, cancellationToken);
        }

        return Result<OpcPackage>.Invalid($"{options.NotFormatPrefix}.");
    }

    private static bool LooksLikeXml(ReadOnlySpan<byte> header)
    {
        ReadOnlySpan<byte> rest = header;

        // A UTF-8 byte order mark, or the UTF-16 ones Word may write.
        if (rest.StartsWith((ReadOnlySpan<byte>)[0xEF, 0xBB, 0xBF]) == true || rest.StartsWith((ReadOnlySpan<byte>)[0xFF, 0xFE]) == true || rest.StartsWith((ReadOnlySpan<byte>)[0xFE, 0xFF]) == true)
        {
            return true;
        }

        foreach (byte value in rest)
        {
            if (value is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n')
            {
                continue;
            }

            return value == (byte)'<';
        }

        return false;
    }

    private static async Task<Result<OpcPackage>> OpenZipAsync(Stream source, OpcOpenOptions options, bool leaveOpen, CancellationToken cancellationToken)
    {
        ZipArchive archive;

        try
        {
            archive = await ZipArchive.CreateAsync(source, ZipArchiveMode.Read, true, null, cancellationToken);
        }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException)
        {
            return Result<OpcPackage>.Invalid($"{OpcErrors.DamagedPrefix}: the zip container is damaged ({exception.Message}).");
        }

        Result checks = await CheckZipAsync(archive, options, cancellationToken);

        if (checks.IsFailure(out string? error, out ResultStatus status) == true)
        {
            await archive.DisposeAsync();

            return Result<OpcPackage>.Failure(error, status);
        }

        Dictionary<string, Part> parts = new(StringComparer.OrdinalIgnoreCase);

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            // Folders in the zip directory are not parts.
            if (entry.FullName.EndsWith('/') == true && entry.Length == 0)
            {
                continue;
            }

            string name = entry.FullName.Replace('\\', '/');
            parts[name] = new Part(name, null, entry.Length, entry, null);
        }

        try
        {
            await ApplyContentTypesAsync(archive, parts, options, cancellationToken);
        }
        catch
        {
            await archive.DisposeAsync();

            throw;
        }

        return new OpcPackage(OpcPackageKind.Zip, options, source, leaveOpen, archive, parts);
    }

    private static async Task<Result> CheckZipAsync(ZipArchive archive, OpcOpenOptions options, CancellationToken cancellationToken)
    {
        if (archive.Entries.Count > options.MaxParts)
        {
            return Result.Invalid($"The file has {archive.Entries.Count:N0} parts, more than {options.MaxParts:N0}, {OpcErrors.LimitMarker}{options.LimitsName}.MaxParts.");
        }

        Result size = OpcZip.CheckSize(archive, options.MaxUncompressedBytes, options.LimitsName + ".MaxUncompressedBytes");

        if (size.IsFailure(out _, out _) == true)
        {
            return size;
        }

        return await OpcZip.CheckEntriesAsync(archive, cancellationToken);
    }

    private static async Task ApplyContentTypesAsync(ZipArchive archive, Dictionary<string, Part> parts, OpcOpenOptions options, CancellationToken cancellationToken)
    {
        if (OpcZip.TryFindEntry(archive, "[Content_Types].xml", out ZipArchiveEntry? entry) == false)
        {
            return;
        }

        Dictionary<string, string> defaults = new(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, string> overrides = new(StringComparer.OrdinalIgnoreCase);

        await using (Stream stream = await OpcZip.OpenEntryAsync(entry, cancellationToken))
        {
            using XmlReader reader = OpcXml.CreateReader(stream, options.MaxXmlDepth);

            while (await reader.ReadAsync() == true)
            {
                if (OpcXml.IsElement(reader, "Default") == true)
                {
                    string? extension = reader.GetAttribute("Extension");
                    string? type = reader.GetAttribute("ContentType");

                    if (extension is not null && type is not null)
                    {
                        defaults[extension] = type;
                    }
                }
                else if (OpcXml.IsElement(reader, "Override") == true)
                {
                    string? name = reader.GetAttribute("PartName");
                    string? type = reader.GetAttribute("ContentType");

                    if (name is not null && type is not null)
                    {
                        overrides[name.TrimStart('/')] = type;
                    }
                }
            }
        }

        foreach (Part part in parts.Values.ToList())
        {
            string? type = null;

            if (overrides.TryGetValue(part.Name, out string? overridden) == true)
            {
                type = overridden;
            }
            else
            {
                int dot = part.Name.LastIndexOf('.');

                if (dot >= 0 && defaults.TryGetValue(part.Name[(dot + 1)..], out string? byExtension) == true)
                {
                    type = byExtension;
                }
            }

            parts[part.Name] = part with { ContentType = type };
        }
    }

    private static async Task<Result<OpcPackage>> OpenFlatAsync(Stream source, OpcOpenOptions options, bool leaveOpen, CancellationToken cancellationToken)
    {
        if (source.Length - source.Position > options.MaxUncompressedBytes)
        {
            return Result<OpcPackage>.Invalid($"The file is larger than {options.MaxUncompressedBytes:N0} bytes, {OpcErrors.LimitMarker}{options.LimitsName}.MaxUncompressedBytes.");
        }

        Dictionary<string, Part> parts = new(StringComparer.OrdinalIgnoreCase);
        bool sawPackage = false;

        // The package wraps every part in pkg:package/pkg:part/pkg:xmlData, three levels more than the part itself has.
        using (XmlReader reader = OpcXml.CreateReader(source, options.MaxXmlDepth + 3))
        {
            while (await reader.ReadAsync() == true)
            {
                if (reader.NodeType != XmlNodeType.Element)
                {
                    continue;
                }

                if (sawPackage == false)
                {
                    if (string.Equals(reader.LocalName, "package", StringComparison.Ordinal) == false || string.Equals(reader.NamespaceURI, OpcNamespaces.FlatOpc, StringComparison.Ordinal) == false)
                    {
                        return Result<OpcPackage>.Invalid($"{options.NotFormatPrefix}: it is an XML file, but not a Flat OPC package.");
                    }

                    sawPackage = true;

                    continue;
                }

                if (OpcXml.IsElement(reader, "part") == false)
                {
                    continue;
                }

                Result<Part> part = await ReadFlatPartAsync(reader, cancellationToken);

                if (part.IsFailure(out string? error, out ResultStatus status, out Part? value) == true)
                {
                    return Result<OpcPackage>.Failure(error, status);
                }

                if (parts.TryAdd(value.Name, value) == false)
                {
                    return Result<OpcPackage>.Invalid($"{OpcErrors.DamagedPrefix}: the part '{value.Name}' is in it twice.");
                }

                if (parts.Count > options.MaxParts)
                {
                    return Result<OpcPackage>.Invalid($"The file has more than {options.MaxParts:N0} parts, {OpcErrors.LimitMarker}{options.LimitsName}.MaxParts.");
                }
            }
        }

        if (sawPackage == false)
        {
            return Result<OpcPackage>.Invalid($"{options.NotFormatPrefix}.");
        }

        // Everything is in memory now; the stream is not needed again.
        if (leaveOpen == false)
        {
            await source.DisposeAsync();
        }

        return new OpcPackage(OpcPackageKind.FlatOpc, options, null, true, null, parts);
    }

    private static async Task<Result<Part>> ReadFlatPartAsync(XmlReader reader, CancellationToken cancellationToken)
    {
        OpcXml.TryGetAttribute(reader, "contentType", out string? contentType);

        if (OpcXml.TryGetAttribute(reader, "name", out string? rawName) == false)
        {
            return Result<Part>.Invalid($"{OpcErrors.DamagedPrefix}: a pkg:part has no pkg:name.");
        }

        if (OpcPath.TryNormalize(rawName, out string? name) == false)
        {
            return Result<Part>.Invalid($"{OpcErrors.DamagedPrefix}: the part name '{rawName}' points outside the package.");
        }

        byte[] data = [];

        if (reader.IsEmptyElement == false)
        {
            int depth = reader.Depth;
            await reader.ReadAsync();

            while (OpcXml.IsEnd(reader, depth) == false && reader.EOF == false)
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (OpcXml.IsElement(reader, "xmlData") == true)
                {
                    data = await ReadXmlDataAsync(reader);

                    continue;
                }

                if (OpcXml.IsElement(reader, "binaryData") == true)
                {
                    string base64 = await reader.ReadElementContentAsStringAsync();

                    try
                    {
                        data = Convert.FromBase64String(base64);
                    }
                    catch (FormatException)
                    {
                        return Result<Part>.Invalid($"{OpcErrors.DamagedPrefix}: the part '{name}' has binary data that is not base64.");
                    }

                    continue;
                }

                await reader.ReadAsync();
            }
        }

        return new Part(name, contentType, data.Length, null, data);
    }

    /// <summary>
    /// The XML inside pkg:xmlData, as a part of its own. XmlWriter declares every namespace the
    /// copied elements use, so a part that relied on a declaration on pkg:package still parses.
    /// </summary>
    private static async Task<byte[]> ReadXmlDataAsync(XmlReader reader)
    {
        int depth = reader.Depth;

        if (reader.IsEmptyElement == true)
        {
            await reader.ReadAsync();

            return [];
        }

        await reader.ReadAsync();
        using MemoryStream buffer = new();
        XmlWriterSettings settings = new() { Async = true, Encoding = new UTF8Encoding(false), CloseOutput = false };

        await using (XmlWriter writer = XmlWriter.Create(buffer, settings))
        {
            await writer.WriteStartDocumentAsync(true);

            while (OpcXml.IsEnd(reader, depth) == false && reader.EOF == false)
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    await writer.WriteNodeAsync(reader, true);

                    continue;
                }

                await reader.ReadAsync();
            }

            await writer.WriteEndDocumentAsync();
        }

        // The end of pkg:xmlData.
        await reader.ReadAsync();

        return buffer.ToArray();
    }

    private bool TryGetPart(string partName, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Part? part)
    {
        return parts.TryGetValue(partName.Replace('\\', '/').TrimStart('/'), out part);
    }

    private sealed record Part(string Name, string? ContentType, long Length, ZipArchiveEntry? Entry, byte[]? Data);
}
