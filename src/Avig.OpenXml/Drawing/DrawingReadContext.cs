namespace Avig.OpenXml.Drawing;

/// <summary>
/// What <see cref="DrawingReader"/> needs from the part the drawing is in: its relationships, to
/// turn an r:id into a link or an image, and a way to read an image part's bytes. The package
/// that owns the part makes one; nothing outside the package is ever followed.
/// </summary>
public sealed class DrawingReadContext
{
    private readonly Dictionary<string, OpcRelationship> relationships = new(StringComparer.Ordinal);
    private readonly Func<OpcRelationship, CancellationToken, Task<Result<byte[]>>>? readImage;
    private readonly Func<OpcRelationship, Result<byte[]>>? readImageNow;

    /// <summary>Creates a context.</summary>
    /// <param name="relationships">The relationships of the part the drawing is in.</param>
    /// <param name="readImage">
    /// Reads the image a relationship points to, such as <see cref="OpcPackage.ReadPartAsync"/>
    /// on the resolved target; it is only called for internal relationships.
    /// </param>
    public DrawingReadContext(IReadOnlyList<OpcRelationship> relationships, Func<OpcRelationship, CancellationToken, Task<Result<byte[]>>> readImage)
    {
        ArgumentNullException.ThrowIfNull(relationships);
        ArgumentNullException.ThrowIfNull(readImage);
        Add(relationships);
        this.readImage = readImage;
    }

    /// <summary>
    /// Creates a context for images already in memory, for <see cref="DrawingReader.Read"/>: a
    /// package that loaded the part and its images first reads the drawing without I/O.
    /// </summary>
    /// <param name="relationships">The relationships of the part the drawing is in.</param>
    /// <param name="readImage">Gives the bytes of the image a relationship points to; only called for internal relationships.</param>
    public DrawingReadContext(IReadOnlyList<OpcRelationship> relationships, Func<OpcRelationship, Result<byte[]>> readImage)
    {
        ArgumentNullException.ThrowIfNull(relationships);
        ArgumentNullException.ThrowIfNull(readImage);
        Add(relationships);
        readImageNow = readImage;
    }

    /// <summary>
    /// The theme's colours by name as <c>#RRGGBB</c> (<c>accent1</c>, <c>tx1</c>, <c>bg1</c> …),
    /// so that fills, lines, shadows and text in theme colours (a:schemeClr) are read with the
    /// colour they show (see <see cref="DrawingColors"/>); null, the default, leaves them out as
    /// before.
    /// </summary>
    public IReadOnlyDictionary<string, string>? SchemeColors { get; init; }

    /// <summary>True when the context was made with a synchronous image reader, for <see cref="DrawingReader.Read"/>.</summary>
    public bool IsSynchronous => readImageNow is not null;

    internal bool TryGetRelationship(string? id, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out OpcRelationship? relationship)
    {
        relationship = null;

        return id is not null && relationships.TryGetValue(id, out relationship) == true;
    }

    internal Task<Result<byte[]>> ReadImageAsync(OpcRelationship relationship, CancellationToken cancellationToken)
    {
        if (readImage is null)
        {
            return Task.FromResult(readImageNow!(relationship));
        }

        return readImage(relationship, cancellationToken);
    }

    internal Result<byte[]> ReadImage(OpcRelationship relationship)
    {
        return readImageNow!(relationship);
    }

    private void Add(IReadOnlyList<OpcRelationship> list)
    {
        foreach (OpcRelationship relationship in list)
        {
            relationships.TryAdd(relationship.Id, relationship);
        }
    }
}

/// <summary>The kinds of <see cref="ShapeTreeItem"/>.</summary>
public enum ShapeTreeItemKind
{
    /// <summary>A shape or a connector (sp, cxnSp).</summary>
    Shape,

    /// <summary>A picture (pic).</summary>
    Picture,

    /// <summary>A group (grpSp).</summary>
    Group,

    /// <summary>A graphic frame (graphicFrame): a chart, a table, SmartArt or an OLE object.</summary>
    GraphicFrame,
}

/// <summary>One element of a shape tree as <see cref="DrawingReader"/> read it, with its place.</summary>
public sealed record ShapeTreeItem
{
    /// <summary>Creates an item.</summary>
    /// <param name="kind">What it is.</param>
    /// <param name="id">The shape id (cNvPr id).</param>
    /// <param name="name">The name (cNvPr name).</param>
    /// <param name="bounds">The place and size from its xfrm, in the coordinates of what it is in.</param>
    public ShapeTreeItem(ShapeTreeItemKind kind, int id, string name, DrawingBounds bounds)
    {
        Kind = kind;
        Id = id;
        Name = name;
        Bounds = bounds;
    }

    /// <summary>What it is.</summary>
    public ShapeTreeItemKind Kind { get; }

    /// <summary>The shape id, unique within the part.</summary>
    public int Id { get; }

    /// <summary>The name in the Selection Pane.</summary>
    public string Name { get; }

    /// <summary>The place and size.</summary>
    public DrawingBounds Bounds { get; }

    /// <summary>The alternative text, or null.</summary>
    public string? Description { get; init; }

    /// <summary>The shape, when <see cref="Kind"/> is <see cref="ShapeTreeItemKind.Shape"/>.</summary>
    public Shape? Shape { get; init; }

    /// <summary>The picture, when <see cref="Kind"/> is <see cref="ShapeTreeItemKind.Picture"/>.</summary>
    public Picture? Picture { get; init; }

    /// <summary>The group, when <see cref="Kind"/> is <see cref="ShapeTreeItemKind.Group"/>.</summary>
    public ShapeGroup? Group { get; init; }

    /// <summary>What a graphic frame shows (a:graphicData uri), such as the chart namespace; null for the other kinds.</summary>
    public string? GraphicDataUri { get; init; }

    /// <summary>The relationship id of the chart part a graphic frame shows, or null. Read the part with <see cref="ChartReader"/>.</summary>
    public string? ChartRelationshipId { get; init; }
}
