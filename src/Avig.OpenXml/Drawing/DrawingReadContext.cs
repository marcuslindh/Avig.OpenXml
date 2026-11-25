namespace Avig.OpenXml.Drawing;

/// <summary>
/// What <see cref="DrawingReader"/> needs from the part the drawing is in: its relationships, to
/// turn an r:id into a link or an image, and a way to read an image part's bytes. The package
/// that owns the part makes one; nothing outside the package is ever followed.
/// </summary>
public sealed class DrawingReadContext
{
    private readonly Dictionary<string, OpcRelationship> relationships = new(StringComparer.Ordinal);
    private readonly Func<OpcRelationship, CancellationToken, Task<Result<byte[]>>> readImage;

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

        foreach (OpcRelationship relationship in relationships)
        {
            this.relationships.TryAdd(relationship.Id, relationship);
        }

        this.readImage = readImage;
    }

    internal bool TryGetRelationship(string? id, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out OpcRelationship? relationship)
    {
        relationship = null;

        return id is not null && relationships.TryGetValue(id, out relationship) == true;
    }

    internal Task<Result<byte[]>> ReadImageAsync(OpcRelationship relationship, CancellationToken cancellationToken)
    {
        return readImage(relationship, cancellationToken);
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
