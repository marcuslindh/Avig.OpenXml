namespace Avig.OpenXml.Drawing;

/// <summary>
/// The namespace the elements of a shape tree are written in, which is the one thing about a
/// shape that depends on where it is: <c>xdr:sp</c> in a spreadsheet drawing, <c>p:sp</c> on a
/// slide. The a: elements inside are the same everywhere. The package that owns the drawing
/// makes one of these; this layer does not know the hosts.
/// </summary>
/// <param name="Prefix">The prefix, such as <c>xdr</c> or <c>p</c>.</param>
/// <param name="Uri">The namespace.</param>
public sealed record ShapeTreeNamespace(string Prefix, string Uri)
{
    /// <summary>Write the empty <c>macro</c> (and on a shape <c>textlink</c>) attributes, as a spreadsheet drawing has them.</summary>
    public bool MacroAttributes { get; init; }

    /// <summary>Write the empty application properties (<c>nvPr</c>) after the non-visual properties, as a slide needs them.</summary>
    public bool ApplicationProperties { get; init; }
}

/// <summary>
/// What one part's drawings refer to, collected while their XML is written: shape ids in
/// order, and a relationship id per chart, image and link. An image or a link used more than
/// once shares one id. The package that owns the part writes the targets afterwards, in the
/// order of <see cref="Targets"/>, since a zip has one entry open at a time.
/// </summary>
public sealed class DrawingContext
{
    private readonly HashSet<string> usedIds;
    private readonly Dictionary<object, string> images = new(ReferenceEqualityComparer.Instance);
    private readonly Dictionary<Hyperlink, string> links = [];
    private readonly List<DrawingTarget> targets = [];
    private int nextShapeId;

    /// <summary>Creates a context.</summary>
    /// <param name="tree">The namespace of the shape tree's elements.</param>
    /// <param name="usedRelationshipIds">The relationship ids the part has already; new ones are added to the set.</param>
    /// <param name="firstShapeId">The first shape id to hand out; ids must be unique within the part.</param>
    public DrawingContext(ShapeTreeNamespace tree, HashSet<string> usedRelationshipIds, int firstShapeId)
    {
        ArgumentNullException.ThrowIfNull(tree);
        ArgumentNullException.ThrowIfNull(usedRelationshipIds);
        Tree = tree;
        usedIds = usedRelationshipIds;
        nextShapeId = firstShapeId;
    }

    /// <summary>The namespace of the shape tree's elements.</summary>
    public ShapeTreeNamespace Tree { get; }

    /// <summary>The charts, images and links, in the order their ids were handed out.</summary>
    public IReadOnlyList<DrawingTarget> Targets => targets;

    /// <summary>The next shape id.</summary>
    /// <returns>The id.</returns>
    public int NextShapeId()
    {
        int id = nextShapeId;
        nextShapeId++;

        return id;
    }

    /// <summary>The first <c>rIdN</c> the part does not use yet, which is then taken.</summary>
    /// <returns>The id.</returns>
    public string NextRelationshipId()
    {
        int number = 1;

        while (usedIds.Contains("rId" + DrawingValues.Text(number)) == true)
        {
            number++;
        }

        string id = "rId" + DrawingValues.Text(number);
        usedIds.Add(id);

        return id;
    }

    /// <summary>A relationship id for a chart part, which the owner writes from <paramref name="chart"/>.</summary>
    /// <param name="chart">Whatever the owner needs to write the chart part later.</param>
    /// <returns>The id.</returns>
    public string Chart(object chart)
    {
        ArgumentNullException.ThrowIfNull(chart);
        string id = NextRelationshipId();
        targets.Add(new DrawingTarget(id, chart, null, null));

        return id;
    }

    /// <summary>The relationship id of an image; the same image data shares one.</summary>
    /// <param name="picture">The image.</param>
    /// <returns>The id.</returns>
    public string Image(Picture picture)
    {
        ArgumentNullException.ThrowIfNull(picture);

        if (images.TryGetValue(picture.Identity, out string? existing) == true)
        {
            return existing;
        }

        string id = NextRelationshipId();
        images[picture.Identity] = id;
        targets.Add(new DrawingTarget(id, null, picture, null));

        return id;
    }

    /// <summary>The relationship id of a link; equal links share one.</summary>
    /// <param name="link">The link.</param>
    /// <returns>The id.</returns>
    public string Link(Hyperlink link)
    {
        ArgumentNullException.ThrowIfNull(link);

        if (links.TryGetValue(link, out string? existing) == true)
        {
            return existing;
        }

        string id = NextRelationshipId();
        links[link] = id;
        targets.Add(new DrawingTarget(id, null, null, link));

        return id;
    }
}

/// <summary>A chart, an image or a link a drawing refers to by <see cref="Id"/>. Exactly one of the three is set.</summary>
/// <param name="Id">The relationship id.</param>
/// <param name="Chart">What the owner gave <see cref="DrawingContext.Chart(object)"/>, or null.</param>
/// <param name="Image">The image, or null.</param>
/// <param name="Link">The link, or null.</param>
public sealed record DrawingTarget(string Id, object? Chart, Picture? Image, Hyperlink? Link);
