namespace Avig.OpenXml.Drawing;

public sealed partial record Shape
{
    private readonly string? title;

    /// <summary>A geometry of its own instead of the preset (a:custGeom), or null. <see cref="Preset"/> is then <c>rect</c>, the box the text sits in.</summary>
    public CustomGeometry? Geometry { get; init; }

    /// <summary>
    /// True for a connector (p:cxnSp): a line or one of the straight, elbow and curved connectors
    /// (<c>straightConnector1</c>, <c>bentConnector2</c> … <c>curvedConnector5</c>). A connector has
    /// no fill and no text, and its ends can be glued to shapes.
    /// </summary>
    public bool IsConnector => Geometry is null && PresetGeometry.IsConnector(Preset);

    /// <summary>Where a connector's start is glued, or null for nowhere.</summary>
    public ConnectionSite? StartConnection { get; init; }

    /// <summary>Where a connector's end is glued, or null for nowhere.</summary>
    public ConnectionSite? EndConnection { get; init; }

    /// <summary>A glow around it, or null for none.</summary>
    public GlowEffect? Glow { get; init; }

    /// <summary>A reflection below it, or null for none.</summary>
    public ReflectionEffect? Reflection { get; init; }

    /// <summary>A blur of the whole shape, or null for none.</summary>
    public BlurEffect? Blur { get; init; }

    /// <summary>One of the older preset shadows, or null for none.</summary>
    public PresetShadow? PresetShadow { get; init; }

    /// <summary>A fill laid over its own, or null for none.</summary>
    public FillOverlayEffect? FillOverlay { get; init; }

    /// <summary>The camera and the light that turn it in 3-D, or null for flat.</summary>
    public Scene3D? Scene3D { get; init; }

    /// <summary>Its bevels, depth, contour and material, or null for none.</summary>
    public Shape3D? Shape3D { get; init; }

    /// <summary>The theme style it refers to (p:style), or null for what Office gives a new shape of its kind.</summary>
    public ShapeStyle? Style { get; init; }

    /// <summary>What a user may not do to it in Office, or null for nothing locked.</summary>
    public ShapeLocks? Locks { get; init; }

    /// <summary>Hidden (cNvPr hidden): it is not shown, but stays in the Selection Pane.</summary>
    public bool IsHidden { get; init; }

    /// <summary>Decorative (adec:decorative): screen readers skip it, so it needs no alternative text.</summary>
    public bool IsDecorative { get; init; }

    /// <summary>The title of the alternative text (cNvPr title), or null.</summary>
    public string? Title
    {
        get => title;
        init => title = DrawingText(value, nameof(Title));
    }

    /// <summary>The names of all 187 of DrawingML's preset geometries (ST_ShapeType), in the schema's order, for <see cref="FromPreset(string)"/>.</summary>
    public static IReadOnlyList<string> PresetNames => PresetGeometry.Names;

    /// <summary>A shape drawn with a geometry of its own: a freeform, a curve or merged shapes.</summary>
    /// <param name="geometry">The paths.</param>
    /// <returns>The shape.</returns>
    public static Shape FromGeometry(CustomGeometry geometry)
    {
        ArgumentNullException.ThrowIfNull(geometry);

        return new Shape("rect") { Geometry = geometry };
    }
}

public sealed partial record ShapeGroup
{
    private readonly double rotation;

    /// <summary>The turn of the whole group in degrees, clockwise, -360 to 360.</summary>
    public double Rotation
    {
        get => rotation;
        init => rotation = DrawingValues.Range(value, -360, 360, nameof(Rotation), "A rotation");
    }

    /// <summary>Mirror the group left to right.</summary>
    public bool FlipHorizontal { get; init; }

    /// <summary>Mirror the group top to bottom.</summary>
    public bool FlipVertical { get; init; }

    /// <summary>What a user may not do to the group in Office, or null for nothing locked.</summary>
    public ShapeLocks? Locks { get; init; }

    /// <summary>Hidden (cNvPr hidden).</summary>
    public bool IsHidden { get; init; }
}
