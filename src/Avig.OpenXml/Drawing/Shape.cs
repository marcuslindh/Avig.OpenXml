namespace Avig.OpenXml.Drawing;

/// <summary>
/// The shapes that <see cref="Shape"/> draws by name. Any other of DrawingML's preset
/// geometries is made with <see cref="Shape.FromPreset(string)"/>, and its kind is null.
/// </summary>
public enum ShapeKind
{
    /// <summary>A rectangle.</summary>
    Rectangle,

    /// <summary>A rectangle with rounded corners; <see cref="Shape.CornerRadius"/> sets how round.</summary>
    RoundedRectangle,

    /// <summary>An ellipse, or a circle in a square.</summary>
    Ellipse,

    /// <summary>An isosceles triangle pointing up.</summary>
    Triangle,

    /// <summary>A diamond.</summary>
    Diamond,

    /// <summary>A hexagon.</summary>
    Hexagon,

    /// <summary>A block arrow pointing right.</summary>
    RightArrow,

    /// <summary>A block arrow pointing left.</summary>
    LeftArrow,

    /// <summary>A block arrow pointing up.</summary>
    UpArrow,

    /// <summary>A block arrow pointing down.</summary>
    DownArrow,

    /// <summary>A chevron pointing right.</summary>
    Chevron,

    /// <summary>A pentagon pointing right, as a step in a process.</summary>
    Pentagon,

    /// <summary>A straight line from the top-left corner to the bottom-right (flip it for the other diagonal); only its <see cref="Shape.Outline"/> shows.</summary>
    Line,

    /// <summary>A text box: a rectangle without fill or line, for text alone.</summary>
    TextBox,
}

/// <summary>An adjust value of a preset geometry (a:gd in a:avLst), such as how far an arrow's head reaches.</summary>
/// <param name="Name">The guide's name, such as <c>adj</c> or <c>adj1</c>.</param>
/// <param name="Value">The value, usually in 100 000ths of the shape's shorter side.</param>
public sealed record ShapeAdjustment(string Name, long Value);

/// <summary>
/// A shape: a preset geometry with a fill, an outline, shadows, text and a link. Everything left
/// unset looks as a shape just inserted in Office does. Avig.Kalkyl puts it on a sheet with
/// <c>AddShape</c>; the package that owns the drawing decides where it goes.
/// </summary>
/// <example>
/// <code>
/// Shape card = new(ShapeKind.RoundedRectangle)
/// {
///     Fill = Fill.Gradient("#2384F9", "#605776", 45),
///     Outline = Outline.None,
///     Shadow = new Shadow(),
///     Text = new ShapeText("Projects", new TextFont { Size = 14, Bold = true, Color = "#FFFFFF" }),
/// };
/// </code>
/// </example>
public sealed partial record Shape
{
    private readonly double? cornerRadius;
    private readonly double rotation;
    private readonly double? softEdges;
    private readonly string? name;
    private readonly string? description;
    private readonly IReadOnlyList<ShapeAdjustment> adjustments = [];

    /// <summary>Creates a shape.</summary>
    /// <param name="kind">The geometry. Any other preset is made with <see cref="FromPreset(string)"/>.</param>
    public Shape(ShapeKind kind)
    {
        if (Enum.IsDefined(kind) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(kind), kind, "Not a shape kind.");
        }

        Kind = kind;
        Preset = PresetGeometry.NameOf(kind);
    }

    private Shape(string preset)
    {
        Kind = null;
        Preset = preset;
    }

    /// <summary>The geometry, or null for a preset without a name here (see <see cref="Preset"/>).</summary>
    public ShapeKind? Kind { get; }

    /// <summary>The name of the preset geometry DrawingML stores (a:prstGeom prst), such as <c>roundRect</c> or <c>star5</c>.</summary>
    public string Preset { get; }

    /// <summary>
    /// How round a <see cref="ShapeKind.RoundedRectangle"/>'s corners are, from 0 (square) to 0.5
    /// (half the shorter side, a pill), or null for Office's 0.1667.
    /// </summary>
    public double? CornerRadius
    {
        get => cornerRadius;
        init
        {
            if (value is double radius)
            {
                DrawingValues.Range(radius, 0, 0.5, nameof(CornerRadius), "A corner radius");
            }

            cornerRadius = value;
        }
    }

    /// <summary>
    /// The geometry's adjust values, or empty for its defaults. <see cref="CornerRadius"/>, when
    /// set, wins over an <c>adj</c> here.
    /// </summary>
    public IReadOnlyList<ShapeAdjustment> Adjustments
    {
        get => adjustments;
        init
        {
            ArgumentNullException.ThrowIfNull(value);

            if (value.Count > 16)
            {
                throw new ArgumentException("A shape has at most 16 adjust values.", nameof(Adjustments));
            }

            foreach (ShapeAdjustment adjustment in value)
            {
                ArgumentNullException.ThrowIfNull(adjustment, nameof(Adjustments));

                if (adjustment.Name is null || adjustment.Name.Length == 0 || adjustment.Name.Length > 64 || adjustment.Name.All(char.IsAsciiLetterOrDigit) == false)
                {
                    throw new ArgumentException($"'{adjustment.Name}' is not a guide name; expected letters and digits, such as adj1.", nameof(Adjustments));
                }
            }

            adjustments = [.. value];
        }
    }

    /// <summary>The inside, or null for the application's default (the theme's first accent colour; none for a text box).</summary>
    public Fill? Fill { get; init; }

    /// <summary>The line around it, or null for the application's default (a darker accent; none for a text box).</summary>
    public Outline? Outline { get; init; }

    /// <summary>A shadow behind it, or null for none.</summary>
    public Shadow? Shadow { get; init; }

    /// <summary>A shadow inside its edge, or null for none.</summary>
    public Shadow? InnerShadow { get; init; }

    /// <summary>Softened edges, in points from 0 to 100, or null for sharp ones.</summary>
    public double? SoftEdges
    {
        get => softEdges;
        init
        {
            if (value is double points)
            {
                DrawingValues.Range(points, 0, 100, nameof(SoftEdges), "Soft edges");
            }

            softEdges = value;
        }
    }

    /// <summary>The text inside, or null for none. A <see cref="ShapeKind.Line"/> has no text.</summary>
    public ShapeText? Text { get; init; }

    /// <summary>The turn in degrees, clockwise, -360 to 360.</summary>
    public double Rotation
    {
        get => rotation;
        init => rotation = DrawingValues.Range(value, -360, 360, nameof(Rotation), "A rotation");
    }

    /// <summary>Mirror it left to right; for a line, it then runs from the top-right corner to the bottom-left.</summary>
    public bool FlipHorizontal { get; init; }

    /// <summary>Mirror it top to bottom.</summary>
    public bool FlipVertical { get; init; }

    /// <summary>Where a click on it goes, or null.</summary>
    public Hyperlink? Link { get; init; }

    /// <summary>The name in the Selection Pane, or null for "Rectangle 2" and so on.</summary>
    public string? Name
    {
        get => name;
        init => name = DrawingText(value, nameof(Name));
    }

    /// <summary>The alternative text that screen readers read, or null.</summary>
    public string? Description
    {
        get => description;
        init => description = DrawingText(value, nameof(Description));
    }

    /// <summary>A shape with any of DrawingML's preset geometries (ST_ShapeType), such as <c>star5</c>, <c>cloud</c> or <c>flowChartDecision</c>.</summary>
    /// <param name="preset">The geometry's name, as DrawingML writes it.</param>
    /// <returns>The shape; one of the named kinds when the preset has one (<c>ellipse</c> is <see cref="ShapeKind.Ellipse"/>).</returns>
    public static Shape FromPreset(string preset)
    {
        ArgumentNullException.ThrowIfNull(preset);

        if (PresetGeometry.IsKnown(preset) == false)
        {
            throw new ArgumentException($"'{preset}' is not one of DrawingML's preset geometries.", nameof(preset));
        }

        if (PresetGeometry.TryKindOf(preset, out ShapeKind kind) == true)
        {
            return new Shape(kind);
        }

        return new Shape(preset);
    }

    internal static string? DrawingText(string? value, string parameter)
    {
        if (value is not null && value.Length > 1000)
        {
            throw new ArgumentException("A drawing's name or alternative text holds at most 1000 characters.", parameter);
        }

        if (value is not null)
        {
            DrawingValues.XmlText(value, parameter);
        }

        return value;
    }
}

/// <summary>
/// Shapes and pictures that move and resize as one, with one link: a dashboard's navigation
/// button made of an icon and a label, for example. Each item has its place inside the group;
/// the group is then put on the sheet or the slide, and its items are scaled to fit.
/// </summary>
public sealed partial record ShapeGroup
{
    private readonly string? name;
    private readonly string? description;

    /// <summary>Creates a group.</summary>
    /// <param name="items">The shapes, pictures and groups, at least one, drawn in this order (the last on top).</param>
    public ShapeGroup(IReadOnlyList<GroupItem> items)
    {
        ArgumentNullException.ThrowIfNull(items);

        if (items.Count == 0 || items.Count > 1000)
        {
            throw new ArgumentException("A group holds 1 to 1000 items.", nameof(items));
        }

        foreach (GroupItem item in items)
        {
            ArgumentNullException.ThrowIfNull(item, nameof(items));
        }

        Items = [.. items];
    }

    /// <summary>The shapes, pictures and groups.</summary>
    public IReadOnlyList<GroupItem> Items { get; }

    /// <summary>Where a click on the group goes, or null. Links on the items themselves are kept too.</summary>
    public Hyperlink? Link { get; init; }

    /// <summary>The name in the Selection Pane, or null for "Group 2" and so on.</summary>
    public string? Name
    {
        get => name;
        init => name = Shape.DrawingText(value, nameof(Name));
    }

    /// <summary>The alternative text that screen readers read, or null.</summary>
    public string? Description
    {
        get => description;
        init => description = Shape.DrawingText(value, nameof(Description));
    }
}

/// <summary>A shape, a picture or a group in a <see cref="ShapeGroup"/>, with its place inside the group.</summary>
public sealed record GroupItem
{
    private GroupItem(Shape? shape, Picture? picture, ShapeGroup? group, DrawingBounds bounds)
    {
        Shape = shape;
        Picture = picture;
        Group = group;
        Bounds = bounds;
    }

    /// <summary>The shape, or null when the item is a picture or a group.</summary>
    public Shape? Shape { get; }

    /// <summary>The picture, or null when the item is a shape or a group.</summary>
    public Picture? Picture { get; }

    /// <summary>The group inside the group, or null when the item is a shape or a picture.</summary>
    public ShapeGroup? Group { get; }

    /// <summary>The place inside the group, measured from the group's top-left corner.</summary>
    public DrawingBounds Bounds { get; }

    /// <summary>A shape at a place inside the group.</summary>
    /// <param name="shape">The shape.</param>
    /// <param name="bounds">The place, from the group's top-left corner.</param>
    /// <returns>The item.</returns>
    public static GroupItem Of(Shape shape, DrawingBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(shape);

        return new GroupItem(shape, null, null, bounds);
    }

    /// <summary>A picture at a place inside the group.</summary>
    /// <param name="picture">The picture.</param>
    /// <param name="bounds">The place, from the group's top-left corner.</param>
    /// <returns>The item.</returns>
    public static GroupItem Of(Picture picture, DrawingBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(picture);

        return new GroupItem(null, picture, null, bounds);
    }

    /// <summary>A group inside the group.</summary>
    /// <param name="group">The group.</param>
    /// <param name="bounds">The place, from the outer group's top-left corner.</param>
    /// <returns>The item.</returns>
    public static GroupItem Of(ShapeGroup group, DrawingBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(group);

        return new GroupItem(null, null, group, bounds);
    }
}
