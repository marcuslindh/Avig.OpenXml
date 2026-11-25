namespace Avig.OpenXml.Drawing;

/// <summary>
/// How the inside of a shape, a chart area or a chart series is painted: <see cref="None"/>, a
/// <see cref="Solid(string, double)"/> colour, a linear <see cref="Gradient(double, GradientStop[])"/>,
/// a two-colour <see cref="Pattern(FillPattern, string, string)"/> or a <see cref="Picture(Drawing.Picture)"/>.
/// </summary>
public sealed partial record Fill
{
    private Fill(FillKind kind, string? color, double transparency, double angle, IReadOnlyList<GradientStop> stops)
    {
        Kind = kind;
        Color = color;
        Transparency = transparency;
        Angle = angle;
        Stops = stops;
    }

    /// <summary>Nothing: the shape or area is see-through.</summary>
    public static Fill None { get; } = new(FillKind.None, null, 0, 0, []);

    /// <summary>Which kind of fill this is.</summary>
    public FillKind Kind { get; }

    /// <summary>The colour of a solid fill, or the foreground of a pattern, as <c>#RRGGBB</c>; null for the other kinds.</summary>
    public string? Color { get; }

    /// <summary>The transparency of a solid fill, from 0 (opaque) to 1 (invisible).</summary>
    public double Transparency { get; }

    /// <summary>The direction of a gradient in degrees, clockwise from left to right.</summary>
    public double Angle { get; }

    /// <summary>The stops of a gradient; empty for the other kinds.</summary>
    public IReadOnlyList<GradientStop> Stops { get; }

    /// <summary>The pattern of a pattern fill.</summary>
    public FillPattern PatternType { get; private init; }

    /// <summary>The background colour of a pattern fill, as <c>#RRGGBB</c>; null for the other kinds.</summary>
    public string? BackgroundColor { get; private init; }

    /// <summary>The image of a picture fill, stretched over the shape; null for the other kinds.</summary>
    public Picture? Image { get; private init; }

    /// <summary>One colour.</summary>
    /// <param name="color">The colour as <c>#RRGGBB</c>.</param>
    /// <param name="transparency">0 for opaque to 1 for invisible.</param>
    /// <returns>The fill.</returns>
    public static Fill Solid(string color, double transparency = 0)
    {
        return new Fill(FillKind.Solid, DrawingValues.Color(color, nameof(color)), DrawingValues.Transparency(transparency, nameof(transparency)), 0, []);
    }

    /// <summary>A linear gradient through two or more colours.</summary>
    /// <param name="angle">
    /// The direction in degrees, clockwise: 0 runs from left to right, 90 from top to bottom and
    /// 45 from the top-left corner to the bottom-right.
    /// </param>
    /// <param name="stops">The colours and where they sit, at least two.</param>
    /// <returns>The fill.</returns>
    public static Fill Gradient(double angle, params GradientStop[] stops)
    {
        ArgumentNullException.ThrowIfNull(stops);
        DrawingValues.Range(angle, -360, 360, nameof(angle), "The angle");

        if (stops.Length < 2 || stops.Length > 10)
        {
            throw new ArgumentException("A gradient has 2 to 10 stops.", nameof(stops));
        }

        foreach (GradientStop stop in stops)
        {
            ArgumentNullException.ThrowIfNull(stop, nameof(stops));
        }

        return new Fill(FillKind.Gradient, null, 0, angle, [.. stops]);
    }

    /// <summary>A linear gradient from one colour to another.</summary>
    /// <param name="from">The colour at the start, as <c>#RRGGBB</c>.</param>
    /// <param name="to">The colour at the end.</param>
    /// <param name="angle">The direction in degrees, as in <see cref="Gradient(double, GradientStop[])"/>; 90 runs from top to bottom.</param>
    /// <returns>The fill.</returns>
    public static Fill Gradient(string from, string to, double angle = 90)
    {
        return Gradient(angle, new GradientStop(0, from), new GradientStop(1, to));
    }

    /// <summary>A pattern in two colours, as on the Pattern fill tab in Office.</summary>
    /// <param name="pattern">The pattern.</param>
    /// <param name="foreground">The colour of the pattern's lines and dots, as <c>#RRGGBB</c>.</param>
    /// <param name="background">The colour behind them.</param>
    /// <returns>The fill.</returns>
    public static Fill Pattern(FillPattern pattern, string foreground, string background)
    {
        if (Enum.IsDefined(pattern) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(pattern), pattern, "Not a fill pattern.");
        }

        return new Fill(FillKind.Pattern, DrawingValues.Color(foreground, nameof(foreground)), 0, 0, [])
        {
            PatternType = pattern,
            BackgroundColor = DrawingValues.Color(background, nameof(background)),
        };
    }

    /// <summary>An image stretched over the shape. A chart cannot take one.</summary>
    /// <param name="image">The image; an SVG is stored with its fallback, as a picture is.</param>
    /// <returns>The fill.</returns>
    public static Fill Picture(Picture image)
    {
        ArgumentNullException.ThrowIfNull(image);

        return new Fill(FillKind.Picture, null, 0, 0, []) { Image = image };
    }
}

/// <summary>The kinds of <see cref="Fill"/>.</summary>
public enum FillKind
{
    /// <summary>No fill.</summary>
    None,

    /// <summary>One colour.</summary>
    Solid,

    /// <summary>A linear gradient.</summary>
    Gradient,

    /// <summary>A pattern in two colours.</summary>
    Pattern,

    /// <summary>An image.</summary>
    Picture,

    /// <summary>The fill of the group the shape is in (a:grpFill).</summary>
    Group,

    /// <summary>The slide's background, as if the shape were a hole in what lies under it (p:sp useBgFill).</summary>
    Background,
}

/// <summary>A colour at a position in a <see cref="Fill.Gradient(double, GradientStop[])"/>.</summary>
public sealed record GradientStop
{
    /// <summary>Creates a stop.</summary>
    /// <param name="position">Where the colour sits, from 0 (the start) to 1 (the end).</param>
    /// <param name="color">The colour as <c>#RRGGBB</c>.</param>
    /// <param name="transparency">0 for opaque to 1 for invisible.</param>
    public GradientStop(double position, string color, double transparency = 0)
    {
        Position = DrawingValues.Range(position, 0, 1, nameof(position), "A stop's position");
        Color = DrawingValues.Color(color, nameof(color));
        Transparency = DrawingValues.Transparency(transparency, nameof(transparency));
    }

    /// <summary>Where the colour sits, from 0 to 1.</summary>
    public double Position { get; }

    /// <summary>The colour as <c>#RRGGBB</c>.</summary>
    public string Color { get; }

    /// <summary>0 for opaque to 1 for invisible.</summary>
    public double Transparency { get; }
}

/// <summary>The patterns of a <see cref="Fill.Pattern(FillPattern, string, string)"/> (ST_PresetPatternVal).</summary>
public enum FillPattern
{
    /// <summary>5 percent.</summary>
    Percent5,

    /// <summary>10 percent.</summary>
    Percent10,

    /// <summary>20 percent.</summary>
    Percent20,

    /// <summary>25 percent.</summary>
    Percent25,

    /// <summary>30 percent.</summary>
    Percent30,

    /// <summary>40 percent.</summary>
    Percent40,

    /// <summary>50 percent.</summary>
    Percent50,

    /// <summary>60 percent.</summary>
    Percent60,

    /// <summary>70 percent.</summary>
    Percent70,

    /// <summary>75 percent.</summary>
    Percent75,

    /// <summary>80 percent.</summary>
    Percent80,

    /// <summary>90 percent.</summary>
    Percent90,

    /// <summary>Horizontal lines.</summary>
    Horizontal,

    /// <summary>Vertical lines.</summary>
    Vertical,

    /// <summary>Light horizontal lines.</summary>
    LightHorizontal,

    /// <summary>Light vertical lines.</summary>
    LightVertical,

    /// <summary>Dark horizontal lines.</summary>
    DarkHorizontal,

    /// <summary>Dark vertical lines.</summary>
    DarkVertical,

    /// <summary>Narrow horizontal lines.</summary>
    NarrowHorizontal,

    /// <summary>Narrow vertical lines.</summary>
    NarrowVertical,

    /// <summary>Dashed horizontal lines.</summary>
    DashedHorizontal,

    /// <summary>Dashed vertical lines.</summary>
    DashedVertical,

    /// <summary>A cross.</summary>
    Cross,

    /// <summary>Downward diagonals.</summary>
    DownwardDiagonal,

    /// <summary>Upward diagonals.</summary>
    UpwardDiagonal,

    /// <summary>Light downward diagonals.</summary>
    LightDownwardDiagonal,

    /// <summary>Light upward diagonals.</summary>
    LightUpwardDiagonal,

    /// <summary>Dark downward diagonals.</summary>
    DarkDownwardDiagonal,

    /// <summary>Dark upward diagonals.</summary>
    DarkUpwardDiagonal,

    /// <summary>Wide downward diagonals.</summary>
    WideDownwardDiagonal,

    /// <summary>Wide upward diagonals.</summary>
    WideUpwardDiagonal,

    /// <summary>Dashed downward diagonals.</summary>
    DashedDownwardDiagonal,

    /// <summary>Dashed upward diagonals.</summary>
    DashedUpwardDiagonal,

    /// <summary>A diagonal cross.</summary>
    DiagonalCross,

    /// <summary>A small checker board.</summary>
    SmallCheck,

    /// <summary>A large checker board.</summary>
    LargeCheck,

    /// <summary>A small grid.</summary>
    SmallGrid,

    /// <summary>A large grid.</summary>
    LargeGrid,

    /// <summary>A dotted grid.</summary>
    DottedGrid,

    /// <summary>Small confetti.</summary>
    SmallConfetti,

    /// <summary>Large confetti.</summary>
    LargeConfetti,

    /// <summary>Horizontal bricks.</summary>
    HorizontalBrick,

    /// <summary>Diagonal bricks.</summary>
    DiagonalBrick,

    /// <summary>Solid diamonds.</summary>
    SolidDiamond,

    /// <summary>Outlined diamonds.</summary>
    OpenDiamond,

    /// <summary>Dotted diamonds.</summary>
    DottedDiamond,

    /// <summary>Plaid.</summary>
    Plaid,

    /// <summary>Spheres.</summary>
    Sphere,

    /// <summary>Weave.</summary>
    Weave,

    /// <summary>Divots.</summary>
    Divot,

    /// <summary>Shingles.</summary>
    Shingle,

    /// <summary>Waves.</summary>
    Wave,

    /// <summary>A trellis.</summary>
    Trellis,

    /// <summary>Zig zag.</summary>
    ZigZag,
}
