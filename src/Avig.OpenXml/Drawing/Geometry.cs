namespace Avig.OpenXml.Drawing;

/// <summary>
/// A geometry of its own (a:custGeom): a freeform, a curve, a scribble or the result of merging
/// shapes. It is one or more paths, each drawn in its own coordinates (<see cref="GeometryPath.Width"/>
/// and <see cref="GeometryPath.Height"/>), which are stretched over the shape's frame.
/// </summary>
/// <example>
/// <code>
/// CustomGeometry triangle = new([new GeometryPath(100, 100, [PathCommand.MoveTo(50, 0), PathCommand.LineTo(100, 100), PathCommand.LineTo(0, 100), PathCommand.Close()])]);
/// Shape freeform = Shape.FromGeometry(triangle) with { Fill = Fill.Solid("#2384F9") };
/// </code>
/// </example>
public sealed record CustomGeometry
{
    /// <summary>The most paths one geometry holds.</summary>
    public const int MaxPaths = 1000;

    /// <summary>Creates a geometry.</summary>
    /// <param name="paths">The paths, 1 to <see cref="MaxPaths"/>, drawn in this order.</param>
    public CustomGeometry(IReadOnlyList<GeometryPath> paths)
    {
        ArgumentNullException.ThrowIfNull(paths);

        if (paths.Count == 0 || paths.Count > MaxPaths)
        {
            throw new ArgumentException($"A geometry holds 1 to {MaxPaths} paths.", nameof(paths));
        }

        foreach (GeometryPath path in paths)
        {
            ArgumentNullException.ThrowIfNull(path, nameof(paths));
        }

        Paths = [.. paths];
    }

    /// <summary>The paths.</summary>
    public IReadOnlyList<GeometryPath> Paths { get; }
}

/// <summary>How the inside of a <see cref="GeometryPath"/> is painted (a:path fill).</summary>
public enum PathFill
{
    /// <summary>With the shape's fill.</summary>
    Normal,

    /// <summary>Not at all: the path is only a line.</summary>
    None,

    /// <summary>With the fill made lighter.</summary>
    Lighten,

    /// <summary>With the fill made a little lighter.</summary>
    LightenLess,

    /// <summary>With the fill made darker.</summary>
    Darken,

    /// <summary>With the fill made a little darker.</summary>
    DarkenLess,
}

/// <summary>One path of a <see cref="CustomGeometry"/> (a:path): the commands that draw it, in its own coordinates.</summary>
public sealed record GeometryPath
{
    /// <summary>The most commands one path holds.</summary>
    public const int MaxCommands = 100_000;

    private readonly PathFill fill;

    /// <summary>Creates a path.</summary>
    /// <param name="width">The width of the path's coordinates, which the shape's width is; 0 for the shape's own EMU.</param>
    /// <param name="height">The height of the path's coordinates.</param>
    /// <param name="commands">The commands, at least one; the first is usually a <see cref="PathCommand.MoveTo(long, long)"/>.</param>
    public GeometryPath(long width, long height, IReadOnlyList<PathCommand> commands)
    {
        ArgumentNullException.ThrowIfNull(commands);

        if (width < 0 || width > DrawingValues.MaxCoordinate || height < 0 || height > DrawingValues.MaxCoordinate)
        {
            throw new ArgumentOutOfRangeException(nameof(width), "A path's width and height go from 0 to 27 273 042 316 900.");
        }

        if (commands.Count == 0 || commands.Count > MaxCommands)
        {
            throw new ArgumentException($"A path holds 1 to {MaxCommands} commands.", nameof(commands));
        }

        foreach (PathCommand command in commands)
        {
            ArgumentNullException.ThrowIfNull(command, nameof(commands));
        }

        Width = width;
        Height = height;
        Commands = [.. commands];
    }

    /// <summary>The width of the path's coordinates.</summary>
    public long Width { get; }

    /// <summary>The height of the path's coordinates.</summary>
    public long Height { get; }

    /// <summary>The commands.</summary>
    public IReadOnlyList<PathCommand> Commands { get; }

    /// <summary>How the inside is painted; <see cref="PathFill.Normal"/> unless set.</summary>
    public PathFill Fill
    {
        get => fill;
        init
        {
            if (Enum.IsDefined(value) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(Fill), value, "Not a path fill.");
            }

            fill = value;
        }
    }

    /// <summary>Whether the shape's line is drawn along the path; true unless set.</summary>
    public bool Stroke { get; init; } = true;

    /// <summary>Whether a 3-D depth may be given to the path; true unless set.</summary>
    public bool ExtrusionOk { get; init; } = true;
}

/// <summary>The kinds of <see cref="PathCommand"/>.</summary>
public enum PathCommandKind
{
    /// <summary>Starts a new figure at a point (a:moveTo).</summary>
    MoveTo,

    /// <summary>A straight line to a point (a:lnTo).</summary>
    LineTo,

    /// <summary>A cubic Bézier curve through two control points to a point (a:cubicBezTo).</summary>
    CubicBezierTo,

    /// <summary>A quadratic Bézier curve through one control point to a point (a:quadBezTo).</summary>
    QuadraticBezierTo,

    /// <summary>An arc of an ellipse from the current point (a:arcTo).</summary>
    ArcTo,

    /// <summary>A line back to the start of the figure, which closes it (a:close).</summary>
    Close,
}

/// <summary>A point in a path's coordinates.</summary>
/// <param name="Left">From the left (x).</param>
/// <param name="Top">From the top (y).</param>
public readonly record struct PathPoint(long Left, long Top);

/// <summary>One command of a <see cref="GeometryPath"/>.</summary>
public sealed record PathCommand
{
    private PathCommand(PathCommandKind kind, IReadOnlyList<PathPoint> points)
    {
        foreach (PathPoint point in points)
        {
            if (Math.Abs(point.Left) > DrawingValues.MaxCoordinate || Math.Abs(point.Top) > DrawingValues.MaxCoordinate)
            {
                throw new ArgumentOutOfRangeException(nameof(points), "A path's points go from -27 273 042 316 900 to 27 273 042 316 900.");
            }
        }

        Kind = kind;
        Points = points;
    }

    /// <summary>What the command does.</summary>
    public PathCommandKind Kind { get; }

    /// <summary>The points: one for a move or a line, two (control, end) for a quadratic curve, three (two controls, end) for a cubic one, none for an arc or a close.</summary>
    public IReadOnlyList<PathPoint> Points { get; }

    /// <summary>The horizontal radius of an arc, in the path's coordinates.</summary>
    public long WidthRadius { get; private init; }

    /// <summary>The vertical radius of an arc.</summary>
    public long HeightRadius { get; private init; }

    /// <summary>Where on the ellipse an arc starts, in degrees clockwise from the right.</summary>
    public double StartAngle { get; private init; }

    /// <summary>How far an arc sweeps, in degrees; negative counter-clockwise.</summary>
    public double SweepAngle { get; private init; }

    /// <summary>Starts a new figure at a point.</summary>
    /// <param name="left">From the left.</param>
    /// <param name="top">From the top.</param>
    /// <returns>The command.</returns>
    public static PathCommand MoveTo(long left, long top)
    {
        return new PathCommand(PathCommandKind.MoveTo, [new PathPoint(left, top)]);
    }

    /// <summary>A straight line to a point.</summary>
    /// <param name="left">From the left.</param>
    /// <param name="top">From the top.</param>
    /// <returns>The command.</returns>
    public static PathCommand LineTo(long left, long top)
    {
        return new PathCommand(PathCommandKind.LineTo, [new PathPoint(left, top)]);
    }

    /// <summary>A cubic Bézier curve.</summary>
    /// <param name="control1">The first control point.</param>
    /// <param name="control2">The second control point.</param>
    /// <param name="end">Where the curve ends.</param>
    /// <returns>The command.</returns>
    public static PathCommand CubicBezierTo(PathPoint control1, PathPoint control2, PathPoint end)
    {
        return new PathCommand(PathCommandKind.CubicBezierTo, [control1, control2, end]);
    }

    /// <summary>A quadratic Bézier curve.</summary>
    /// <param name="control">The control point.</param>
    /// <param name="end">Where the curve ends.</param>
    /// <returns>The command.</returns>
    public static PathCommand QuadraticBezierTo(PathPoint control, PathPoint end)
    {
        return new PathCommand(PathCommandKind.QuadraticBezierTo, [control, end]);
    }

    /// <summary>An arc of an ellipse, from the current point.</summary>
    /// <param name="widthRadius">The ellipse's horizontal radius, in the path's coordinates.</param>
    /// <param name="heightRadius">The ellipse's vertical radius.</param>
    /// <param name="startAngle">Where on the ellipse the arc starts, in degrees clockwise from the right, -360 to 360.</param>
    /// <param name="sweepAngle">How far it sweeps, -360 to 360 degrees.</param>
    /// <returns>The command.</returns>
    public static PathCommand ArcTo(long widthRadius, long heightRadius, double startAngle, double sweepAngle)
    {
        if (widthRadius < 0 || widthRadius > DrawingValues.MaxCoordinate || heightRadius < 0 || heightRadius > DrawingValues.MaxCoordinate)
        {
            throw new ArgumentOutOfRangeException(nameof(widthRadius), "An arc's radii go from 0 to 27 273 042 316 900.");
        }

        return new PathCommand(PathCommandKind.ArcTo, [])
        {
            WidthRadius = widthRadius,
            HeightRadius = heightRadius,
            StartAngle = DrawingValues.Range(startAngle, -360, 360, nameof(startAngle), "An arc's start angle"),
            SweepAngle = DrawingValues.Range(sweepAngle, -360, 360, nameof(sweepAngle), "An arc's sweep"),
        };
    }

    /// <summary>Closes the figure with a line back to its start.</summary>
    /// <returns>The command.</returns>
    public static PathCommand Close()
    {
        return new PathCommand(PathCommandKind.Close, []);
    }
}

/// <summary>
/// Where the end of a connector is glued (a:stCxn, a:endCxn): a shape by its id on the slide or
/// sheet, and the connection site on it. A rectangle's sites are 0 at the top, 1 at the left, 2 at
/// the bottom and 3 at the right; other geometries have their own (an ellipse has 8).
/// </summary>
/// <param name="ShapeId">The id of the shape (cNvPr id).</param>
/// <param name="Site">The connection site, from 0.</param>
public sealed record ConnectionSite(int ShapeId, int Site)
{
    /// <summary>The id of the shape, 1 or more.</summary>
    public int ShapeId { get; init; } = DrawingValues.Range(ShapeId, 1, int.MaxValue, nameof(ShapeId), "A shape id");

    /// <summary>The connection site, 0 to 1000.</summary>
    public int Site { get; init; } = DrawingValues.Range(Site, 0, 1000, nameof(Site), "A connection site");
}

/// <summary>
/// What a user may not do to a shape in Office (a:spLocks, a:cxnSpLocks, a:grpSpLocks). Office
/// honours them in its editor; they do not protect the file.
/// </summary>
public sealed record ShapeLocks
{
    /// <summary>Cannot be grouped.</summary>
    public bool NoGroup { get; init; }

    /// <summary>A group cannot be ungrouped.</summary>
    public bool NoUngroup { get; init; }

    /// <summary>Cannot be selected.</summary>
    public bool NoSelect { get; init; }

    /// <summary>Cannot be rotated.</summary>
    public bool NoRotate { get; init; }

    /// <summary>Keeps its aspect ratio when resized.</summary>
    public bool NoChangeAspect { get; init; }

    /// <summary>Cannot be moved.</summary>
    public bool NoMove { get; init; }

    /// <summary>Cannot be resized.</summary>
    public bool NoResize { get; init; }

    /// <summary>Its points cannot be edited.</summary>
    public bool NoEditPoints { get; init; }

    /// <summary>Its adjust handles cannot be moved.</summary>
    public bool NoAdjustHandles { get; init; }

    /// <summary>Its arrowheads cannot be changed.</summary>
    public bool NoChangeArrowheads { get; init; }

    /// <summary>Its geometry cannot be changed.</summary>
    public bool NoChangeShapeType { get; init; }

    /// <summary>Its text cannot be edited.</summary>
    public bool NoTextEdit { get; init; }

    internal bool Any => NoGroup || NoUngroup || NoSelect || NoRotate || NoChangeAspect || NoMove || NoResize || NoEditPoints || NoAdjustHandles || NoChangeArrowheads || NoChangeShapeType || NoTextEdit;
}

/// <summary>One of the theme references of a <see cref="ShapeStyle"/> (a:lnRef, a:fillRef, a:effectRef).</summary>
/// <param name="Index">
/// The style in the theme's list: 0 for none, 1 to 3 for subtle to intense (a:fillRef also takes
/// 1001 to 1003 for the background fills).
/// </param>
/// <param name="Color">The theme colour it is painted with, such as <c>accent1</c>, <c>dk1</c> or <c>phClr</c>.</param>
/// <param name="Shade">A shade of the colour, 0 to 1 (DrawingML a:shade), or null for none.</param>
public sealed record StyleReference(int Index, string Color = "accent1", double? Shade = null)
{
    /// <summary>The style in the theme's list.</summary>
    public int Index { get; init; } = StyleIndex(Index);

    /// <summary>The theme colour.</summary>
    public string Color { get; init; } = ShapeStyle.SchemeColor(Color, nameof(Color));

    /// <summary>A shade of the colour, or null.</summary>
    public double? Shade { get; init; } = Fraction(Shade);

    private static int StyleIndex(int index)
    {
        if ((index >= 0 && index <= 3) || (index >= 1001 && index <= 1003))
        {
            return index;
        }

        throw new ArgumentOutOfRangeException(nameof(index), index, "A style index is 0 to 3, or 1001 to 1003 for a background fill.");
    }

    private static double? Fraction(double? value)
    {
        if (value is double fraction)
        {
            DrawingValues.Range(fraction, 0, 1, nameof(value), "A shade");
        }

        return value;
    }
}

/// <summary>
/// The theme style a shape refers to (p:style), as Office's Shape Styles gallery sets it: the
/// line, fill and effect of the theme's format scheme, and the font colour. What the shape sets
/// itself wins over it.
/// </summary>
public sealed record ShapeStyle
{
    private static readonly HashSet<string> SchemeColors = new(StringComparer.Ordinal)
    {
        "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink",
        "phClr", "dk1", "lt1", "dk2", "lt2",
    };

    private readonly string fontColor = "lt1";

    /// <summary>The line: one of the theme's line styles; unless set, what Office gives a new shape (style 2 in accent1, shaded 15 %).</summary>
    public StyleReference Line { get; init; } = new(2, "accent1", 0.15);

    /// <summary>The fill: one of the theme's fill styles.</summary>
    public StyleReference Fill { get; init; } = new(1);

    /// <summary>The effect: one of the theme's effect styles.</summary>
    public StyleReference Effect { get; init; } = new(0);

    /// <summary>The theme font of the text: <c>minor</c> (body), <c>major</c> (headings) or <c>none</c>.</summary>
    public ThemeFont Font { get; init; } = ThemeFont.Minor;

    /// <summary>The theme colour of the text, such as <c>lt1</c> or <c>dk1</c>.</summary>
    public string FontColor
    {
        get => fontColor;
        init => fontColor = SchemeColor(value, nameof(FontColor));
    }

    internal static string SchemeColor(string value, string parameter)
    {
        ArgumentNullException.ThrowIfNull(value, parameter);

        if (SchemeColors.Contains(value) == false)
        {
            throw new ArgumentException($"'{value}' is not a theme colour; expected one such as accent1, dk1, lt1 or phClr.", parameter);
        }

        return value;
    }
}

/// <summary>The theme font a <see cref="ShapeStyle"/> gives the text (a:fontRef idx).</summary>
public enum ThemeFont
{
    /// <summary>The body font.</summary>
    Minor,

    /// <summary>The headings font.</summary>
    Major,

    /// <summary>No theme font.</summary>
    None,
}
