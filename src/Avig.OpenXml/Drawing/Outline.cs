namespace Avig.OpenXml.Drawing;

/// <summary>The dash patterns of an <see cref="Outline"/>, as on Office's Dash type menu.</summary>
public enum LineDash
{
    /// <summary>A solid line.</summary>
    Solid,

    /// <summary>Round dots.</summary>
    RoundDot,

    /// <summary>Square dots.</summary>
    SquareDot,

    /// <summary>Dashes.</summary>
    Dash,

    /// <summary>Dash, dot.</summary>
    DashDot,

    /// <summary>Long dashes.</summary>
    LongDash,

    /// <summary>Long dash, dot.</summary>
    LongDashDot,

    /// <summary>Long dash, dot, dot.</summary>
    LongDashDotDot,

    /// <summary>Dots as long as the line is wide, with gaps three times that (<c>dot</c>).</summary>
    Dot,

    /// <summary>Short dash, dot, at the line's width (<c>sysDashDot</c>).</summary>
    SystemDashDot,

    /// <summary>Short dash, dot, dot, at the line's width (<c>sysDashDotDot</c>).</summary>
    SystemDashDotDot,
}

/// <summary>
/// The line around a shape, the line of a line shape or a chart series, or a chart's border:
/// a colour, a width and a dash pattern, or <see cref="None"/>.
/// </summary>
public sealed partial record Outline
{
    private readonly double transparency;

    private Outline()
    {
    }

    /// <summary>Creates a line.</summary>
    /// <param name="color">The colour as <c>#RRGGBB</c>.</param>
    /// <param name="width">The width in points, 0 to 1584; Office's default is 0.75.</param>
    public Outline(string color, double width = 0.75)
    {
        Color = DrawingValues.Color(color, nameof(color));
        Width = DrawingValues.Range(width, 0, 1584, nameof(width), "A line's width");
    }

    /// <summary>No line.</summary>
    public static Outline None { get; } = new();

    /// <summary>The colour as <c>#RRGGBB</c>, or null for <see cref="None"/>.</summary>
    public string? Color { get; }

    /// <summary>The width in points.</summary>
    public double Width { get; }

    /// <summary>The dash pattern.</summary>
    public LineDash Dash { get; init; }

    /// <summary>0 for opaque to 1 for invisible.</summary>
    public double Transparency
    {
        get => transparency;
        init => transparency = DrawingValues.Transparency(value, nameof(Transparency));
    }
}

/// <summary>
/// A shadow behind a shape, a picture or a chart series (or inside a shape, as
/// <see cref="Shape.InnerShadow"/>). The defaults are Office's "Offset: Bottom Right".
/// </summary>
public sealed partial record Shadow
{
    private readonly string color = "#000000";
    private readonly double transparency = 0.6;
    private readonly double blur = 4;
    private readonly double distance = 3;
    private readonly double angle = 45;

    /// <summary>The colour as <c>#RRGGBB</c>; black unless set.</summary>
    public string Color
    {
        get => color;
        init => color = DrawingValues.Color(value, nameof(Color));
    }

    /// <summary>0 for opaque to 1 for invisible; 0.6 unless set.</summary>
    public double Transparency
    {
        get => transparency;
        init => transparency = DrawingValues.Transparency(value, nameof(Transparency));
    }

    /// <summary>How soft the edge is, in points, 0 to 100; 4 unless set.</summary>
    public double Blur
    {
        get => blur;
        init => blur = DrawingValues.Range(value, 0, 100, nameof(Blur), "A shadow's blur");
    }

    /// <summary>How far the shadow falls from the shape, in points, 0 to 200; 3 unless set.</summary>
    public double Distance
    {
        get => distance;
        init => distance = DrawingValues.Range(value, 0, 200, nameof(Distance), "A shadow's distance");
    }

    /// <summary>The direction it falls in, in degrees clockwise from the right: 45 is down to the right, 90 straight down.</summary>
    public double Angle
    {
        get => angle;
        init => angle = DrawingValues.Range(value, -360, 360, nameof(Angle), "A shadow's angle");
    }
}
