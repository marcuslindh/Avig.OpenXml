namespace Avig.OpenXml.Drawing;

/// <summary>Lines drawn side by side as one (a:ln cmpd), as on Office's Compound type menu.</summary>
public enum LineCompound
{
    /// <summary>One line.</summary>
    SingleLine,

    /// <summary>Two lines of the same width.</summary>
    DoubleLine,

    /// <summary>A thick line and a thin one.</summary>
    ThickThin,

    /// <summary>A thin line and a thick one.</summary>
    ThinThick,

    /// <summary>Three lines.</summary>
    TripleLine,
}

/// <summary>How a line's ends look (a:ln cap), as on Office's Cap type menu.</summary>
public enum LineCap
{
    /// <summary>Cut off at the end.</summary>
    Flat,

    /// <summary>Rounded beyond the end.</summary>
    Round,

    /// <summary>Square beyond the end.</summary>
    Square,
}

/// <summary>How a line turns a corner (a:round, a:bevel, a:miter), as on Office's Join type menu.</summary>
public enum LineJoin
{
    /// <summary>Rounded.</summary>
    Round,

    /// <summary>Cut off.</summary>
    Bevel,

    /// <summary>Pointed.</summary>
    Miter,
}

/// <summary>Where a shape's line sits on its edge (a:ln algn).</summary>
public enum LineAlignment
{
    /// <summary>Centred on the edge.</summary>
    Center,

    /// <summary>Inside the edge.</summary>
    Inset,
}

/// <summary>The arrowheads of a line (ST_LineEndType), as on Office's Begin and End Arrow type menus.</summary>
public enum LineEndType
{
    /// <summary>None.</summary>
    None,

    /// <summary>A filled triangle (Office's Arrow).</summary>
    Triangle,

    /// <summary>A stealth arrow.</summary>
    Stealth,

    /// <summary>A diamond.</summary>
    Diamond,

    /// <summary>An oval.</summary>
    Oval,

    /// <summary>An open arrow (<c>arrow</c>).</summary>
    Arrow,
}

/// <summary>The width or the length of an arrowhead (ST_LineEndWidth, ST_LineEndLength).</summary>
public enum LineEndSize
{
    /// <summary>Medium, the default.</summary>
    Medium,

    /// <summary>Small.</summary>
    Small,

    /// <summary>Large.</summary>
    Large,
}

/// <summary>An arrowhead at the start (a:headEnd) or the end (a:tailEnd) of a line.</summary>
/// <param name="Type">The arrowhead.</param>
/// <param name="Width">How wide it is.</param>
/// <param name="Length">How long it is.</param>
public sealed record LineEnd(LineEndType Type, LineEndSize Width = LineEndSize.Medium, LineEndSize Length = LineEndSize.Medium)
{
    /// <summary>The arrowhead.</summary>
    public LineEndType Type { get; init; } = Defined(Type);

    /// <summary>How wide it is.</summary>
    public LineEndSize Width { get; init; } = Defined(Width);

    /// <summary>How long it is.</summary>
    public LineEndSize Length { get; init; } = Defined(Length);

    private static T Defined<T>(T value)
        where T : struct, Enum
    {
        if (Enum.IsDefined(value) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(value), value, "Not a value of " + typeof(T).Name + ".");
        }

        return value;
    }
}

/// <summary>One dash and the gap after it in a line's own dash pattern (a:custDash/a:ds), each as a fraction of the line's width (1 is as long as the line is wide).</summary>
/// <param name="Dash">The dash's length, 0 to 1000 widths.</param>
/// <param name="Space">The gap's length, 0 to 1000 widths.</param>
public sealed record DashStop(double Dash, double Space)
{
    /// <summary>The dash's length, in widths of the line.</summary>
    public double Dash { get; init; } = DrawingValues.Range(Dash, 0, 1000, nameof(Dash), "A dash");

    /// <summary>The gap's length, in widths of the line.</summary>
    public double Space { get; init; } = DrawingValues.Range(Space, 0, 1000, nameof(Space), "A gap");
}

public sealed partial record Outline
{
    private readonly LineCompound compound;
    private readonly LineAlignment alignment;
    private readonly double? miterLimit;
    private readonly IReadOnlyList<DashStop> customDash = [];

    /// <summary>Creates a line painted with a gradient or a pattern (a solid fill makes the same line as <see cref="Outline(string, double)"/>).</summary>
    /// <param name="fill">A solid, gradient or pattern fill.</param>
    /// <param name="width">The width in points, 0 to 1584.</param>
    public Outline(Fill fill, double width = 0.75)
        : this(FirstColor(fill), width)
    {
        if (fill.Kind == FillKind.Solid)
        {
            Transparency = fill.Transparency;

            return;
        }

        LineFill = fill;
    }

    /// <summary>The gradient or pattern the line is painted with, or null for <see cref="Color"/>.</summary>
    public Fill? LineFill { get; }

    /// <summary>Lines side by side; <see cref="LineCompound.SingleLine"/> unless set.</summary>
    public LineCompound Compound
    {
        get => compound;
        init => compound = Defined(value, nameof(Compound));
    }

    /// <summary>How the ends look, or null for the application's default (flat in Office).</summary>
    public LineCap? Cap { get; init; }

    /// <summary>How corners are turned, or null for the application's default.</summary>
    public LineJoin? Join { get; init; }

    /// <summary>For a <see cref="LineJoin.Miter"/>, how far the point may reach, 0 to 1000 widths (a:miter lim), or null for Office's 8.</summary>
    public double? MiterLimit
    {
        get => miterLimit;
        init
        {
            if (value is double limit)
            {
                DrawingValues.Range(limit, 0, 1000, nameof(MiterLimit), "A miter limit");
            }

            miterLimit = value;
        }
    }

    /// <summary>Where the line sits on the edge; centred unless set.</summary>
    public LineAlignment Alignment
    {
        get => alignment;
        init => alignment = Defined(value, nameof(Alignment));
    }

    /// <summary>The arrowhead at the start, or null for none.</summary>
    public LineEnd? Head { get; init; }

    /// <summary>The arrowhead at the end, or null for none.</summary>
    public LineEnd? Tail { get; init; }

    /// <summary>A dash pattern of the line's own, which wins over <see cref="Dash"/>; empty for none.</summary>
    public IReadOnlyList<DashStop> CustomDash
    {
        get => customDash;
        init
        {
            ArgumentNullException.ThrowIfNull(value);

            if (value.Count > 100)
            {
                throw new ArgumentException("A dash pattern has at most 100 dashes.", nameof(CustomDash));
            }

            foreach (DashStop stop in value)
            {
                ArgumentNullException.ThrowIfNull(stop, nameof(CustomDash));
            }

            customDash = [.. value];
        }
    }

    private static string FirstColor(Fill fill)
    {
        ArgumentNullException.ThrowIfNull(fill);

        return fill.Kind switch
        {
            FillKind.Solid or FillKind.Pattern => fill.Color!,
            FillKind.Gradient => fill.Stops[0].Color,
            _ => throw new ArgumentException("A line is painted with a solid, gradient or pattern fill.", nameof(fill)),
        };
    }

    private static T Defined<T>(T value, string parameter)
        where T : struct, Enum
    {
        if (Enum.IsDefined(value) == false)
        {
            throw new ArgumentOutOfRangeException(parameter, value, "Not a value of " + typeof(T).Name + ".");
        }

        return value;
    }
}

public sealed partial record Shadow
{
    private readonly double scaleX = 1;
    private readonly double scaleY = 1;
    private readonly double skewX;
    private readonly double skewY;

    /// <summary>The horizontal scale of a shadow behind, -2 to 2; 1 unless set. Office's perspective shadows scale and skew it.</summary>
    public double ScaleX
    {
        get => scaleX;
        init => scaleX = DrawingValues.Range(value, -2, 2, nameof(ScaleX), "A shadow's scale");
    }

    /// <summary>The vertical scale of a shadow behind, -2 to 2; 1 unless set.</summary>
    public double ScaleY
    {
        get => scaleY;
        init => scaleY = DrawingValues.Range(value, -2, 2, nameof(ScaleY), "A shadow's scale");
    }

    /// <summary>The horizontal skew of a shadow behind, in degrees, -90 to 90.</summary>
    public double SkewX
    {
        get => skewX;
        init => skewX = DrawingValues.Range(value, -90, 90, nameof(SkewX), "A shadow's skew");
    }

    /// <summary>The vertical skew of a shadow behind, in degrees, -90 to 90.</summary>
    public double SkewY
    {
        get => skewY;
        init => skewY = DrawingValues.Range(value, -90, 90, nameof(SkewY), "A shadow's skew");
    }

    /// <summary>What a shadow behind is lined up to when it is scaled or skewed; the top-left corner unless set.</summary>
    public RectangleAlignment Alignment { get; init; } = RectangleAlignment.TopLeft;

    /// <summary>Whether a shadow behind turns with the shape.</summary>
    public bool RotateWithShape { get; init; }
}
