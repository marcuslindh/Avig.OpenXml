namespace Avig.OpenXml.Drawing;

/// <summary>How the series of a column, bar, line or area chart stand to each other.</summary>
public enum ChartGrouping
{
    /// <summary>Side by side (columns and bars), or each on its own (lines and areas); the default.</summary>
    Default,

    /// <summary>On top of each other.</summary>
    Stacked,

    /// <summary>On top of each other, each category adding up to 100 %.</summary>
    PercentStacked,

    /// <summary>One behind the other along a depth axis, in a 3-D chart (<see cref="ChartSpace.ThreeD"/>): the 3-D column, area and line charts.</summary>
    Depth,
}

/// <summary>The shape of the columns or bars of a 3-D chart.</summary>
public enum BarShape
{
    /// <summary>A box; the default.</summary>
    Box,

    /// <summary>A cylinder.</summary>
    Cylinder,

    /// <summary>A cone, its point at the value.</summary>
    Cone,

    /// <summary>A cone cut off at the value, its point at the axis' maximum.</summary>
    ConeToMax,

    /// <summary>A pyramid, its point at the value.</summary>
    Pyramid,

    /// <summary>A pyramid cut off at the value, its point at the axis' maximum.</summary>
    PyramidToMax,
}

/// <summary>How a radar chart draws its series.</summary>
public enum RadarStyle
{
    /// <summary>Lines, with markers unless the series' <see cref="ChartSpaceSeries.Marker"/> is None; the default.</summary>
    Marker,

    /// <summary>Filled areas.</summary>
    Filled,
}

/// <summary>Which points of a pie-of-pie or bar-of-pie chart go to the second plot.</summary>
public enum ChartSplit
{
    /// <summary>The application's choice: the last points; the default.</summary>
    Auto,

    /// <summary>The last <see cref="ChartSpace.SplitValue"/> points.</summary>
    Position,

    /// <summary>The points below <see cref="ChartSpace.SplitValue"/>.</summary>
    Value,

    /// <summary>The points below <see cref="ChartSpace.SplitValue"/> percent of the whole.</summary>
    Percent,
}

/// <summary>The kind of a trendline.</summary>
public enum TrendlineKind
{
    /// <summary>A straight line.</summary>
    Linear,

    /// <summary>An exponential curve.</summary>
    Exponential,

    /// <summary>A logarithmic curve.</summary>
    Logarithmic,

    /// <summary>A moving average over <see cref="Trendline.Period"/> points.</summary>
    MovingAverage,

    /// <summary>A polynomial of <see cref="Trendline.Order"/>.</summary>
    Polynomial,

    /// <summary>A power curve.</summary>
    Power,
}

/// <summary>Which way error bars go.</summary>
public enum ErrorBarDirection
{
    /// <summary>Along the value (y) axis; the default.</summary>
    Y,

    /// <summary>Along the x axis, for scatter and bubble charts.</summary>
    X,
}

/// <summary>Which side of a point error bars are drawn on.</summary>
public enum ErrorBarSide
{
    /// <summary>Both; the default.</summary>
    Both,

    /// <summary>Above (or right of) the point.</summary>
    Plus,

    /// <summary>Below (or left of) the point.</summary>
    Minus,
}

/// <summary>What sets the length of error bars.</summary>
public enum ErrorBarAmount
{
    /// <summary>A fixed amount, <see cref="ErrorBars.Value"/>; the default.</summary>
    Fixed,

    /// <summary>A percentage of each value.</summary>
    Percentage,

    /// <summary>A number of standard deviations.</summary>
    StandardDeviation,

    /// <summary>The standard error.</summary>
    StandardError,
}

/// <summary>The viewing angle and depth of a 3-D chart (c:view3D).</summary>
public sealed record ChartView3D
{
    private readonly int rotationX = 15;
    private readonly int rotationY = 20;
    private readonly int perspective = 30;
    private readonly int depthPercent = 100;
    private readonly int gapDepth = 150;

    /// <summary>The tilt, -90 to 90 degrees; 15 by default (pie charts take 0 to 90).</summary>
    public int RotationX
    {
        get => rotationX;
        init => rotationX = DrawingValues.Range(value, -90, 90, nameof(RotationX), "The X rotation");
    }

    /// <summary>The turn, 0 to 360 degrees; 20 by default.</summary>
    public int RotationY
    {
        get => rotationY;
        init => rotationY = DrawingValues.Range(value, 0, 360, nameof(RotationY), "The Y rotation");
    }

    /// <summary>The field of view, 0 to 240 (twice the degrees); 30 by default. Has no effect with <see cref="RightAngleAxes"/>.</summary>
    public int Perspective
    {
        get => perspective;
        init => perspective = DrawingValues.Range(value, 0, 240, nameof(Perspective), "The perspective");
    }

    /// <summary>Keep the axes at right angles, without perspective; on by default, as for column and bar charts.</summary>
    public bool RightAngleAxes { get; init; } = true;

    /// <summary>The depth in percent of the width, 20 to 2000; 100 by default.</summary>
    public int DepthPercent
    {
        get => depthPercent;
        init => depthPercent = DrawingValues.Range(value, 20, 2000, nameof(DepthPercent), "The depth");
    }

    /// <summary>The space between series in depth, in percent of a column's depth, 0 to 500; 150 by default.</summary>
    public int GapDepth
    {
        get => gapDepth;
        init => gapDepth = DrawingValues.Range(value, 0, 500, nameof(GapDepth), "The gap depth");
    }
}

/// <summary>A table of the chart's values under the category axis (c:dTable).</summary>
public sealed record ChartDataTable
{
    /// <summary>Draw lines between the rows.</summary>
    public bool HorizontalBorders { get; init; } = true;

    /// <summary>Draw lines between the columns.</summary>
    public bool VerticalBorders { get; init; } = true;

    /// <summary>Draw a line around the table.</summary>
    public bool Outline { get; init; } = true;

    /// <summary>Show the legend's keys beside the series' names.</summary>
    public bool ShowLegendKeys { get; init; } = true;
}

/// <summary>A trendline through a series' points (c:trendline).</summary>
public sealed record Trendline
{
    private readonly int order = 2;
    private readonly int period = 2;
    private readonly string? name;

    /// <summary>The kind of curve.</summary>
    public TrendlineKind Kind { get; init; }

    /// <summary>The order of a polynomial, 2 to 6.</summary>
    public int Order
    {
        get => order;
        init => order = DrawingValues.Range(value, 2, 6, nameof(Order), "A polynomial's order");
    }

    /// <summary>The points a moving average is taken over, 2 to 255.</summary>
    public int Period
    {
        get => period;
        init => period = DrawingValues.Range(value, 2, 255, nameof(Period), "A moving average's period");
    }

    /// <summary>The name in the legend, or null for the application's ("Linear (Series 1)").</summary>
    public string? Name
    {
        get => name;
        init
        {
            if (value is not null && value.Length > 255)
            {
                throw new ArgumentException("A trendline's name holds at most 255 characters.", nameof(Name));
            }

            if (value is not null)
            {
                DrawingValues.XmlText(value, nameof(Name));
            }

            name = value;
        }
    }

    /// <summary>Periods to draw ahead of the last point, 0 or more.</summary>
    public double Forward { get; init; }

    /// <summary>Periods to draw before the first point, 0 or more.</summary>
    public double Backward { get; init; }

    /// <summary>Where the line crosses the value axis, or null to let it be worked out.</summary>
    public double? Intercept { get; init; }

    /// <summary>Show the equation on the chart.</summary>
    public bool DisplayEquation { get; init; }

    /// <summary>Show the R-squared value on the chart.</summary>
    public bool DisplayRSquared { get; init; }

    /// <summary>The line, or null for the application's.</summary>
    public Outline? Outline { get; init; }
}

/// <summary>Error bars on a series' points (c:errBars).</summary>
public sealed record ErrorBars
{
    private readonly double value = 1;

    /// <summary>Along which axis.</summary>
    public ErrorBarDirection Direction { get; init; }

    /// <summary>On which side of the point.</summary>
    public ErrorBarSide Side { get; init; }

    /// <summary>What sets their length.</summary>
    public ErrorBarAmount Amount { get; init; }

    /// <summary>The amount, the percentage or the number of standard deviations; not used for the standard error. 0 or more, 1 by default.</summary>
    public double Value
    {
        get => value;
        init
        {
            if (double.IsFinite(value) == false || value < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(Value), value, "An error bar's value must be a finite number of 0 or more.");
            }

            this.value = value;
        }
    }

    /// <summary>Leave out the cross line at the end.</summary>
    public bool NoEndCap { get; init; }

    /// <summary>The line, or null for the application's.</summary>
    public Outline? Outline { get; init; }
}

/// <summary>More of an axis: minor gridlines, the direction, a logarithmic scale and the minor unit.</summary>
public sealed partial record ChartAxis
{
    private readonly double? logBase;
    private readonly double? minorUnit;

    /// <summary>Draw the minor gridlines too.</summary>
    public bool MinorGridlines { get; init; }

    /// <summary>The minor gridlines' colour, width and dash, or null for the default.</summary>
    public Outline? MinorGridlineOutline { get; init; }

    /// <summary>Run the axis the other way: categories from right to left, or values from the top down.</summary>
    public bool Reversed { get; init; }

    /// <summary>A logarithmic scale with this base, 2 to 1000, or null for a linear one. For value axes.</summary>
    public double? LogBase
    {
        get => logBase;
        init
        {
            if (value is double given)
            {
                DrawingValues.Range(given, 2, 1000, nameof(LogBase), "A logarithm's base");
            }

            logBase = value;
        }
    }

    /// <summary>The step between minor gridlines on a value axis, or null for automatic.</summary>
    public double? MinorUnit
    {
        get => minorUnit;
        init
        {
            if (value is double unit && (double.IsFinite(unit) == false || unit <= 0))
            {
                throw new ArgumentOutOfRangeException(nameof(MinorUnit), unit, "The minor unit must be above 0.");
            }

            minorUnit = value;
        }
    }
}
