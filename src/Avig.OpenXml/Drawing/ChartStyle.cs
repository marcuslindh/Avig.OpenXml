namespace Avig.OpenXml.Drawing;

/// <summary>The kind of chart, as on the Insert menu of Excel and PowerPoint.</summary>
public enum ChartType
{
    /// <summary>Vertical bars.</summary>
    Column,

    /// <summary>Horizontal bars.</summary>
    Bar,

    /// <summary>Lines with markers.</summary>
    Line,

    /// <summary>Filled areas under lines.</summary>
    Area,

    /// <summary>A pie, from the first series.</summary>
    Pie,

    /// <summary>A ring, one per series.</summary>
    Doughnut,

    /// <summary>Points at x and y; the series' categories are the x values and must be numbers.</summary>
    Scatter,
}

/// <summary>Where the legend goes.</summary>
public enum LegendPosition
{
    /// <summary>No legend.</summary>
    None,

    /// <summary>To the right of the plot; the default.</summary>
    Right,

    /// <summary>Below the plot.</summary>
    Bottom,

    /// <summary>Above the plot.</summary>
    Top,

    /// <summary>To the left of the plot.</summary>
    Left,
}

/// <summary>The marker shapes of a line or scatter series.</summary>
public enum MarkerStyle
{
    /// <summary>No marker: the line alone.</summary>
    None,

    /// <summary>A circle.</summary>
    Circle,

    /// <summary>A square.</summary>
    Square,

    /// <summary>A diamond.</summary>
    Diamond,

    /// <summary>A triangle.</summary>
    Triangle,

    /// <summary>An x.</summary>
    X,

    /// <summary>A star.</summary>
    Star,

    /// <summary>A short horizontal dash.</summary>
    Dash,

    /// <summary>A dot.</summary>
    Dot,

    /// <summary>A plus.</summary>
    Plus,
}

/// <summary>Where a data label sits by its point. Which ones a chart type takes is in <see cref="DataLabels.Position"/>.</summary>
public enum DataLabelPosition
{
    /// <summary>The application's choice for the chart type.</summary>
    Default,

    /// <summary>On the middle of the point.</summary>
    Center,

    /// <summary>Inside the end of a column, bar or slice.</summary>
    InsideEnd,

    /// <summary>Inside the base of a column or bar.</summary>
    InsideBase,

    /// <summary>Outside the end of a column, bar or slice.</summary>
    OutsideEnd,

    /// <summary>Above a line or scatter point.</summary>
    Above,

    /// <summary>Below a line or scatter point.</summary>
    Below,

    /// <summary>Left of a line or scatter point.</summary>
    Left,

    /// <summary>Right of a line or scatter point.</summary>
    Right,

    /// <summary>Where a pie's label fits best.</summary>
    BestFit,
}

/// <summary>Labels on the points of a chart: the value, the percentage, the category or the series' name.</summary>
public sealed record DataLabels
{
    private readonly string? numberFormat;

    /// <summary>Show the value.</summary>
    public bool ShowValue { get; init; } = true;

    /// <summary>Show the share of the whole, for pie and doughnut charts.</summary>
    public bool ShowPercent { get; init; }

    /// <summary>Show the category.</summary>
    public bool ShowCategory { get; init; }

    /// <summary>Show the series' name.</summary>
    public bool ShowSeriesName { get; init; }

    /// <summary>
    /// Where the labels sit. Column and bar take <c>Center</c>, <c>InsideEnd</c>, <c>InsideBase</c>
    /// and, unless stacked, <c>OutsideEnd</c>; line and scatter take <c>Center</c>, <c>Above</c>,
    /// <c>Below</c>, <c>Left</c> and <c>Right</c>; pie takes <c>Center</c>, <c>InsideEnd</c>,
    /// <c>OutsideEnd</c> and <c>BestFit</c>; area and doughnut take <c>Default</c> only. Office
    /// repairs a file with any other, so <see cref="ChartSpace.Validate"/> throws an ArgumentException for it.
    /// </summary>
    public DataLabelPosition Position { get; init; }

    /// <summary>How the labels look, or null for the chart's font.</summary>
    public TextFont? Font { get; init; }

    /// <summary>The number format of the values, such as <c>0%</c>, or null for the cells' own.</summary>
    public string? NumberFormat
    {
        get => numberFormat;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 255))
            {
                throw new ArgumentException("A number format holds 1 to 255 characters.", nameof(NumberFormat));
            }

            numberFormat = value;
        }
    }
}

/// <summary>How an axis of a chart looks: shown or hidden, gridlines, its range, font and number format.</summary>
public sealed record ChartAxis
{
    private readonly double? minimum;
    private readonly double? maximum;
    private readonly double? majorUnit;
    private readonly string? numberFormat;

    /// <summary>Show the axis and its labels. Off, the chart keeps its scale but draws no axis, as in a dashboard.</summary>
    public bool Visible { get; init; } = true;

    /// <summary>Draw the major gridlines, or null for the default: on for the value axis, off for the category axis.</summary>
    public bool? Gridlines { get; init; }

    /// <summary>The gridlines' colour, width and dash, or null for the light grey default.</summary>
    public Outline? GridlineOutline { get; init; }

    /// <summary>The axis line, or null for the default; <see cref="Outline.None"/> hides the line and keeps the labels.</summary>
    public Outline? Line { get; init; }

    /// <summary>The lowest value on a value axis, or null for automatic.</summary>
    public double? Minimum
    {
        get => minimum;
        init => minimum = Finite(value, nameof(Minimum));
    }

    /// <summary>The highest value on a value axis, or null for automatic.</summary>
    public double? Maximum
    {
        get => maximum;
        init => maximum = Finite(value, nameof(Maximum));
    }

    /// <summary>The step between gridlines on a value axis, or null for automatic.</summary>
    public double? MajorUnit
    {
        get => majorUnit;
        init
        {
            if (value is double unit && (double.IsFinite(unit) == false || unit <= 0))
            {
                throw new ArgumentOutOfRangeException(nameof(MajorUnit), unit, "The major unit must be above 0.");
            }

            majorUnit = value;
        }
    }

    /// <summary>The number format of the labels, such as <c># ##0 "kr"</c>, or null for the cells' own.</summary>
    public string? NumberFormat
    {
        get => numberFormat;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 255))
            {
                throw new ArgumentException("A number format holds 1 to 255 characters.", nameof(NumberFormat));
            }

            numberFormat = value;
        }
    }

    /// <summary>How the labels look, or null for the chart's font.</summary>
    public TextFont? Font { get; init; }

    private static double? Finite(double? value, string parameter)
    {
        if (value is double number && double.IsFinite(number) == false)
        {
            throw new ArgumentOutOfRangeException(parameter, number, "An axis bound must be a finite number.");
        }

        return value;
    }
}
