namespace Avig.OpenXml.Drawing;

/// <summary>What a series of the other chart types has: bubble sizes, exploded slices, trendlines and error bars.</summary>
public sealed partial record ChartSpaceSeries
{
    private readonly int? explosion;
    private readonly IReadOnlyList<Trendline> trendlines = [];
    private readonly IReadOnlyList<ErrorBars> errorBars = [];

    /// <summary>The sizes of a bubble series' bubbles (c:bubbleSize), or null for all the same.</summary>
    public ChartData? BubbleSizes { get; init; }

    /// <summary>How far a pie's or doughnut's slices are pulled out, in percent of the radius, 0 to 400, or null for not at all.</summary>
    public int? Explosion
    {
        get => explosion;
        init
        {
            if (value is int given)
            {
                DrawingValues.Range(given, 0, 400, nameof(Explosion), "An explosion");
            }

            explosion = value;
        }
    }

    /// <summary>Trendlines through the points, at most 8. For column, bar, line, area, scatter and bubble series that are not stacked or 3-D.</summary>
    public IReadOnlyList<Trendline> Trendlines
    {
        get => trendlines;
        init
        {
            ArgumentNullException.ThrowIfNull(value);

            if (value.Count > 8)
            {
                throw new ArgumentException("A series holds at most 8 trendlines.", nameof(Trendlines));
            }

            foreach (Trendline item in value)
            {
                ArgumentNullException.ThrowIfNull(item, nameof(Trendlines));
            }

            trendlines = [.. value];
        }
    }

    /// <summary>Error bars, at most one per direction (two for scatter and bubble series, along x and y). For column, bar, line, area, scatter and bubble series that are not 3-D.</summary>
    public IReadOnlyList<ErrorBars> ErrorBars
    {
        get => errorBars;
        init
        {
            ArgumentNullException.ThrowIfNull(value);

            if (value.Count > 2 || value.Select(item => item?.Direction).Distinct().Count() != value.Count)
            {
                throw new ArgumentException("A series holds at most one set of error bars per direction.", nameof(ErrorBars));
            }

            foreach (ErrorBars item in value)
            {
                ArgumentNullException.ThrowIfNull(item, nameof(ErrorBars));
            }

            errorBars = [.. value];
        }
    }
}

/// <summary>What the other chart types have: 3-D, percent stacking, bar shapes, radar, bubble, stock, surface and pie-of-pie options, a data table and the chart's style.</summary>
public sealed partial record ChartSpace
{
    private readonly int? secondPlotSize;
    private readonly double? splitValue;
    private readonly int? bubbleScale;
    private readonly int? style;
    private readonly string? externalData;

    /// <summary>How the series stand to each other: side by side, stacked, stacked to 100 %, or in depth. <see cref="Stacked"/> is the same as <see cref="ChartGrouping.Stacked"/>.</summary>
    public ChartGrouping Grouping
    {
        get => grouping;
        init
        {
            if (Enum.IsDefined(value) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(Grouping), value, "Not a grouping.");
            }

            grouping = value;
        }
    }

    /// <summary>
    /// Draw the chart in 3-D with this view, or null for 2-D. Column, bar, line, area, pie and
    /// surface charts can be 3-D; a surface without it is a contour seen from above.
    /// </summary>
    public ChartView3D? ThreeD { get; init; }

    /// <summary>The shape of a 3-D chart's columns or bars.</summary>
    public BarShape BarShape { get; init; }

    /// <summary>Lines with markers, or filled areas, in a radar chart.</summary>
    public RadarStyle RadarStyle { get; init; }

    /// <summary>Draw a surface as a wireframe, without fills.</summary>
    public bool Wireframe { get; init; }

    /// <summary>
    /// Draw lines between a scatter chart's points in the application's colours. Off, as by
    /// default, a scatter series without its own <see cref="ChartSpaceSeries.Outline"/> or
    /// <see cref="ChartSpaceSeries.Smooth"/> shows its points alone.
    /// </summary>
    public bool ScatterLines { get; init; }

    /// <summary>Give the bubbles of a bubble chart a 3-D look.</summary>
    public bool Bubble3D { get; init; }

    /// <summary>The bubbles' size in percent of the default, 0 to 300, or null for 100.</summary>
    public int? BubbleScale
    {
        get => bubbleScale;
        init
        {
            if (value is int scale)
            {
                DrawingValues.Range(scale, 0, 300, nameof(BubbleScale), "A bubble scale");
            }

            bubbleScale = value;
        }
    }

    /// <summary>Show bubbles with a negative size.</summary>
    public bool ShowNegativeBubbles { get; init; }

    /// <summary>The second plot of a pie-of-pie or bar-of-pie chart, in percent of the first, 5 to 200, or null for 75.</summary>
    public int? SecondPlotSize
    {
        get => secondPlotSize;
        init
        {
            if (value is int size)
            {
                DrawingValues.Range(size, 5, 200, nameof(SecondPlotSize), "The second plot's size");
            }

            secondPlotSize = value;
        }
    }

    /// <summary>Which points of a pie-of-pie or bar-of-pie chart go to the second plot.</summary>
    public ChartSplit Split { get; init; }

    /// <summary>The number of points, the value or the percentage <see cref="Split"/> divides at, or null for the application's.</summary>
    public double? SplitValue
    {
        get => splitValue;
        init
        {
            if (value is double given && (double.IsFinite(given) == false || given < 0))
            {
                throw new ArgumentOutOfRangeException(nameof(SplitValue), given, "A split value must be a finite number of 0 or more.");
            }

            splitValue = value;
        }
    }

    /// <summary>
    /// Draw lines from the highest to the lowest value in each category: on by default in a stock
    /// chart (null), off in a line chart. For line and stock charts.
    /// </summary>
    public bool? HighLowLines { get; init; }

    /// <summary>
    /// Draw bars from the first to the last series in each category: on by default (null) in a
    /// stock chart of four series (open, high, low, close), off otherwise. For line and stock charts.
    /// </summary>
    public bool? UpDownBars { get; init; }

    /// <summary>Draw lines from each point down to the category axis. For line and area charts.</summary>
    public bool DropLines { get; init; }

    /// <summary>
    /// Draw lines between the stacks of a stacked column or bar chart, and from a pie-of-pie's or
    /// bar-of-pie's first plot to its second: on by default there (null), off otherwise.
    /// </summary>
    public bool? SeriesLines { get; init; }

    /// <summary>The depth axis of a 3-D chart with <see cref="ChartGrouping.Depth"/>, of a 3-D line chart and of a surface, or null for the default.</summary>
    public ChartAxis? SeriesAxis { get; init; }

    /// <summary>A table of the values under the plot, or null for none. For charts with a category axis.</summary>
    public ChartDataTable? DataTable { get; init; }

    /// <summary>One of the application's 48 built-in chart styles (c:style), 1 to 48, or null for the default 2.</summary>
    public int? Style
    {
        get => style;
        init
        {
            if (value is int given)
            {
                DrawingValues.Range(given, 1, 48, nameof(Style), "A chart style");
            }

            style = value;
        }
    }

    /// <summary>
    /// The relationship id of the workbook that holds the chart's data (c:externalData), or null
    /// for none. The package that owns the chart writes the workbook and the relationship; with it
    /// the application's Edit Data opens the workbook.
    /// </summary>
    public string? ExternalData
    {
        get => externalData;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 255))
            {
                throw new ArgumentException("A relationship id holds 1 to 255 characters.", nameof(ExternalData));
            }

            if (value is not null)
            {
                DrawingValues.XmlText(value, nameof(ExternalData));
            }

            externalData = value;
        }
    }

    /// <summary>True for the types with a category (or x) axis and a value axis.</summary>
    internal bool HasAxes => Type is not (ChartType.Pie or ChartType.Doughnut or ChartType.PieOfPie or ChartType.BarOfPie);

    /// <summary>Throws ArgumentException for options the type does not take, which Office would repair.</summary>
    private void ValidateKinds()
    {
        if (ThreeD is not null && (Type is ChartType.Column or ChartType.Bar or ChartType.Line or ChartType.Area or ChartType.Pie or ChartType.Surface) == false)
        {
            throw new ArgumentException($"A {Type} chart cannot be 3-D; column, bar, line, area, pie and surface charts can.", nameof(ThreeD));
        }

        if (Grouping is ChartGrouping.Stacked or ChartGrouping.PercentStacked && (Type is ChartType.Column or ChartType.Bar or ChartType.Line or ChartType.Area) == false)
        {
            throw new ArgumentException($"A {Type} chart cannot be stacked; column, bar, line and area charts can.", nameof(Grouping));
        }

        if (Grouping == ChartGrouping.Depth && (ThreeD is null || (Type is ChartType.Column or ChartType.Bar or ChartType.Line or ChartType.Area) == false))
        {
            throw new ArgumentException("Series in depth need a 3-D column, bar, line or area chart (ThreeD).", nameof(Grouping));
        }

        if (BarShape != BarShape.Box && (ThreeD is null || (Type is ChartType.Column or ChartType.Bar) == false))
        {
            throw new ArgumentException("Cylinders, cones and pyramids need a 3-D column or bar chart (ThreeD).", nameof(BarShape));
        }

        if (ThreeD is not null && Series.Any(item => TypeOf(item) != Type || item.SecondaryAxis == true))
        {
            throw new ArgumentException("A 3-D chart cannot combine types or use a secondary axis.", nameof(ThreeD));
        }

        if (Type == ChartType.Pie && ThreeD is ChartView3D pieView && pieView.RotationX < 0)
        {
            throw new ArgumentException("A 3-D pie tilts 0 to 90 degrees (ThreeD.RotationX).", nameof(ThreeD));
        }

        if (Type == ChartType.Surface && Series.Any(item => item.PointColors.Count > 0 || item.DataLabels is not null) == true || (Type == ChartType.Surface && DataLabels is not null))
        {
            throw new ArgumentException("A surface takes no point colours or data labels; its bands are coloured by value.", nameof(Series));
        }

        if (Type == ChartType.Stock)
        {
            int stock = Series.Count(item => TypeOf(item) == ChartType.Stock);

            if (stock is < 3 or > 4)
            {
                throw new ArgumentException($"A stock chart has three series (high, low, close) or four (open, high, low, close), not {stock}.", nameof(Series));
            }
        }

        foreach (ChartSpaceSeries item in Series)
        {
            ChartType type = TypeOf(item);
            bool flat = ThreeD is null;
            bool plain = type is ChartType.Column or ChartType.Bar or ChartType.Line or ChartType.Area or ChartType.Scatter or ChartType.Bubble;

            if (item.Trendlines.Count > 0 && (plain == false || flat == false || (type == Type && Stacked == true)))
            {
                throw new ArgumentException($"A {type} series cannot take trendlines here; 2-D column, bar, line, area, scatter and bubble series that are not stacked can.", nameof(Series));
            }

            if (item.ErrorBars.Count > 0 && (plain == false || flat == false))
            {
                throw new ArgumentException($"A {type} series cannot take error bars here; 2-D column, bar, line, area, scatter and bubble series can.", nameof(Series));
            }

            if (item.ErrorBars.Any(bars => bars.Direction == ErrorBarDirection.X) == true && (type is ChartType.Scatter or ChartType.Bubble) == false)
            {
                throw new ArgumentException("Error bars along x are for scatter and bubble series.", nameof(Series));
            }

            if (item.Explosion is not null && (type is ChartType.Pie or ChartType.Doughnut or ChartType.PieOfPie or ChartType.BarOfPie) == false)
            {
                throw new ArgumentException("Only the slices of a pie or doughnut can be pulled out (Explosion).", nameof(Series));
            }
        }
    }
}
