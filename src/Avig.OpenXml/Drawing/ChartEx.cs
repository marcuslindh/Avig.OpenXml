namespace Avig.OpenXml.Drawing;

/// <summary>The chart types of the newer chart engine (cx:chartSpace, Office 2016 and later).</summary>
public enum ChartExType
{
    /// <summary>A running total, each value going up or down from the last; <see cref="ChartExSeries.Subtotals"/> stand on the axis.</summary>
    Waterfall,

    /// <summary>Nested rectangles sized by value, grouped by the category levels.</summary>
    Treemap,

    /// <summary>Rings of the category levels, the innermost the outermost level, sized by value.</summary>
    Sunburst,

    /// <summary>The values counted into bins (<see cref="ChartExSpace.Binning"/>).</summary>
    Histogram,

    /// <summary>A histogram sorted from the largest bin, with a line of the cumulative percentage.</summary>
    Pareto,

    /// <summary>Quartiles, whiskers and outliers of each series' values, per category.</summary>
    BoxWhisker,

    /// <summary>Bars centred on each other, from the largest, as stages of a process.</summary>
    Funnel,

    /// <summary>A filled map: the categories are countries, regions or postal codes, coloured by value.</summary>
    RegionMap,
}

/// <summary>Where the labels of a treemap's parent categories go.</summary>
public enum TreemapParentLabels
{
    /// <summary>Over the top left of each group; the default.</summary>
    Overlapping,

    /// <summary>In a banner above each group.</summary>
    Banner,

    /// <summary>No parent labels.</summary>
    None,
}

/// <summary>The bins of a histogram or Pareto chart (cx:binning); all null for the application's own.</summary>
public sealed record ChartExBinning
{
    private readonly double? binWidth;
    private readonly int? binCount;

    /// <summary>The width of each bin, above 0, or null.</summary>
    public double? BinWidth
    {
        get => binWidth;
        init
        {
            if (value is double width && (double.IsFinite(width) == false || width <= 0))
            {
                throw new ArgumentOutOfRangeException(nameof(BinWidth), width, "A bin width must be above 0.");
            }

            binWidth = value;
        }
    }

    /// <summary>The number of bins, 1 to 1000, or null; with <see cref="BinWidth"/> too, the width wins.</summary>
    public int? BinCount
    {
        get => binCount;
        init
        {
            if (value is int count)
            {
                DrawingValues.Range(count, 1, 1000, nameof(BinCount), "A bin count");
            }

            binCount = value;
        }
    }

    /// <summary>Values at or below this go to one bin of their own, or null for none.</summary>
    public double? Underflow { get; init; }

    /// <summary>Values above this go to one bin of their own, or null for none.</summary>
    public double? Overflow { get; init; }

    /// <summary>Bins hold their lower bound instead of their upper one: [a, b) instead of (a, b].</summary>
    public bool ClosedLeft { get; init; }
}

/// <summary>The categories of a <see cref="ChartExSpace"/>: one level of names, or several for a treemap or sunburst.</summary>
public sealed record ChartExCategories
{
    private readonly IReadOnlyList<IReadOnlyList<string?>> levels = [];
    private readonly string? formula;

    /// <summary>Creates the categories.</summary>
    /// <param name="levels">The levels, the outermost (the root of a treemap) first and the leaves last; each a name per point.</param>
    public ChartExCategories(IReadOnlyList<IReadOnlyList<string?>> levels)
    {
        Levels = levels;
    }

    /// <summary>The levels, the outermost first and the leaves last; each a name per point.</summary>
    public IReadOnlyList<IReadOnlyList<string?>> Levels
    {
        get => levels;
        init
        {
            ArgumentNullException.ThrowIfNull(value);

            if (value.Count == 0 || value.Count > 32)
            {
                throw new ArgumentException("Categories have 1 to 32 levels.", nameof(Levels));
            }

            List<IReadOnlyList<string?>> copy = [];

            foreach (IReadOnlyList<string?> level in value)
            {
                ArgumentNullException.ThrowIfNull(level, nameof(Levels));

                foreach (string? text in level)
                {
                    if (text is not null)
                    {
                        DrawingValues.XmlText(text, nameof(Levels));
                    }
                }

                copy.Add([.. level]);
            }

            levels = copy;
        }
    }

    /// <summary>The reference over every level, such as <c>Sheet1!$A$2:$C$17</c>, or null for names alone.</summary>
    public string? Formula
    {
        get => formula;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 8192))
            {
                throw new ArgumentException("A formula holds 1 to 8192 characters.", nameof(Formula));
            }

            if (value is not null)
            {
                DrawingValues.XmlText(value, nameof(Formula));
            }

            formula = value;
        }
    }

    /// <summary>One level of names.</summary>
    /// <param name="names">The names.</param>
    /// <returns>The categories.</returns>
    public static ChartExCategories Flat(IReadOnlyList<string?> names)
    {
        ArgumentNullException.ThrowIfNull(names);

        return new ChartExCategories([names]);
    }
}

/// <summary>One series of a <see cref="ChartExSpace"/>: its values, its name and how it looks.</summary>
public sealed record ChartExSeries
{
    private readonly string? name;
    private readonly string? color;
    private readonly IReadOnlyList<int> subtotals = [];

    /// <summary>Creates a series.</summary>
    /// <param name="values">The numbers, one per category.</param>
    public ChartExSeries(ChartData values)
    {
        ArgumentNullException.ThrowIfNull(values);
        Values = values;
    }

    /// <summary>The numbers, one per category.</summary>
    public ChartData Values { get; init; }

    /// <summary>The name in the legend, or null for the application's.</summary>
    public string? Name
    {
        get => name;
        init
        {
            if (value is not null && value.Length > 255)
            {
                throw new ArgumentException("A series name holds at most 255 characters.", nameof(Name));
            }

            if (value is not null)
            {
                DrawingValues.XmlText(value, nameof(Name));
            }

            name = value;
        }
    }

    /// <summary>The reference the name is read from, such as <c>Sheet1!$B$1</c>, or null.</summary>
    public string? NameFormula { get; init; }

    /// <summary>The colour as <c>#RRGGBB</c>, or null for the application's own.</summary>
    public string? Color
    {
        get => color;
        init
        {
            string? normalized = null;

            if (value is not null)
            {
                normalized = DrawingValues.Color(value, nameof(Color));
            }

            color = normalized;
        }
    }

    /// <summary>The points of a waterfall that are totals, standing on the axis instead of following the last; 0-based.</summary>
    public IReadOnlyList<int> Subtotals
    {
        get => subtotals;
        init
        {
            ArgumentNullException.ThrowIfNull(value);

            if (value.Any(index => index < 0) == true)
            {
                throw new ArgumentOutOfRangeException(nameof(Subtotals), "A subtotal is the index of a point, 0 or more.");
            }

            subtotals = [.. value.Distinct().Order()];
        }
    }

    /// <summary>Labels on the points, or null for none.</summary>
    public DataLabels? DataLabels { get; init; }
}

/// <summary>
/// A chart part of the newer chart engine (cx:chartSpace in chartExN.xml): waterfall, treemap,
/// sunburst, histogram, Pareto, box and whisker, funnel and filled map. The data sits in the part
/// (cx:chartData), with the formulas it was read from; where the part is placed and what its
/// formulas point to is up to the package that owns it.
/// </summary>
public sealed record ChartExSpace
{
    private readonly IReadOnlyList<ChartExSeries> series = [];
    private readonly string? title;
    private readonly string? externalData;

    /// <summary>Creates a chart.</summary>
    /// <param name="type">The kind of chart.</param>
    public ChartExSpace(ChartExType type)
    {
        if (Enum.IsDefined(type) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(type), type, "Not a chart type.");
        }

        Type = type;
    }

    /// <summary>The kind of chart.</summary>
    public ChartExType Type { get; }

    /// <summary>The title above the chart, or null for none.</summary>
    public string? Title
    {
        get => title;
        init
        {
            if (value is not null && value.Length > 255)
            {
                throw new ArgumentException("A chart title holds at most 255 characters.", nameof(Title));
            }

            if (value is not null)
            {
                DrawingValues.XmlText(value, nameof(Title));
            }

            title = value;
        }
    }

    /// <summary>The categories the series share, or null for 1, 2, 3 … (a histogram needs none).</summary>
    public ChartExCategories? Categories { get; init; }

    /// <summary>The series: one for most types, up to 255 for a box and whisker chart.</summary>
    public IReadOnlyList<ChartExSeries> Series
    {
        get => series;
        init
        {
            ArgumentNullException.ThrowIfNull(value);

            if (value.Count > 255)
            {
                throw new ArgumentException("A chart holds at most 255 series.", nameof(Series));
            }

            foreach (ChartExSeries item in value)
            {
                ArgumentNullException.ThrowIfNull(item, nameof(Series));
            }

            series = [.. value];
        }
    }

    /// <summary>Where the legend goes, or null for the type's own: on top for a waterfall, treemap, box and whisker and Pareto, right of a map, none for the rest.</summary>
    public LegendPosition? Legend { get; init; }

    /// <summary>Where a treemap's parent labels go.</summary>
    public TreemapParentLabels ParentLabels { get; init; }

    /// <summary>Draw lines from each bar of a waterfall to the next.</summary>
    public bool ConnectorLines { get; init; } = true;

    /// <summary>The bins of a histogram or Pareto chart, or null for automatic ones.</summary>
    public ChartExBinning? Binning { get; init; }

    /// <summary>Count the quartiles of a box and whisker chart with the median included (inclusive) instead of left out (exclusive, the default).</summary>
    public bool InclusiveQuartiles { get; init; }

    /// <summary>Draw a line through the means of a box and whisker chart's boxes.</summary>
    public bool MeanLine { get; init; }

    /// <summary>Mark the mean of each box.</summary>
    public bool MeanMarkers { get; init; } = true;

    /// <summary>Show the points inside the whiskers.</summary>
    public bool InnerPoints { get; init; }

    /// <summary>Show the points outside the whiskers.</summary>
    public bool Outliers { get; init; } = true;

    /// <summary>Show the category axis (or the funnel's stage names), or null for the type's own.</summary>
    public bool? CategoryAxis { get; init; }

    /// <summary>Show the value axis, or null for the type's own.</summary>
    public bool? ValueAxis { get; init; }

    /// <summary>Draw the value axis' major gridlines, or null for the type's own.</summary>
    public bool? Gridlines { get; init; }

    /// <summary>Labels on every series' points, or null for the type's own. A series' own <see cref="ChartExSeries.DataLabels"/> wins.</summary>
    public DataLabels? DataLabels { get; init; }

    /// <summary>The relationship id of the workbook that holds the chart's data (cx:externalData), or null for none.</summary>
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

    /// <summary>True for the types with axes: waterfall, histogram, Pareto, box and whisker and funnel.</summary>
    public bool HasAxes => Type is ChartExType.Waterfall or ChartExType.Histogram or ChartExType.Pareto or ChartExType.BoxWhisker or ChartExType.Funnel;

    /// <summary>The legend the chart is written with: its own, or the type's.</summary>
    public LegendPosition LegendOrDefault => Legend ?? Type switch
    {
        ChartExType.Waterfall or ChartExType.Treemap or ChartExType.BoxWhisker or ChartExType.Pareto => LegendPosition.Top,
        ChartExType.RegionMap => LegendPosition.Right,
        _ => LegendPosition.None,
    };

    /// <summary>Throws ArgumentException for what the application would refuse: no series, more than one where the type takes one, a hierarchy where it takes a flat list, or subtotals outside a waterfall.</summary>
    public void Validate()
    {
        if (Series.Count == 0)
        {
            throw new ArgumentException("A chart needs a series.", nameof(Series));
        }

        if (Series.Count > 1 && Type != ChartExType.BoxWhisker)
        {
            throw new ArgumentException($"A {Type} chart has one series; only a box and whisker chart takes more.", nameof(Series));
        }

        if (Categories is not null && Categories.Levels.Count > 1 && (Type is ChartExType.Treemap or ChartExType.Sunburst or ChartExType.BoxWhisker) == false)
        {
            throw new ArgumentException($"A {Type} chart takes one level of categories; treemaps and sunbursts take more.", nameof(Categories));
        }

        if (Series.Any(item => item.Subtotals.Count > 0) == true && Type != ChartExType.Waterfall)
        {
            throw new ArgumentException("Only a waterfall has subtotals.", nameof(Series));
        }

        if (Type == ChartExType.RegionMap && Categories is null)
        {
            throw new ArgumentException("A filled map needs categories: the countries, regions or postal codes.", nameof(Categories));
        }
    }
}
