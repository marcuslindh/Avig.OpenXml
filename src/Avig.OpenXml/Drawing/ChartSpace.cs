namespace Avig.OpenXml.Drawing;

/// <summary>
/// The data of a series, as a chart part stores it: a <see cref="Formula"/> that points into a
/// workbook (the sheet itself in Excel, the embedded workbook in PowerPoint), the values last
/// read from it, or both. The chart part never says what the formula means; the package that
/// owns the chart does.
/// </summary>
public sealed record ChartData
{
    private readonly string? formula;
    private readonly IReadOnlyList<double?>? numbers;
    private readonly IReadOnlyList<string?>? texts;
    private readonly string? formatCode;

    /// <summary>The reference, such as <c>'Sales 2026'!$B$2:$B$13</c>, or null for values alone.</summary>
    public string? Formula
    {
        get => formula;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 8192))
            {
                throw new ArgumentException("A series formula holds 1 to 8192 characters.", nameof(Formula));
            }

            if (value is not null)
            {
                DrawingValues.XmlText(value, nameof(Formula));
            }

            formula = value;
        }
    }

    /// <summary>The numbers last read from the reference (c:numCache) or given alone (c:numLit); null for none. A null point is a gap.</summary>
    public IReadOnlyList<double?>? Numbers
    {
        get => numbers;
        init
        {
            if (value is not null)
            {
                foreach (double? number in value)
                {
                    if (number is double given && double.IsFinite(given) == false)
                    {
                        throw new ArgumentException("A chart value must be a finite number.", nameof(Numbers));
                    }
                }

                numbers = [.. value];

                return;
            }

            numbers = null;
        }
    }

    /// <summary>The texts last read from the reference (c:strCache) or given alone (c:strLit); null for none.</summary>
    public IReadOnlyList<string?>? Texts
    {
        get => texts;
        init
        {
            if (value is not null)
            {
                foreach (string? text in value)
                {
                    if (text is not null)
                    {
                        DrawingValues.XmlText(text, nameof(Texts));
                    }
                }

                texts = [.. value];

                return;
            }

            texts = null;
        }
    }

    /// <summary>The number format of <see cref="Numbers"/> (c:formatCode), or null for General.</summary>
    public string? FormatCode
    {
        get => formatCode;
        init
        {
            if (value is not null && value.Length > 255)
            {
                throw new ArgumentException("A number format holds at most 255 characters.", nameof(FormatCode));
            }

            formatCode = value;
        }
    }

    /// <summary>A reference without stored values: the application reads them when it opens the file, as Excel does.</summary>
    /// <param name="formula">The reference.</param>
    /// <returns>The data.</returns>
    public static ChartData Reference(string formula)
    {
        ArgumentNullException.ThrowIfNull(formula);

        return new ChartData { Formula = formula };
    }

    /// <summary>Numbers, with or without the reference they came from.</summary>
    /// <param name="numbers">The values; null is a gap.</param>
    /// <param name="formula">The reference, or null for values alone.</param>
    /// <returns>The data.</returns>
    public static ChartData FromNumbers(IReadOnlyList<double?> numbers, string? formula = null)
    {
        ArgumentNullException.ThrowIfNull(numbers);

        return new ChartData { Numbers = numbers, Formula = formula };
    }

    /// <summary>Texts, such as category names, with or without the reference they came from.</summary>
    /// <param name="texts">The texts.</param>
    /// <param name="formula">The reference, or null for texts alone.</param>
    /// <returns>The data.</returns>
    public static ChartData FromTexts(IReadOnlyList<string?> texts, string? formula = null)
    {
        ArgumentNullException.ThrowIfNull(texts);

        return new ChartData { Texts = texts, Formula = formula };
    }
}

/// <summary>One series of a <see cref="ChartSpace"/>: its data, its type and how it looks.</summary>
public sealed record ChartSpaceSeries
{
    private readonly string? name;
    private readonly string? color;
    private ChartType? type;
    private IReadOnlyList<string?> pointColors = [];
    private MarkerStyle? marker;
    private int? markerSize;

    /// <summary>Creates a series.</summary>
    /// <param name="values">The numbers (the y values of a scatter chart).</param>
    public ChartSpaceSeries(ChartData values)
    {
        ArgumentNullException.ThrowIfNull(values);
        Values = values;
    }

    /// <summary>The numbers (the y values of a scatter chart).</summary>
    public ChartData Values { get; init; }

    /// <summary>The labels along the axis (the x values of a scatter chart), or null for 1, 2, 3 ...</summary>
    public ChartData? Categories { get; init; }

    /// <summary>The name in the legend, or null for "Series 1" and so on. With <see cref="NameFormula"/>, the name last read from it.</summary>
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

    /// <summary>The reference the name is read from, such as <c>Sheet1!$B$1</c>, or null for a name given as text.</summary>
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

    /// <summary>
    /// Draw this series as another type than the chart's, for a combination chart: columns with
    /// a line, say. Column, line and area mix; null keeps the chart's type.
    /// </summary>
    public ChartType? Type
    {
        get => type;
        init
        {
            if (value is ChartType given && Enum.IsDefined(given) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(Type), given, "Not a chart type.");
            }

            type = value;
        }
    }

    /// <summary>Plot the series against a second value axis on the right, for numbers of another size.</summary>
    public bool SecondaryAxis { get; init; }

    /// <summary>The inside of the columns, bars, areas or slices; wins over <see cref="Color"/>. <see cref="Fill.None"/> hides them, as the first series of a funnel.</summary>
    public Fill? Fill { get; init; }

    /// <summary>The line of a line or scatter series, or the border of columns, bars, areas and slices; wins over <see cref="Color"/> for lines.</summary>
    public Outline? Outline { get; init; }

    /// <summary>A shadow behind the series, or null for none.</summary>
    public Shadow? Shadow { get; init; }

    /// <summary>
    /// A colour per point, as <c>#RRGGBB</c>, or null to leave that point as the series is: the
    /// slices of a pie, or one bar picked out among the rest.
    /// </summary>
    public IReadOnlyList<string?> PointColors
    {
        get => pointColors;
        init
        {
            ArgumentNullException.ThrowIfNull(value);
            List<string?> colors = [];

            foreach (string? given in value)
            {
                string? normalized = null;

                if (given is not null)
                {
                    normalized = DrawingValues.Color(given, nameof(PointColors));
                }

                colors.Add(normalized);
            }

            pointColors = colors;
        }
    }

    /// <summary>Curve the line through the points instead of straight segments. For line and scatter series.</summary>
    public bool Smooth { get; init; }

    /// <summary>The marker on each point of a line or scatter series, or null for the application's own.</summary>
    public MarkerStyle? Marker
    {
        get => marker;
        init
        {
            if (value is MarkerStyle given && Enum.IsDefined(given) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(Marker), given, "Not a marker style.");
            }

            marker = value;
        }
    }

    /// <summary>The marker's size in points, 2 to 72, or null for the default 5.</summary>
    public int? MarkerSize
    {
        get => markerSize;
        init
        {
            if (value is int size)
            {
                DrawingValues.Range(size, 2, 72, nameof(MarkerSize), "A marker size");
            }

            markerSize = value;
        }
    }

    /// <summary>Labels on this series' points, or null for the chart's <see cref="ChartSpace.DataLabels"/>.</summary>
    public DataLabels? DataLabels { get; init; }
}

/// <summary>
/// A chart part (c:chartSpace): the type, the series, the axes, the legend and how it all looks.
/// Where it is placed and what its formulas point to is up to the package that owns it; Avig.Kalkyl's
/// <c>Chart</c> builds one over cells in the workbook.
/// </summary>
public sealed record ChartSpace
{
    private readonly IReadOnlyList<ChartSpaceSeries> series = [];
    private readonly string? title;
    private int? holeSize;
    private int firstSliceAngle;
    private int? gapWidth;
    private int? overlap;

    /// <summary>Creates a chart.</summary>
    /// <param name="type">The kind of chart.</param>
    public ChartSpace(ChartType type)
    {
        if (Enum.IsDefined(type) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(type), type, "Not a chart type.");
        }

        Type = type;
    }

    /// <summary>The kind of chart.</summary>
    public ChartType Type { get; }

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

    /// <summary>The series, at most 255.</summary>
    public IReadOnlyList<ChartSpaceSeries> Series
    {
        get => series;
        init
        {
            ArgumentNullException.ThrowIfNull(value);

            if (value.Count > 255)
            {
                throw new ArgumentException("A chart holds at most 255 series.", nameof(Series));
            }

            foreach (ChartSpaceSeries item in value)
            {
                ArgumentNullException.ThrowIfNull(item, nameof(Series));
            }

            series = [.. value];
        }
    }

    /// <summary>Where the legend goes.</summary>
    public LegendPosition Legend { get; init; } = LegendPosition.Right;

    /// <summary>
    /// Stack the series on top of each other instead of side by side. For column, bar, line and
    /// area charts.
    /// </summary>
    public bool Stacked { get; init; }

    /// <summary>The title under the category axis (the x axis of a scatter chart), or null.</summary>
    public string? CategoryAxisTitle { get; init; }

    /// <summary>The title beside the value axis, or null.</summary>
    public string? ValueAxisTitle { get; init; }

    /// <summary>The size of a doughnut's hole in percent of its diameter, 10 to 90, or null for the default 50.</summary>
    public int? HoleSize
    {
        get => holeSize;
        init
        {
            if (value is int size)
            {
                DrawingValues.Range(size, 10, 90, nameof(HoleSize), "A hole size");
            }

            holeSize = value;
        }
    }

    /// <summary>Where a pie's or doughnut's first slice starts, in degrees clockwise from the top, 0 to 360.</summary>
    public int FirstSliceAngle
    {
        get => firstSliceAngle;
        init => firstSliceAngle = DrawingValues.Range(value, 0, 360, nameof(FirstSliceAngle), "The first slice's angle");
    }

    /// <summary>The space between groups of columns or bars in percent of a column's width, 0 to 500, or null for the default 150.</summary>
    public int? GapWidth
    {
        get => gapWidth;
        init
        {
            if (value is int gap)
            {
                DrawingValues.Range(gap, 0, 500, nameof(GapWidth), "A gap width");
            }

            gapWidth = value;
        }
    }

    /// <summary>
    /// How much the columns or bars of a group overlap in percent, -100 to 100, or null for the
    /// default 0 (100 when stacked). 100 draws one series on top of another, as a progress bar
    /// over its background.
    /// </summary>
    public int? Overlap
    {
        get => overlap;
        init
        {
            if (value is int given)
            {
                DrawingValues.Range(given, -100, 100, nameof(Overlap), "An overlap");
            }

            overlap = value;
        }
    }

    /// <summary>
    /// Give each point its own colour, or null for the default: on for pie and doughnut
    /// charts, off for the rest. Off, a doughnut's slices share the series' colour.
    /// </summary>
    public bool? VaryColors { get; init; }

    /// <summary>Round the corners of the chart's border.</summary>
    public bool RoundedCorners { get; init; }

    /// <summary>The chart area behind everything, or null for white; <see cref="Fill.None"/> lets what is behind show through.</summary>
    public Fill? Background { get; init; }

    /// <summary>The plot area inside the axes, or null for white; see-through instead when <see cref="Background"/> is set.</summary>
    public Fill? PlotBackground { get; init; }

    /// <summary>The chart's border, or null for a thin grey one; <see cref="Outline.None"/> removes it.</summary>
    public Outline? Border { get; init; }

    /// <summary>The font of all the chart's text, or null for the default; the title, legend, axes and labels can override it.</summary>
    public TextFont? Font { get; init; }

    /// <summary>The title's font, or null for the chart's.</summary>
    public TextFont? TitleFont { get; init; }

    /// <summary>The legend's font, or null for the chart's.</summary>
    public TextFont? LegendFont { get; init; }

    /// <summary>The category axis (the x axis of a scatter chart), or null for the default.</summary>
    public ChartAxis? CategoryAxis { get; init; }

    /// <summary>The value axis, or null for the default.</summary>
    public ChartAxis? ValueAxis { get; init; }

    /// <summary>The second value axis, on the right, of the series with <see cref="ChartSpaceSeries.SecondaryAxis"/>, or null for the default.</summary>
    public ChartAxis? SecondaryValueAxis { get; init; }

    /// <summary>Labels on every series' points, or null for none. A series' own <see cref="ChartSpaceSeries.DataLabels"/> wins.</summary>
    public DataLabels? DataLabels { get; init; }

    /// <summary>The type a series is drawn as: its own, or the chart's.</summary>
    /// <param name="item">A series of the chart.</param>
    /// <returns>The type.</returns>
    public ChartType TypeOf(ChartSpaceSeries item)
    {
        ArgumentNullException.ThrowIfNull(item);

        if (item.Type is ChartType own)
        {
            return own;
        }

        return Type;
    }

    /// <summary>
    /// Throws ArgumentException for what Office would repair: a mix of types it does not
    /// combine, a label position the type does not take, an axis whose minimum is not below its
    /// maximum, or a picture fill, which a chart part cannot refer to.
    /// </summary>
    public void Validate()
    {
        foreach (ChartAxis? axis in new[] { CategoryAxis, ValueAxis, SecondaryValueAxis })
        {
            if (axis is not null && axis.Minimum is double minimum && axis.Maximum is double maximum && minimum >= maximum)
            {
                throw new ArgumentException($"An axis' minimum ({minimum}) must be below its maximum ({maximum}).", nameof(ValueAxis));
            }
        }

        foreach (Fill? fill in Series.Select(item => item.Fill).Append(Background).Append(PlotBackground))
        {
            if (fill is not null && fill.Kind == FillKind.Picture)
            {
                throw new ArgumentException("A chart cannot take a picture fill.", nameof(Series));
            }
        }

        bool mixed = Series.Any(item => TypeOf(item) != Type || item.SecondaryAxis == true);

        if (mixed == true)
        {
            foreach (ChartSpaceSeries item in Series)
            {
                if ((TypeOf(item) is ChartType.Column or ChartType.Line or ChartType.Area) == false)
                {
                    throw new ArgumentException($"A combination chart mixes column, line and area series; {TypeOf(item)} cannot be one of them.", nameof(Series));
                }
            }

            if (Series.All(item => item.SecondaryAxis == true) == true)
            {
                throw new ArgumentException("At least one series must stay on the primary axis.", nameof(Series));
            }
        }

        foreach (ChartSpaceSeries item in Series)
        {
            DataLabels? labels = item.DataLabels ?? DataLabels;

            if (labels is null || labels.Position == DataLabelPosition.Default)
            {
                continue;
            }

            ChartType type = TypeOf(item);
            bool stacked = Stacked == true && type == Type;
            bool allowed = type switch
            {
                ChartType.Column or ChartType.Bar => labels.Position is DataLabelPosition.Center or DataLabelPosition.InsideEnd or DataLabelPosition.InsideBase
                    || (labels.Position == DataLabelPosition.OutsideEnd && stacked == false),
                ChartType.Line or ChartType.Scatter => labels.Position is DataLabelPosition.Center or DataLabelPosition.Above or DataLabelPosition.Below or DataLabelPosition.Left or DataLabelPosition.Right,
                ChartType.Pie => labels.Position is DataLabelPosition.Center or DataLabelPosition.InsideEnd or DataLabelPosition.OutsideEnd or DataLabelPosition.BestFit,
                _ => false,
            };

            if (allowed == false)
            {
                throw new ArgumentException($"A {type} series does not take data labels at {labels.Position}; Excel would repair the file. See DataLabels.Position.", nameof(DataLabels));
            }
        }
    }
}
