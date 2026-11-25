using System.Xml.Linq;

namespace Avig.OpenXml.Drawing;

/// <summary>
/// Reads a chart part (c:chartSpace) back into a <see cref="ChartSpace"/>: the types, the series
/// with their formulas and stored values, the axes, the legend and how it all looks. Elements are
/// read by local name. What the types cannot hold (3-D charts, trendlines, error bars, theme
/// colours) is left out; values outside what the types take are clamped.
/// </summary>
public static class ChartReader
{
    /// <summary>The most points read from one cache, so that a hostile ptCount cannot make a list without end.</summary>
    private const int MaxPoints = 1_048_576;

    /// <summary>Reads a chart part.</summary>
    /// <param name="chartSpace">The root element, c:chartSpace.</param>
    /// <returns>The chart; Invalid when the element is not a chartSpace or holds no plot this layer can read.</returns>
    public static Result<ChartSpace> Read(XElement chartSpace)
    {
        ArgumentNullException.ThrowIfNull(chartSpace);

        if (chartSpace.Name.LocalName != "chartSpace")
        {
            return Result<ChartSpace>.Invalid($"<{chartSpace.Name.LocalName}> is not a chart part's root, c:chartSpace.");
        }

        XElement? chart = Child(chartSpace, "chart");
        XElement? plotArea = Child(chart, "plotArea");
        List<(ChartType Type, XElement Element)> plots = [];

        foreach (XElement plot in plotArea?.Elements() ?? [])
        {
            if (TryPlotType(plot, out ChartType type) == true)
            {
                plots.Add((type, plot));
            }
        }

        if (plots.Count == 0)
        {
            return Result<ChartSpace>.Invalid($"{OpcErrors.DamagedPrefix}: the chart has no plot of a type that can be read.");
        }

        // The chart's own type: a stacked plot's, or else that of the plot the first series is in.
        (ChartType Type, XElement Element) main = plots.Find(plot => Value(plot.Element, "grouping") is "stacked" or "percentStacked");

        if (main.Element is null)
        {
            main = plots.OrderBy(plot => Children(plot.Element, "ser").Select(series => Int(Value(series, "idx"), int.MaxValue)).DefaultIfEmpty(int.MaxValue).Min()).First();
        }

        Dictionary<string, XElement> axes = new(StringComparer.Ordinal);

        foreach (XElement axis in plotArea!.Elements().Where(element => element.Name.LocalName is "catAx" or "valAx" or "dateAx" or "serAx"))
        {
            if (Value(axis, "axId") is string id)
            {
                axes.TryAdd(id, axis);
            }
        }

        List<string> mainAxes = Children(main.Element, "axId").Select(element => element.Attribute("val")?.Value ?? string.Empty).ToList();

        // A stock chart with the volume as columns is a stock chart, though the columns come first.
        ChartType chartType = main.Type;

        if (plots.Exists(plot => plot.Type == ChartType.Stock) == true)
        {
            chartType = ChartType.Stock;
        }

        List<(int Index, ChartSpaceSeries Series)> series = [];

        foreach ((ChartType type, XElement plot) in plots)
        {
            List<string> plotAxes = Children(plot, "axId").Select(element => element.Attribute("val")?.Value ?? string.Empty).ToList();
            bool secondary = plotAxes.Count >= 2 && mainAxes.Count >= 2 && plotAxes[1] != mainAxes[1];

            foreach (XElement element in Children(plot, "ser"))
            {
                if (series.Count >= 255)
                {
                    break;
                }

                ChartSpaceSeries read = ReadSeries(element, type, chartType) with { SecondaryAxis = secondary };

                series.Add((Int(Value(element, "idx"), series.Count), read));
            }
        }

        ChartSpace space = new(chartType)
        {
            Title = TitleText(Child(chart, "title")),
            TitleFont = TitleFont(Child(chart, "title")),
            Series = series.OrderBy(item => item.Index).Select(item => item.Series).ToList(),
            Grouping = GroupingOf(main.Element, axes),
            VaryColors = Flag(main.Element, "varyColors"),
            RoundedCorners = Flag(chartSpace, "roundedCorners") ?? true,
            HoleSize = Clamped(Value(main.Element, "holeSize"), 10, 90),
            FirstSliceAngle = Clamped(Value(main.Element, "firstSliceAng"), 0, 360) ?? 0,
            GapWidth = Clamped(Value(main.Element, "gapWidth"), 0, 500),
            Overlap = Clamped(Value(main.Element, "overlap"), -100, 100),
        };

        space = ReadKinds(space, chartSpace, chart, main, plots, axes);
        space = ReadLegend(space, Child(chart, "legend"));
        space = ReadAxes(space, main, mainAxes, axes, plots);
        XElement? chartProperties = Child(chartSpace, "spPr");

        if (Child(plotArea, "spPr") is XElement plotProperties && DrawingReader.TryReadFill(plotProperties, out Fill? plotFill) == true)
        {
            space = space with { PlotBackground = plotFill };
        }

        if (chartProperties is not null && DrawingReader.TryReadFill(chartProperties, out Fill? background) == true)
        {
            space = space with { Background = background };
        }

        if (chartProperties is not null && DrawingReader.TryReadOutline(chartProperties, out Outline? border) == true)
        {
            space = space with { Border = border };
        }

        return space with { Font = TextPropertiesFont(chartSpace) };
    }

    private static bool TryPlotType(XElement plot, out ChartType type)
    {
        type = ChartType.Column;

        switch (plot.Name.LocalName)
        {
            case "barChart" or "bar3DChart":
                type = ChartType.Column;

                if (Value(plot, "barDir") == "bar")
                {
                    type = ChartType.Bar;
                }

                return true;

            case "lineChart" or "line3DChart":
                type = ChartType.Line;

                return true;

            case "areaChart" or "area3DChart":
                type = ChartType.Area;

                return true;

            case "pieChart" or "pie3DChart":
                type = ChartType.Pie;

                return true;

            case "doughnutChart":
                type = ChartType.Doughnut;

                return true;

            case "scatterChart":
                type = ChartType.Scatter;

                return true;

            case "ofPieChart":
                type = ChartType.PieOfPie;

                if (Value(plot, "ofPieType") == "bar")
                {
                    type = ChartType.BarOfPie;
                }

                return true;

            case "radarChart":
                type = ChartType.Radar;

                return true;

            case "bubbleChart":
                type = ChartType.Bubble;

                return true;

            case "stockChart":
                type = ChartType.Stock;

                return true;

            case "surfaceChart" or "surface3DChart":
                type = ChartType.Surface;

                return true;

            default:
                return false;
        }
    }

    /// <summary>The grouping of a plot: stacked, stacked to 100 %, or (in a 3-D column, bar or area plot with a depth axis) in depth.</summary>
    private static ChartGrouping GroupingOf(XElement plot, Dictionary<string, XElement> axes)
    {
        switch (Value(plot, "grouping"))
        {
            case "stacked":
                return ChartGrouping.Stacked;

            case "percentStacked":
                return ChartGrouping.PercentStacked;

            case "standard" when plot.Name.LocalName == "line3DChart":
                return ChartGrouping.Depth;

            case "standard" when plot.Name.LocalName is "bar3DChart" or "area3DChart":
                List<string> ids = Children(plot, "axId").Select(element => element.Attribute("val")?.Value ?? string.Empty).ToList();

                if (ids.Count >= 3 && axes.TryGetValue(ids[2], out XElement? depth) == true && depth.Name.LocalName == "serAx")
                {
                    return ChartGrouping.Depth;
                }

                return ChartGrouping.Default;

            default:
                return ChartGrouping.Default;
        }
    }

    /// <summary>What the other chart types add: 3-D, shapes, radar, surface, bubble, pie-of-pie, stock lines, the depth axis, the data table, the style and the workbook.</summary>
    private static ChartSpace ReadKinds(ChartSpace space, XElement chartSpace, XElement? chart, (ChartType Type, XElement Element) main, List<(ChartType Type, XElement Element)> plots, Dictionary<string, XElement> axes)
    {
        XElement plot = main.Element;

        if (plot.Name.LocalName.EndsWith("3DChart", StringComparison.Ordinal) == true)
        {
            space = space with { ThreeD = ReadView3D(Child(chart, "view3D")) };
        }

        if (plot.Name.LocalName == "bar3DChart")
        {
            space = space with { BarShape = BarShapeOf(Value(plot, "shape")) };
        }

        space = space with
        {
            RadarStyle = RadarStyleOf(Value(plot, "radarStyle")),
            Wireframe = Flag(plot, "wireframe") ?? false,
            ShowNegativeBubbles = Flag(plot, "showNegBubbles") ?? false,
            BubbleScale = Clamped(Value(plot, "bubbleScale"), 0, 300),
            Bubble3D = Flag(plot, "bubble3D") ?? Children(plot, "ser").Any(item => Flag(item, "bubble3D") == true),
            SecondPlotSize = Clamped(Value(plot, "secondPieSize"), 5, 200),
            Split = SplitOf(Value(plot, "splitType")),
        };

        if (Number(Value(plot, "splitPos")) is double split && split >= 0)
        {
            space = space with { SplitValue = split };
        }

        // The stock lines are read so that the defaults (a stock plot's high-low lines, its
        // up-down bars with four series, a pie-of-pie's series lines) stay null.
        foreach ((ChartType type, XElement element) in plots)
        {
            bool stock = type == ChartType.Stock;

            if (type is ChartType.Line or ChartType.Stock)
            {
                bool highLow = Child(element, "hiLowLines") is not null;
                bool upDown = Child(element, "upDownBars") is not null;

                if (highLow != stock)
                {
                    space = space with { HighLowLines = highLow };
                }

                if (upDown != (stock == true && Children(element, "ser").Count() == 4))
                {
                    space = space with { UpDownBars = upDown };
                }
            }

            if (Child(element, "dropLines") is not null)
            {
                space = space with { DropLines = true };
            }

            if (type is ChartType.PieOfPie or ChartType.BarOfPie && Child(element, "serLines") is null)
            {
                space = space with { SeriesLines = false };
            }

            if (type is ChartType.Column or ChartType.Bar && Child(element, "serLines") is not null)
            {
                space = space with { SeriesLines = true };
            }

            if (type == ChartType.Scatter && Children(element, "ser").Any(AutomaticLine) == true)
            {
                space = space with { ScatterLines = true };
            }
        }

        XElement? plotArea = Child(chart, "plotArea");

        if (Child(plotArea, "serAx") is XElement seriesAxis)
        {
            space = space with { SeriesAxis = ReadAxis(seriesAxis) };
        }

        if (Child(plotArea, "dTable") is XElement table)
        {
            space = space with
            {
                DataTable = new ChartDataTable
                {
                    HorizontalBorders = Flag(table, "showHorzBorder") ?? false,
                    VerticalBorders = Flag(table, "showVertBorder") ?? false,
                    Outline = Flag(table, "showOutline") ?? false,
                    ShowLegendKeys = Flag(table, "showKeys") ?? false,
                },
            };
        }

        if (StyleOf(chartSpace) is int style)
        {
            space = space with { Style = style };
        }

        XAttribute? data = Child(chartSpace, "externalData")?.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "id");

        if (data is not null && data.Value.Length > 0 && data.Value.Length <= 255)
        {
            space = space with { ExternalData = Clean(data.Value) };
        }

        return space;
    }

    /// <summary>A scatter series with a line in the application's colour: an a:ln that neither hides the line nor gives it a colour this layer reads.</summary>
    private static bool AutomaticLine(XElement series)
    {
        XElement? properties = Child(series, "spPr");
        XElement? line = Child(properties, "ln");

        return line is not null && Child(line, "noFill") is null && DrawingReader.TryReadOutline(properties!, out _) == false;
    }

    /// <summary>The chart's built-in style: c:style, or c14:style (100 more) where the mc:AlternateContent was resolved to it.</summary>
    private static int? StyleOf(XElement chartSpace)
    {
        XElement? style = chartSpace.Elements().FirstOrDefault(element => element.Name.LocalName == "style");

        if (style is null)
        {
            XElement? fallback = Child(Child(chartSpace, "AlternateContent"), "Fallback");
            style = Child(fallback, "style");
        }

        int? read = null;

        if (style is not null && int.TryParse(style.Attribute("val")?.Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int value) == true)
        {
            if (value > 100)
            {
                value -= 100;
            }

            if (value >= 1 && value <= 48)
            {
                read = value;
            }
        }

        return read;
    }

    private static ChartView3D ReadView3D(XElement? view)
    {
        ChartView3D read = new()
        {
            RotationX = Clamped(Value(view, "rotX"), -90, 90) ?? 15,
            RotationY = Clamped(Value(view, "rotY"), 0, 360) ?? 20,
            DepthPercent = Clamped(Value(view, "depthPercent"), 20, 2000) ?? 100,
            RightAngleAxes = Flag(view, "rAngAx") ?? true,
            Perspective = Clamped(Value(view, "perspective"), 0, 240) ?? 30,
        };

        return read;
    }

    private static BarShape BarShapeOf(string? name)
    {
        return name switch
        {
            "cylinder" => BarShape.Cylinder,
            "cone" => BarShape.Cone,
            "coneToMax" => BarShape.ConeToMax,
            "pyramid" => BarShape.Pyramid,
            "pyramidToMax" => BarShape.PyramidToMax,
            _ => BarShape.Box,
        };
    }

    private static RadarStyle RadarStyleOf(string? name)
    {
        if (name == "filled")
        {
            return RadarStyle.Filled;
        }

        return RadarStyle.Marker;
    }

    private static ChartSplit SplitOf(string? name)
    {
        return name switch
        {
            "pos" => ChartSplit.Position,
            "val" => ChartSplit.Value,
            "percent" => ChartSplit.Percent,
            _ => ChartSplit.Auto,
        };
    }

    private static Trendline ReadTrendline(XElement element)
    {
        TrendlineKind kind = Value(element, "trendlineType") switch
        {
            "exp" => TrendlineKind.Exponential,
            "log" => TrendlineKind.Logarithmic,
            "movingAvg" => TrendlineKind.MovingAverage,
            "poly" => TrendlineKind.Polynomial,
            "power" => TrendlineKind.Power,
            _ => TrendlineKind.Linear,
        };

        Trendline read = new()
        {
            Kind = kind,
            Order = Clamped(Value(element, "order"), 2, 6) ?? 2,
            Period = Clamped(Value(element, "period"), 2, 255) ?? 2,
            Name = Limit(Child(element, "name")?.Value, 255),
            Forward = Math.Max(0, Number(Value(element, "forward")) ?? 0),
            Backward = Math.Max(0, Number(Value(element, "backward")) ?? 0),
            Intercept = Number(Value(element, "intercept")),
            DisplayEquation = Flag(element, "dispEq") ?? false,
            DisplayRSquared = Flag(element, "dispRSqr") ?? false,
        };

        if (Child(element, "spPr") is XElement properties && DrawingReader.TryReadOutline(properties, out Outline? line) == true)
        {
            read = read with { Outline = line };
        }

        return read;
    }

    private static ErrorBars ReadErrorBars(XElement element)
    {
        ErrorBars read = new()
        {
            Direction = Value(element, "errDir") switch
            {
                "x" => ErrorBarDirection.X,
                _ => ErrorBarDirection.Y,
            },
            Side = Value(element, "errBarType") switch
            {
                "plus" => ErrorBarSide.Plus,
                "minus" => ErrorBarSide.Minus,
                _ => ErrorBarSide.Both,
            },
            Amount = Value(element, "errValType") switch
            {
                "percentage" => ErrorBarAmount.Percentage,
                "stdDev" => ErrorBarAmount.StandardDeviation,
                "stdErr" => ErrorBarAmount.StandardError,
                _ => ErrorBarAmount.Fixed,
            },
            Value = Math.Max(0, Number(Value(element, "val")) ?? 1),
            NoEndCap = Flag(element, "noEndCap") ?? false,
        };

        if (Child(element, "spPr") is XElement properties && DrawingReader.TryReadOutline(properties, out Outline? line) == true)
        {
            read = read with { Outline = line };
        }

        return read;
    }

    private static ChartSpaceSeries ReadSeries(XElement element, ChartType type, ChartType chartType)
    {
        string valueName = "val";
        string categoryName = "cat";

        if (type is ChartType.Scatter or ChartType.Bubble)
        {
            valueName = "yVal";
            categoryName = "xVal";
        }

        ChartData values = Data(Child(element, valueName));
        ChartData? categories = null;

        if (Child(element, categoryName) is XElement categoryElement)
        {
            categories = Data(categoryElement);
        }

        ChartSpaceSeries series = new(values) { Categories = categories };

        XElement? text = Child(element, "tx");

        if (Child(text, "strRef") is XElement nameReference)
        {
            series = series with
            {
                NameFormula = Child(nameReference, "f")?.Value,
                Name = Limit(Children(Child(nameReference, "strCache"), "pt").Select(point => Child(point, "v")?.Value).FirstOrDefault(), 255),
            };
        }
        else if (Child(text, "v") is XElement nameValue)
        {
            series = series with { Name = Limit(nameValue.Value, 255) };
        }

        if (type != chartType)
        {
            series = series with { Type = type };
        }

        if (Child(element, "spPr") is XElement properties)
        {
            if (DrawingReader.TryReadFill(properties, out Fill? fill) == true)
            {
                series = series with { Fill = fill };
            }

            if (DrawingReader.TryReadOutline(properties, out Outline? outline) == true)
            {
                series = series with { Outline = outline };
            }

            DrawingReader.ReadEffects(Child(properties, "effectLst"), out Shadow? shadow, out _, out _);
            series = series with { Shadow = shadow };
        }

        if (Child(element, "marker") is XElement marker)
        {
            series = series with
            {
                Marker = DrawingNames.MarkerOf(Value(marker, "symbol")),
                MarkerSize = Clamped(Value(marker, "size"), 2, 72),
            };

            // The writer paints the markers in the series' colour, which is all that is left of it
            // when the line itself is hidden, as in a scatter chart without lines.
            if (Child(marker, "spPr") is XElement markerProperties && DrawingReader.TryReadFill(markerProperties, out Fill? markerFill) == true && markerFill.Kind == FillKind.Solid)
            {
                series = series with { Color = markerFill.Color };
            }
        }

        List<string?> pointColors = [];

        foreach (XElement point in Children(element, "dPt"))
        {
            int index = Int(Value(point, "idx"), -1);

            if (index < 0 || index >= 4096)
            {
                continue;
            }

            while (pointColors.Count <= index)
            {
                pointColors.Add(null);
            }

            XElement? pointProperties = Child(point, "spPr");

            if (pointProperties is not null && DrawingReader.TryReadFill(pointProperties, out Fill? pointFill) == true && pointFill.Kind == FillKind.Solid)
            {
                pointColors[index] = pointFill.Color;
            }
            else if (pointProperties is not null && DrawingReader.TryReadOutline(pointProperties, out Outline? pointLine) == true && pointLine.Color is not null)
            {
                pointColors[index] = pointLine.Color;
            }
        }

        if (Child(element, "dLbls") is XElement labels)
        {
            series = series with { DataLabels = ReadLabels(labels) };
        }

        if (Child(element, "bubbleSize") is XElement sizes)
        {
            series = series with { BubbleSizes = Data(sizes) };
        }

        series = series with
        {
            Explosion = Clamped(Value(element, "explosion"), 0, 400),
            Trendlines = Children(element, "trendline").Take(8).Select(ReadTrendline).ToList(),
            ErrorBars = Children(element, "errBars").Select(ReadErrorBars).DistinctBy(bars => bars.Direction).ToList(),
        };

        return series with
        {
            PointColors = pointColors,
            Smooth = Flag(element, "smooth") ?? false,
        };
    }

    private static DataLabels ReadLabels(XElement labels)
    {
        DataLabels read = new()
        {
            ShowValue = Flag(labels, "showVal") ?? false,
            ShowPercent = Flag(labels, "showPercent") ?? false,
            ShowCategory = Flag(labels, "showCatName") ?? false,
            ShowSeriesName = Flag(labels, "showSerName") ?? false,
            Position = DrawingNames.PositionOf(Value(labels, "dLblPos")),
            Font = TextPropertiesFont(labels),
        };

        if (Child(labels, "numFmt")?.Attribute("formatCode")?.Value is string format && format.Length > 0 && format.Length <= 255)
        {
            read = read with { NumberFormat = format };
        }

        return read;
    }

    private static ChartSpace ReadLegend(ChartSpace space, XElement? legend)
    {
        if (legend is null)
        {
            return space with { Legend = LegendPosition.None };
        }

        return space with
        {
            Legend = DrawingNames.LegendOf(Value(legend, "legendPos")),
            LegendFont = TextPropertiesFont(legend),
        };
    }

    private static ChartSpace ReadAxes(ChartSpace space, (ChartType Type, XElement Element) main, List<string> mainAxes, Dictionary<string, XElement> axes, List<(ChartType Type, XElement Element)> plots)
    {
        if (space.HasAxes == false || mainAxes.Count < 2)
        {
            return space;
        }

        if (axes.TryGetValue(mainAxes[0], out XElement? category) == true)
        {
            space = space with
            {
                CategoryAxis = ReadAxis(category),
                CategoryAxisTitle = TitleText(Child(category, "title")),
            };
        }

        if (axes.TryGetValue(mainAxes[1], out XElement? value) == true)
        {
            space = space with
            {
                ValueAxis = ReadAxis(value),
                ValueAxisTitle = TitleText(Child(value, "title")),
            };
        }

        foreach ((ChartType _, XElement plot) in plots)
        {
            List<string> plotAxes = Children(plot, "axId").Select(element => element.Attribute("val")?.Value ?? string.Empty).ToList();

            if (plotAxes.Count >= 2 && plotAxes[1] != mainAxes[1] && axes.TryGetValue(plotAxes[1], out XElement? secondary) == true)
            {
                return space with { SecondaryValueAxis = ReadAxis(secondary) };
            }
        }

        return space;
    }

    private static ChartAxis ReadAxis(XElement axis)
    {
        XElement? scaling = Child(axis, "scaling");
        double? maximum = Number(Value(scaling, "max"));
        double? minimum = Number(Value(scaling, "min"));

        if (minimum is double low && maximum is double high && low >= high)
        {
            minimum = null;
            maximum = null;
        }

        ChartAxis read = new()
        {
            Visible = (Flag(axis, "delete") ?? false) == false,
            Gridlines = Child(axis, "majorGridlines") is not null,
            Minimum = minimum,
            Maximum = maximum,
            Font = TextPropertiesFont(axis),
        };

        if (Number(Value(axis, "majorUnit")) is double unit && unit > 0)
        {
            read = read with { MajorUnit = unit };
        }

        if (Number(Value(axis, "minorUnit")) is double minor && minor > 0)
        {
            read = read with { MinorUnit = minor };
        }

        if (Number(Value(scaling, "logBase")) is double logBase && logBase >= 2 && logBase <= 1000)
        {
            read = read with { LogBase = logBase };
        }

        read = read with
        {
            Reversed = Value(scaling, "orientation") == "maxMin",
            MinorGridlines = Child(axis, "minorGridlines") is not null,
        };

        if (Child(Child(axis, "minorGridlines"), "spPr") is XElement minorProperties && DrawingReader.TryReadOutline(minorProperties, out Outline? minorLine) == true)
        {
            read = read with { MinorGridlineOutline = minorLine };
        }

        if (Child(Child(axis, "majorGridlines"), "spPr") is XElement gridProperties && DrawingReader.TryReadOutline(gridProperties, out Outline? gridline) == true)
        {
            read = read with { GridlineOutline = gridline };
        }

        if (Child(axis, "spPr") is XElement properties && DrawingReader.TryReadOutline(properties, out Outline? line) == true)
        {
            read = read with { Line = line };
        }

        XElement? numberFormat = Child(axis, "numFmt");

        if (numberFormat is not null && (Bool(numberFormat.Attribute("sourceLinked")?.Value) ?? false) == false
            && numberFormat.Attribute("formatCode")?.Value is string format && format.Length > 0 && format.Length <= 255)
        {
            read = read with { NumberFormat = format };
        }

        return read;
    }

    /// <summary>The text of a c:title's rich text, its runs joined, or null.</summary>
    private static string? TitleText(XElement? title)
    {
        XElement? rich = Child(Child(title, "tx"), "rich");
        string? text = null;

        if (rich is not null)
        {
            text = Limit(string.Concat(rich.Descendants().Where(element => element.Name.LocalName == "t").Select(element => element.Value)), 255);
        }

        return text;
    }

    private static TextFont? TitleFont(XElement? title)
    {
        XElement? run = Child(Child(Child(Child(title, "tx"), "rich"), "p"), "r");
        TextFont? font = null;

        if (Child(run, "rPr") is XElement properties && DrawingReader.TryReadFont(properties, out TextFont? read) == true)
        {
            font = read;
        }

        return font;
    }

    /// <summary>The look a c:txPr gives the text of the element it is in, or null.</summary>
    private static TextFont? TextPropertiesFont(XElement? element)
    {
        XElement? properties = Child(Child(Child(Child(element, "txPr"), "p"), "pPr"), "defRPr");
        TextFont? font = null;

        if (properties is not null && DrawingReader.TryReadFont(properties, out TextFont? read) == true)
        {
            font = read;
        }

        return font;
    }

    /// <summary>A cat, val, xVal or yVal: the formula and the stored values, of a reference or a literal.</summary>
    private static ChartData Data(XElement? element)
    {
        XElement? source = element?.Elements().FirstOrDefault(child => child.Name.LocalName is "numRef" or "strRef" or "numLit" or "strLit" or "multiLvlStrRef");
        ChartData data = new();

        if (source is null)
        {
            return data;
        }

        string? formula = Child(source, "f")?.Value;

        if (formula is not null && formula.Length > 0 && formula.Length <= 8192)
        {
            data = data with { Formula = Clean(formula) };
        }

        bool numbers = source.Name.LocalName is "numRef" or "numLit";
        XElement? points = source.Name.LocalName switch
        {
            "numRef" => Child(source, "numCache"),
            "strRef" => Child(source, "strCache"),
            "multiLvlStrRef" => null,
            _ => source,
        };

        if (points is null)
        {
            return data;
        }

        int count = Math.Clamp(Int(Value(points, "ptCount"), 0), 0, MaxPoints);

        foreach (XElement point in Children(points, "pt"))
        {
            int index = Int(point.Attribute("idx")?.Value, -1);

            if (index >= count && index < MaxPoints)
            {
                count = index + 1;
            }
        }

        if (numbers == true)
        {
            double?[] values = new double?[count];

            foreach (XElement point in Children(points, "pt"))
            {
                int index = Int(point.Attribute("idx")?.Value, -1);

                if (index >= 0 && index < count && Number(Child(point, "v")?.Value) is double number)
                {
                    values[index] = number;
                }
            }

            data = data with { Numbers = values };

            if (Child(points, "formatCode")?.Value is string format && format != "General" && format.Length <= 255)
            {
                data = data with { FormatCode = format };
            }

            return data;
        }

        string?[] texts = new string?[count];

        foreach (XElement point in Children(points, "pt"))
        {
            int index = Int(point.Attribute("idx")?.Value, -1);

            if (index >= 0 && index < count)
            {
                texts[index] = Clean(Child(point, "v")?.Value ?? string.Empty);
            }
        }

        return data with { Texts = texts };
    }

    private static XElement? Child(XElement? parent, string localName)
    {
        return parent?.Elements().FirstOrDefault(child => child.Name.LocalName == localName);
    }

    private static IEnumerable<XElement> Children(XElement? parent, string localName)
    {
        if (parent is null)
        {
            return [];
        }

        return parent.Elements().Where(child => child.Name.LocalName == localName);
    }

    /// <summary>The val attribute of a child element, or null.</summary>
    private static string? Value(XElement? parent, string localName)
    {
        return Child(parent, localName)?.Attribute("val")?.Value;
    }

    /// <summary>A child's val as a boolean; a child without val is true, as the schema says; no child is null.</summary>
    private static bool? Flag(XElement? parent, string localName)
    {
        XElement? child = Child(parent, localName);
        bool? flag = null;

        if (child is not null)
        {
            flag = Bool(child.Attribute("val")?.Value) ?? true;
        }

        return flag;
    }

    private static bool? Bool(string? text)
    {
        bool? value = text switch
        {
            "1" or "true" => true,
            "0" or "false" => false,
            _ => null,
        };

        return value;
    }

    private static int Int(string? text, int fallback)
    {
        if (int.TryParse(text, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out int value) == true)
        {
            return value;
        }

        return fallback;
    }

    private static int? Clamped(string? text, int minimum, int maximum)
    {
        int? value = null;

        if (int.TryParse(text, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out int number) == true)
        {
            value = Math.Clamp(number, minimum, maximum);
        }

        return value;
    }

    private static double? Number(string? text)
    {
        double? value = null;

        if (DrawingValues.TryDouble(text, out double number) == true)
        {
            value = number;
        }

        return value;
    }

    private static string? Limit(string? text, int length)
    {
        string? limited = text;

        if (text is not null)
        {
            limited = Clean(text[..Math.Min(text.Length, length)]);
        }

        return limited;
    }

    private static string Clean(string value)
    {
        return new string(value.Where(System.Xml.XmlConvert.IsXmlChar).ToArray());
    }
}
