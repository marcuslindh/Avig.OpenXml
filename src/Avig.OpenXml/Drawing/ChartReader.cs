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

                ChartSpaceSeries read = ReadSeries(element, type, main.Type) with { SecondaryAxis = secondary };

                series.Add((Int(Value(element, "idx"), series.Count), read));
            }
        }

        ChartSpace space = new(main.Type)
        {
            Title = TitleText(Child(chart, "title")),
            TitleFont = TitleFont(Child(chart, "title")),
            Series = series.OrderBy(item => item.Index).Select(item => item.Series).ToList(),
            Stacked = Value(main.Element, "grouping") is "stacked" or "percentStacked",
            VaryColors = Flag(main.Element, "varyColors"),
            RoundedCorners = Flag(chartSpace, "roundedCorners") ?? true,
            HoleSize = Clamped(Value(main.Element, "holeSize"), 10, 90),
            FirstSliceAngle = Clamped(Value(main.Element, "firstSliceAng"), 0, 360) ?? 0,
            GapWidth = Clamped(Value(main.Element, "gapWidth"), 0, 500),
            Overlap = Clamped(Value(main.Element, "overlap"), -100, 100),
        };

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
            case "barChart":
                type = ChartType.Column;

                if (Value(plot, "barDir") == "bar")
                {
                    type = ChartType.Bar;
                }

                return true;

            case "lineChart":
                type = ChartType.Line;

                return true;

            case "areaChart":
                type = ChartType.Area;

                return true;

            case "pieChart":
                type = ChartType.Pie;

                return true;

            case "doughnutChart":
                type = ChartType.Doughnut;

                return true;

            case "scatterChart":
                type = ChartType.Scatter;

                return true;

            default:
                return false;
        }
    }

    private static ChartSpaceSeries ReadSeries(XElement element, ChartType type, ChartType chartType)
    {
        string valueName = "val";
        string categoryName = "cat";

        if (type == ChartType.Scatter)
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
        if (main.Type is ChartType.Pie or ChartType.Doughnut || mainAxes.Count < 2)
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
