using System.Globalization;
using System.Xml.Linq;

namespace Avig.OpenXml.Drawing;

/// <summary>
/// Reads a chart part of the newer chart engine (cx:chartSpace) back into a
/// <see cref="ChartExSpace"/>: the type, the categories and values with their formulas, the
/// series' names, colours and labels, and the type's options. Elements are read by local name;
/// values outside what the types take are clamped.
/// </summary>
public static class ChartExReader
{
    /// <summary>The most points read from one level, so that a hostile ptCount cannot make a list without end.</summary>
    private const int MaxPoints = 1_048_576;

    /// <summary>Reads a chart part.</summary>
    /// <param name="chartSpace">The root element, cx:chartSpace.</param>
    /// <returns>The chart; Invalid when the element is not a chartSpace or has no series of a type this layer knows.</returns>
    public static Result<ChartExSpace> Read(XElement chartSpace)
    {
        ArgumentNullException.ThrowIfNull(chartSpace);

        if (chartSpace.Name.LocalName != "chartSpace" || chartSpace.Name.NamespaceName != ChartExWriter.Namespace)
        {
            return Result<ChartExSpace>.Invalid($"<{chartSpace.Name.LocalName}> is not the root of a chart part of the newer chart engine, cx:chartSpace.");
        }

        XElement? chartData = Child(chartSpace, "chartData");
        XElement? chart = Child(chartSpace, "chart");
        XElement? plotArea = Child(chart, "plotArea");
        List<XElement> seriesElements = Children(Child(plotArea, "plotAreaRegion"), "series").ToList();
        List<XElement> own = seriesElements.Where(element => Attr(element, "layoutId") != "paretoLine" && Attr(element, "ownerIdx") is null).ToList();

        if (own.Count == 0 || TryType(Attr(own[0], "layoutId"), seriesElements, out ChartExType type) == false)
        {
            return Result<ChartExSpace>.Invalid($"{OpcErrors.DamagedPrefix}: the chart has no series of a type that can be read.");
        }

        Dictionary<string, XElement> data = new(StringComparer.Ordinal);

        foreach (XElement element in Children(chartData, "data"))
        {
            data.TryAdd(Attr(element, "id") ?? string.Empty, element);
        }

        List<ChartExSeries> series = [];
        ChartExCategories? categories = null;

        foreach (XElement element in own.Take(255))
        {
            data.TryGetValue(Value(element, "dataId") ?? string.Empty, out XElement? block);
            categories ??= Categories(Child(block, "strDim"));
            series.Add(ReadSeries(element, block));
        }

        ChartExSpace read = new(type)
        {
            Title = TitleText(Child(chart, "title")),
            Series = series,
            Categories = categories,
            Legend = LegendOf(Child(chart, "legend")),
        };

        XElement? layout = Child(own[0], "layoutPr");
        read = read with
        {
            ParentLabels = Value(layout, "parentLabelLayout") switch
            {
                "banner" => TreemapParentLabels.Banner,
                "none" => TreemapParentLabels.None,
                _ => TreemapParentLabels.Overlapping,
            },
            ConnectorLines = Flag(Child(layout, "visibility"), "connectorLines") ?? true,
            MeanLine = Flag(Child(layout, "visibility"), "meanLine") ?? false,
            MeanMarkers = Flag(Child(layout, "visibility"), "meanMarker") ?? true,
            InnerPoints = Flag(Child(layout, "visibility"), "nonoutliers") ?? false,
            Outliers = Flag(Child(layout, "visibility"), "outliers") ?? true,
            InclusiveQuartiles = Attr(Child(layout, "statistics"), "quartileMethod") == "inclusive",
            Binning = Binning(Child(layout, "binning")),
        };

        List<XElement> axes = Children(plotArea, "axis").ToList();
        XElement? categoryAxis = axes.FirstOrDefault(axis => Child(axis, "catScaling") is not null);
        XElement? valueAxis = axes.FirstOrDefault(axis => Child(axis, "valScaling") is not null);

        if (Flag(categoryAxis, "hidden") == true)
        {
            read = read with { CategoryAxis = false };
        }

        if (Flag(valueAxis, "hidden") == true)
        {
            read = read with { ValueAxis = false };
        }

        if (valueAxis is not null && Child(valueAxis, "majorGridlines") is null)
        {
            read = read with { Gridlines = false };
        }

        XAttribute? external = Child(chartData, "externalData")?.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "id");

        if (external is not null && external.Value.Length > 0 && external.Value.Length <= 255)
        {
            read = read with { ExternalData = Clean(external.Value) };
        }

        return read;
    }

    private static bool TryType(string? layoutId, List<XElement> series, out ChartExType type)
    {
        type = ChartExType.Histogram;

        switch (layoutId)
        {
            case "waterfall":
                type = ChartExType.Waterfall;

                return true;

            case "treemap":
                type = ChartExType.Treemap;

                return true;

            case "sunburst":
                type = ChartExType.Sunburst;

                return true;

            case "boxWhisker":
                type = ChartExType.BoxWhisker;

                return true;

            case "funnel":
                type = ChartExType.Funnel;

                return true;

            case "regionMap":
                type = ChartExType.RegionMap;

                return true;

            case "clusteredColumn":
                if (series.Any(element => Attr(element, "layoutId") == "paretoLine") == true)
                {
                    type = ChartExType.Pareto;
                }

                return true;

            default:
                return false;
        }
    }

    private static ChartExSeries ReadSeries(XElement element, XElement? block)
    {
        XElement? values = Children(block, "numDim").FirstOrDefault();
        ChartExSeries series = new(Numbers(values));
        XElement? text = Child(Child(element, "tx"), "txData");

        if (text is not null)
        {
            series = series with
            {
                Name = Limit(Child(text, "v")?.Value, 255),
                NameFormula = Formula(Child(text, "f")),
            };
        }

        if (Child(element, "spPr") is XElement properties && DrawingReader.TryReadFill(properties, out Fill? fill) == true && fill.Kind == FillKind.Solid)
        {
            series = series with { Color = fill.Color };
        }

        if (Child(element, "dataLabels") is XElement labels)
        {
            XElement? visibility = Child(labels, "visibility");
            DataLabels read = new()
            {
                Position = DrawingNames.PositionOf(Attr(labels, "pos")),
                ShowValue = Flag(visibility, "value") ?? false,
                ShowCategory = Flag(visibility, "categoryName") ?? false,
                ShowSeriesName = Flag(visibility, "seriesName") ?? false,
            };

            if (Attr(Child(labels, "numFmt"), "formatCode") is string format && format.Length > 0 && format.Length <= 255)
            {
                read = read with { NumberFormat = format };
            }

            series = series with { DataLabels = read };
        }

        List<int> subtotals = [];

        foreach (XElement subtotal in Children(Child(Child(element, "layoutPr"), "subtotals"), "idx").Take(MaxPoints))
        {
            if (int.TryParse(Attr(subtotal, "val"), NumberStyles.None, CultureInfo.InvariantCulture, out int point) == true)
            {
                subtotals.Add(point);
            }
        }

        return series with { Subtotals = subtotals };
    }

    private static ChartExCategories? Categories(XElement? dimension)
    {
        ChartExCategories? read = null;
        List<IReadOnlyList<string?>> levels = [];

        foreach (XElement level in Children(dimension, "lvl").Take(32))
        {
            int count = Count(level);
            string?[] names = new string?[count];

            foreach (XElement point in Children(level, "pt"))
            {
                if (int.TryParse(Attr(point, "idx"), NumberStyles.None, CultureInfo.InvariantCulture, out int index) == true && index < count)
                {
                    names[index] = Clean(point.Value);
                }
            }

            levels.Add(names);
        }

        // The file has the leaves first; the type has the outermost level first.
        levels.Reverse();

        if (levels.Count > 0)
        {
            read = new ChartExCategories(levels) { Formula = Formula(Child(dimension, "f")) };
        }

        return read;
    }

    private static ChartData Numbers(XElement? dimension)
    {
        XElement? level = Child(dimension, "lvl");
        int count = Count(level);
        double?[] numbers = new double?[count];

        foreach (XElement point in Children(level, "pt"))
        {
            if (int.TryParse(Attr(point, "idx"), NumberStyles.None, CultureInfo.InvariantCulture, out int index) == true && index < count && DrawingValues.TryDouble(point.Value, out double number) == true)
            {
                numbers[index] = number;
            }
        }

        ChartData data = new() { Numbers = numbers, Formula = Formula(Child(dimension, "f")) };

        if (Attr(level, "formatCode") is string format && format != "General" && format.Length <= 255)
        {
            data = data with { FormatCode = format };
        }

        return data;
    }

    private static int Count(XElement? level)
    {
        int count = 0;

        if (int.TryParse(Attr(level, "ptCount"), NumberStyles.None, CultureInfo.InvariantCulture, out int declared) == true)
        {
            count = Math.Clamp(declared, 0, MaxPoints);
        }

        foreach (XElement point in Children(level, "pt"))
        {
            if (int.TryParse(Attr(point, "idx"), NumberStyles.None, CultureInfo.InvariantCulture, out int index) == true && index >= count && index < MaxPoints)
            {
                count = index + 1;
            }
        }

        return count;
    }

    private static ChartExBinning? Binning(XElement? binning)
    {
        ChartExBinning? read = null;

        if (binning is not null)
        {
            read = new ChartExBinning
            {
                ClosedLeft = Attr(binning, "intervalClosed") == "l",
                Underflow = Number(Attr(binning, "underflow")),
                Overflow = Number(Attr(binning, "overflow")),
            };

            if (Number(Value(binning, "binSize") ?? Child(binning, "binSize")?.Value) is double width && width > 0)
            {
                read = read with { BinWidth = width };
            }

            if (int.TryParse(Value(binning, "binCount") ?? Child(binning, "binCount")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out int count) == true)
            {
                read = read with { BinCount = Math.Clamp(count, 1, 1000) };
            }
        }

        return read;
    }

    private static string? TitleText(XElement? title)
    {
        string? text = null;
        XElement? tx = Child(title, "tx");

        if (Child(Child(tx, "txData"), "v") is XElement value)
        {
            text = Limit(value.Value, 255);
        }
        else if (Child(tx, "rich") is XElement rich)
        {
            text = Limit(string.Concat(rich.Descendants().Where(element => element.Name.LocalName == "t").Select(element => element.Value)), 255);
        }

        return text;
    }

    private static LegendPosition? LegendOf(XElement? legend)
    {
        if (legend is null)
        {
            return LegendPosition.None;
        }

        return DrawingNames.LegendOf(Attr(legend, "pos"));
    }

    private static string? Formula(XElement? formula)
    {
        string? text = null;

        if (formula is not null && formula.Value.Length > 0 && formula.Value.Length <= 8192)
        {
            text = Clean(formula.Value);
        }

        return text;
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

    private static string? Attr(XElement? element, string localName)
    {
        return element?.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == localName)?.Value;
    }

    private static string? Value(XElement? parent, string localName)
    {
        return Attr(Child(parent, localName), "val");
    }

    private static bool? Flag(XElement? element, string localName)
    {
        bool? flag = Attr(element, localName) switch
        {
            "1" or "true" => true,
            "0" or "false" => false,
            _ => null,
        };

        return flag;
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
