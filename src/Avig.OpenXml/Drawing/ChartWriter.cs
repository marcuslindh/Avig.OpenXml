using System.Xml;

namespace Avig.OpenXml.Drawing;

/// <summary>
/// Writes a chart part (c:chartSpace). Every element is written in the order the schema gives,
/// which Office checks: a misplaced one makes it repair the file. The series' formulas are written
/// as they are given; what they point to is up to the package that owns the chart.
/// </summary>
public static class ChartWriter
{
    private const string ChartUri = OpcNamespaces.DrawingChart;
    private const string DrawingUri = OpcNamespaces.DrawingMain;

    private const string CategoryAxisId = "500000001";
    private const string ValueAxisId = "500000002";
    private const string SecondaryCategoryAxisId = "500000003";
    private const string SecondaryValueAxisId = "500000004";

    /// <summary>A complete chart part, from the XML declaration on. Call <see cref="ChartSpace.Validate"/> first.</summary>
    /// <param name="xml">The writer of the part.</param>
    /// <param name="chart">The chart.</param>
    /// <returns>A task.</returns>
    public static async Task WriteChartSpaceAsync(XmlWriter xml, ChartSpace chart)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(chart);
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync("c", "chartSpace", ChartUri);
        await xml.WriteAttributeStringAsync("xmlns", "a", null, DrawingUri);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, OpcNamespaces.Relationships);

        // Written always: the schema's default (rounded) is the opposite of Excel's.
        await WriteValueAsync(xml, "roundedCorners", Flag(chart.RoundedCorners));
        await xml.WriteStartElementAsync("c", "chart", ChartUri);

        if (chart.Title is not null)
        {
            await WriteTitleAsync(xml, chart.Title, false, chart.TitleFont);
            await WriteValueAsync(xml, "autoTitleDeleted", "0");
        }
        else
        {
            // Otherwise Excel shows the name of a lone series as the title.
            await WriteValueAsync(xml, "autoTitleDeleted", "1");
        }

        await xml.WriteStartElementAsync("c", "plotArea", ChartUri);
        await xml.WriteElementStringAsync("c", "layout", ChartUri, string.Empty);
        List<SeriesGroup> groups = Groups(chart);

        foreach (SeriesGroup group in groups)
        {
            await WriteGroupAsync(xml, chart, group);
        }

        await WriteAxesAsync(xml, chart, groups.Exists(group => group.Secondary == true));

        // Excel paints a plot area without its own fill white, which would hide a background.
        Fill? plotFill = chart.PlotBackground;

        if (plotFill is null && chart.Background is not null)
        {
            plotFill = Fill.None;
        }

        if (plotFill is not null)
        {
            await xml.WriteStartElementAsync("c", "spPr", ChartUri);
            await DrawingWriter.WriteFillAsync(xml, plotFill);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();

        if (chart.Legend != LegendPosition.None)
        {
            await xml.WriteStartElementAsync("c", "legend", ChartUri);
            await WriteValueAsync(xml, "legendPos", LegendName(chart.Legend));
            await WriteValueAsync(xml, "overlay", "0");
            await WriteTextPropertiesAsync(xml, chart.LegendFont);
            await xml.WriteEndElementAsync();
        }

        await WriteValueAsync(xml, "plotVisOnly", "1");
        await WriteValueAsync(xml, "dispBlanksAs", "gap");
        await xml.WriteEndElementAsync();

        if (chart.Background is not null || chart.Border is not null)
        {
            await xml.WriteStartElementAsync("c", "spPr", ChartUri);

            if (chart.Background is not null)
            {
                await DrawingWriter.WriteFillAsync(xml, chart.Background);
            }

            if (chart.Border is not null)
            {
                await DrawingWriter.WriteOutlineAsync(xml, chart.Border);
            }

            await xml.WriteEndElementAsync();
        }

        await WriteTextPropertiesAsync(xml, chart.Font);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>One plot element (barChart, lineChart ...): the series of one type on one axis.</summary>
    private sealed record SeriesGroup(ChartType Type, bool Secondary, List<(int Index, ChartSpaceSeries Series)> Members);

    /// <summary>
    /// The series grouped by type and axis: primary before secondary, and areas behind columns
    /// behind lines, since later groups are drawn on top.
    /// </summary>
    private static List<SeriesGroup> Groups(ChartSpace chart)
    {
        IReadOnlyList<ChartSpaceSeries> series = chart.Series;
        List<SeriesGroup> groups = [];

        for (int index = 0; index < series.Count; index++)
        {
            ChartType type = chart.TypeOf(series[index]);
            bool secondary = series[index].SecondaryAxis;
            SeriesGroup? group = groups.Find(candidate => candidate.Type == type && candidate.Secondary == secondary);

            if (group is null)
            {
                group = new SeriesGroup(type, secondary, []);
                groups.Add(group);
            }

            group.Members.Add((index, series[index]));
        }

        return groups.OrderBy(group => group.Secondary).ThenBy(group => Rank(group.Type)).ToList();
    }

    private static int Rank(ChartType type)
    {
        return type switch
        {
            ChartType.Area => 0,
            ChartType.Column or ChartType.Bar => 1,
            ChartType.Line => 2,
            _ => 3,
        };
    }

    private static async Task WriteGroupAsync(XmlWriter xml, ChartSpace chart, SeriesGroup group)
    {
        // Stacking applies to the chart's own type; series drawn as another type stand alone.
        bool stacked = chart.Stacked == true && group.Type == chart.Type;
        bool pie = group.Type is ChartType.Pie or ChartType.Doughnut;
        string varyColors = Flag(chart.VaryColors ?? pie);

        switch (group.Type)
        {
            case ChartType.Column or ChartType.Bar:
                await xml.WriteStartElementAsync("c", "barChart", ChartUri);

                if (group.Type == ChartType.Bar)
                {
                    await WriteValueAsync(xml, "barDir", "bar");
                }
                else
                {
                    await WriteValueAsync(xml, "barDir", "col");
                }

                if (stacked == true)
                {
                    await WriteValueAsync(xml, "grouping", "stacked");
                }
                else
                {
                    await WriteValueAsync(xml, "grouping", "clustered");
                }

                await WriteValueAsync(xml, "varyColors", varyColors);
                await WriteSeriesAsync(xml, chart, group);
                await WriteValueAsync(xml, "gapWidth", DrawingValues.Text(chart.GapWidth ?? 150));

                if (chart.Overlap is int overlap)
                {
                    await WriteValueAsync(xml, "overlap", DrawingValues.Text(overlap));
                }
                else if (stacked == true)
                {
                    await WriteValueAsync(xml, "overlap", "100");
                }

                await WriteAxisIdsAsync(xml, group.Secondary);
                await xml.WriteEndElementAsync();

                break;

            case ChartType.Line:
                await xml.WriteStartElementAsync("c", "lineChart", ChartUri);
                await WriteValueAsync(xml, "grouping", Grouping(stacked));
                await WriteValueAsync(xml, "varyColors", varyColors);
                await WriteSeriesAsync(xml, chart, group);
                await WriteValueAsync(xml, "marker", "1");
                await WriteAxisIdsAsync(xml, group.Secondary);
                await xml.WriteEndElementAsync();

                break;

            case ChartType.Area:
                await xml.WriteStartElementAsync("c", "areaChart", ChartUri);
                await WriteValueAsync(xml, "grouping", Grouping(stacked));
                await WriteValueAsync(xml, "varyColors", varyColors);
                await WriteSeriesAsync(xml, chart, group);
                await WriteAxisIdsAsync(xml, group.Secondary);
                await xml.WriteEndElementAsync();

                break;

            case ChartType.Pie:
                await xml.WriteStartElementAsync("c", "pieChart", ChartUri);
                await WriteValueAsync(xml, "varyColors", varyColors);
                await WriteSeriesAsync(xml, chart, group);
                await WriteValueAsync(xml, "firstSliceAng", DrawingValues.Text(chart.FirstSliceAngle));
                await xml.WriteEndElementAsync();

                break;

            case ChartType.Doughnut:
                await xml.WriteStartElementAsync("c", "doughnutChart", ChartUri);
                await WriteValueAsync(xml, "varyColors", varyColors);
                await WriteSeriesAsync(xml, chart, group);
                await WriteValueAsync(xml, "firstSliceAng", DrawingValues.Text(chart.FirstSliceAngle));
                await WriteValueAsync(xml, "holeSize", DrawingValues.Text(chart.HoleSize ?? 50));
                await xml.WriteEndElementAsync();

                break;

            default:
                await xml.WriteStartElementAsync("c", "scatterChart", ChartUri);
                await WriteValueAsync(xml, "scatterStyle", "lineMarker");
                await WriteValueAsync(xml, "varyColors", varyColors);
                await WriteSeriesAsync(xml, chart, group);
                await WriteAxisIdsAsync(xml, group.Secondary);
                await xml.WriteEndElementAsync();

                break;
        }
    }

    private static string Grouping(bool stacked)
    {
        if (stacked == true)
        {
            return "stacked";
        }

        return "standard";
    }

    /// <summary>The series, each with its children in the order its type's schema gives.</summary>
    private static async Task WriteSeriesAsync(XmlWriter xml, ChartSpace chart, SeriesGroup group)
    {
        ChartType type = group.Type;
        bool pie = type is ChartType.Pie or ChartType.Doughnut;
        bool lines = type is ChartType.Line or ChartType.Scatter;
        bool bars = type is ChartType.Column or ChartType.Bar;

        foreach ((int index, ChartSpaceSeries series) in group.Members)
        {
            await xml.WriteStartElementAsync("c", "ser", ChartUri);
            await WriteValueAsync(xml, "idx", DrawingValues.Text(index));
            await WriteValueAsync(xml, "order", DrawingValues.Text(index));

            if (series.NameFormula is not null)
            {
                await xml.WriteStartElementAsync("c", "tx", ChartUri);
                await xml.WriteStartElementAsync("c", "strRef", ChartUri);
                await xml.WriteElementStringAsync("c", "f", ChartUri, series.NameFormula);

                if (series.Name is not null)
                {
                    await WritePointsAsync(xml, true, "strCache", null, 1, [series.Name]);
                }

                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();
            }
            else if (series.Name is not null)
            {
                await xml.WriteStartElementAsync("c", "tx", ChartUri);
                await xml.WriteElementStringAsync("c", "v", ChartUri, series.Name);
                await xml.WriteEndElementAsync();
            }

            // A pie colours each slice, so one colour for the whole series would hide them,
            // unless the chart turns that off.
            string? color = series.Color;

            if (pie == true && (chart.VaryColors ?? true) == true)
            {
                color = null;
            }

            await WriteSeriesShapeAsync(xml, type, series, color);

            if (lines == true)
            {
                await WriteMarkerAsync(xml, series, color ?? series.Outline?.Color);
            }

            if (bars == true)
            {
                await WriteValueAsync(xml, "invertIfNegative", "0");
            }

            for (int point = 0; point < series.PointColors.Count; point++)
            {
                if (series.PointColors[point] is string pointColor)
                {
                    await WritePointAsync(xml, type, point, pointColor, series);
                }
            }

            if ((series.DataLabels ?? chart.DataLabels) is DataLabels labels)
            {
                await WriteDataLabelsAsync(xml, labels);
            }

            if (type == ChartType.Scatter)
            {
                if (series.Categories is not null)
                {
                    await WriteReferenceAsync(xml, "xVal", true, series.Categories);
                }

                await WriteReferenceAsync(xml, "yVal", true, series.Values);
            }
            else
            {
                if (series.Categories is not null)
                {
                    await WriteReferenceAsync(xml, "cat", false, series.Categories);
                }

                await WriteReferenceAsync(xml, "val", true, series.Values);
            }

            if (lines == true)
            {
                await WriteValueAsync(xml, "smooth", Flag(series.Smooth));
            }

            await xml.WriteEndElementAsync();
        }
    }

    /// <summary>The series' c:spPr: the fill, line and shadow its type uses, when any is set.</summary>
    private static async Task WriteSeriesShapeAsync(XmlWriter xml, ChartType type, ChartSpaceSeries series, string? color)
    {
        Fill? fill = null;
        Outline? outline = series.Outline;

        if (type is ChartType.Line or ChartType.Scatter)
        {
            if (outline is null && color is not null)
            {
                outline = new Outline(color, 2.25);
            }

            // A plain scatter chart shows points without lines between them, unless asked.
            if (type == ChartType.Scatter && series.Outline is null && series.Smooth == false)
            {
                outline = Outline.None;
            }
        }
        else
        {
            fill = series.Fill;

            if (fill is null && color is not null)
            {
                fill = Fill.Solid(color);
            }
        }

        if (fill is null && outline is null && series.Shadow is null)
        {
            return;
        }

        await xml.WriteStartElementAsync("c", "spPr", ChartUri);

        if (fill is not null)
        {
            await DrawingWriter.WriteFillAsync(xml, fill);
        }

        if (outline is not null)
        {
            if (type == ChartType.Scatter && outline.Color is null)
            {
                // As Excel writes a scatter series without lines.
                await xml.WriteStartElementAsync("a", "ln", DrawingUri);
                await xml.WriteAttributeStringAsync(null, "w", null, "19050");
                await xml.WriteElementStringAsync("a", "noFill", DrawingUri, string.Empty);
                await xml.WriteEndElementAsync();
            }
            else
            {
                await DrawingWriter.WriteOutlineAsync(xml, outline);
            }
        }

        if (series.Shadow is not null)
        {
            await DrawingWriter.WriteEffectsAsync(xml, series.Shadow, null, null);
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteMarkerAsync(XmlWriter xml, ChartSpaceSeries series, string? color)
    {
        if (series.Marker is null && series.MarkerSize is null && color is null)
        {
            return;
        }

        await xml.WriteStartElementAsync("c", "marker", ChartUri);
        MarkerStyle style = series.Marker ?? MarkerStyle.Circle;
        await WriteValueAsync(xml, "symbol", MarkerName(style));

        if (style != MarkerStyle.None)
        {
            await WriteValueAsync(xml, "size", DrawingValues.Text(series.MarkerSize ?? 5));

            if (color is not null)
            {
                await xml.WriteStartElementAsync("c", "spPr", ChartUri);
                await DrawingWriter.WriteFillAsync(xml, Fill.Solid(color));
                await DrawingWriter.WriteOutlineAsync(xml, new Outline(color, 0.75));
                await xml.WriteEndElementAsync();
            }
        }

        await xml.WriteEndElementAsync();
    }

    private static string MarkerName(MarkerStyle style)
    {
        return style switch
        {
            MarkerStyle.None => "none",
            MarkerStyle.Square => "square",
            MarkerStyle.Diamond => "diamond",
            MarkerStyle.Triangle => "triangle",
            MarkerStyle.X => "x",
            MarkerStyle.Star => "star",
            MarkerStyle.Dash => "dash",
            MarkerStyle.Dot => "dot",
            MarkerStyle.Plus => "plus",
            _ => "circle",
        };
    }

    /// <summary>A c:dPt that gives one point its own colour.</summary>
    private static async Task WritePointAsync(XmlWriter xml, ChartType type, int point, string color, ChartSpaceSeries series)
    {
        await xml.WriteStartElementAsync("c", "dPt", ChartUri);
        await WriteValueAsync(xml, "idx", DrawingValues.Text(point));

        if (type is ChartType.Column or ChartType.Bar)
        {
            await WriteValueAsync(xml, "invertIfNegative", "0");
        }

        await WriteValueAsync(xml, "bubble3D", "0");
        await xml.WriteStartElementAsync("c", "spPr", ChartUri);

        if (type is ChartType.Line or ChartType.Scatter)
        {
            double width = series.Outline?.Width ?? 2.25;
            await DrawingWriter.WriteOutlineAsync(xml, new Outline(color, width));
        }
        else
        {
            await DrawingWriter.WriteFillAsync(xml, Fill.Solid(color));

            if (series.Outline is not null)
            {
                await DrawingWriter.WriteOutlineAsync(xml, series.Outline);
            }
        }

        if (series.Shadow is not null)
        {
            await DrawingWriter.WriteEffectsAsync(xml, series.Shadow, null, null);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteDataLabelsAsync(XmlWriter xml, DataLabels labels)
    {
        await xml.WriteStartElementAsync("c", "dLbls", ChartUri);

        if (labels.NumberFormat is not null)
        {
            await xml.WriteStartElementAsync("c", "numFmt", ChartUri);
            await xml.WriteAttributeStringAsync(null, "formatCode", null, labels.NumberFormat);
            await xml.WriteAttributeStringAsync(null, "sourceLinked", null, "0");
            await xml.WriteEndElementAsync();
        }

        // Excel's own labels have no box around them.
        await xml.WriteStartElementAsync("c", "spPr", ChartUri);
        await xml.WriteElementStringAsync("a", "noFill", DrawingUri, string.Empty);
        await xml.WriteStartElementAsync("a", "ln", DrawingUri);
        await xml.WriteElementStringAsync("a", "noFill", DrawingUri, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await WriteTextPropertiesAsync(xml, labels.Font);

        if (labels.Position != DataLabelPosition.Default)
        {
            await WriteValueAsync(xml, "dLblPos", PositionName(labels.Position));
        }

        await WriteValueAsync(xml, "showLegendKey", "0");
        await WriteValueAsync(xml, "showVal", Flag(labels.ShowValue));
        await WriteValueAsync(xml, "showCatName", Flag(labels.ShowCategory));
        await WriteValueAsync(xml, "showSerName", Flag(labels.ShowSeriesName));
        await WriteValueAsync(xml, "showPercent", Flag(labels.ShowPercent));
        await WriteValueAsync(xml, "showBubbleSize", "0");
        await xml.WriteEndElementAsync();
    }

    private static string PositionName(DataLabelPosition position)
    {
        return position switch
        {
            DataLabelPosition.Center => "ctr",
            DataLabelPosition.InsideEnd => "inEnd",
            DataLabelPosition.InsideBase => "inBase",
            DataLabelPosition.OutsideEnd => "outEnd",
            DataLabelPosition.Above => "t",
            DataLabelPosition.Below => "b",
            DataLabelPosition.Left => "l",
            DataLabelPosition.Right => "r",
            _ => "bestFit",
        };
    }

    private static async Task WriteAxesAsync(XmlWriter xml, ChartSpace chart, bool secondary)
    {
        if (chart.Type is ChartType.Pie or ChartType.Doughnut)
        {
            return;
        }

        ChartAxis categoryAxis = chart.CategoryAxis ?? new ChartAxis();
        ChartAxis valueAxis = chart.ValueAxis ?? new ChartAxis();
        string categoryPosition = "b";
        string valuePosition = "l";

        if (chart.Type == ChartType.Bar)
        {
            categoryPosition = "l";
            valuePosition = "b";
        }

        if (chart.Type == ChartType.Scatter)
        {
            // Two value axes: x along the bottom, y up the side.
            await WriteValueAxisAsync(xml, new AxisPlace(CategoryAxisId, ValueAxisId, "b", "midCat", false), chart.CategoryAxisTitle, categoryAxis, false);
            await WriteValueAxisAsync(xml, new AxisPlace(ValueAxisId, CategoryAxisId, "l", "midCat", false), chart.ValueAxisTitle, valueAxis, true);

            return;
        }

        await WriteCategoryAxisAsync(xml, CategoryAxisId, ValueAxisId, categoryPosition, chart.CategoryAxisTitle, categoryAxis);
        await WriteValueAxisAsync(xml, new AxisPlace(ValueAxisId, CategoryAxisId, valuePosition, "between", false), chart.ValueAxisTitle, valueAxis, true);

        if (secondary == true)
        {
            // The secondary series share the categories, so their own category axis is hidden;
            // their value axis sits on the right, crossing at the far end.
            await WriteCategoryAxisAsync(xml, SecondaryCategoryAxisId, SecondaryValueAxisId, "b", null, new ChartAxis { Visible = false });
            await WriteValueAxisAsync(xml, new AxisPlace(SecondaryValueAxisId, SecondaryCategoryAxisId, "r", "between", true), null, chart.SecondaryValueAxis ?? new ChartAxis(), false);
        }
    }

    private static async Task WriteCategoryAxisAsync(XmlWriter xml, string id, string crossing, string position, string? title, ChartAxis axis)
    {
        await xml.WriteStartElementAsync("c", "catAx", ChartUri);
        await WriteValueAsync(xml, "axId", id);
        await WriteScalingAsync(xml, axis, false);
        await WriteValueAsync(xml, "delete", Flag(axis.Visible == false));
        await WriteValueAsync(xml, "axPos", position);
        await WriteGridlinesAsync(xml, axis, false);

        if (title is not null)
        {
            await WriteTitleAsync(xml, title, false, null);
        }

        if (axis.NumberFormat is not null)
        {
            await WriteNumberFormatAsync(xml, axis.NumberFormat, false);
        }

        await WriteValueAsync(xml, "majorTickMark", "out");
        await WriteValueAsync(xml, "minorTickMark", "none");
        await WriteValueAsync(xml, "tickLblPos", "nextTo");
        await WriteAxisShapeAsync(xml, axis);
        await WriteTextPropertiesAsync(xml, axis.Font);
        await WriteValueAsync(xml, "crossAx", crossing);
        await WriteValueAsync(xml, "crosses", "autoZero");
        await WriteValueAsync(xml, "auto", "1");
        await WriteValueAsync(xml, "lblAlgn", "ctr");
        await WriteValueAsync(xml, "lblOffset", "100");
        await WriteValueAsync(xml, "noMultiLvlLbl", "0");
        await xml.WriteEndElementAsync();
    }

    /// <summary>Where a value axis sits: its id, the axis it crosses, its side, how it meets the categories, and whether it crosses at the far end.</summary>
    private readonly record struct AxisPlace(string Id, string Crossing, string Position, string CrossBetween, bool CrossesAtMaximum);

    private static async Task WriteValueAxisAsync(XmlWriter xml, AxisPlace place, string? title, ChartAxis axis, bool gridlines)
    {
        await xml.WriteStartElementAsync("c", "valAx", ChartUri);
        await WriteValueAsync(xml, "axId", place.Id);
        await WriteScalingAsync(xml, axis, true);
        await WriteValueAsync(xml, "delete", Flag(axis.Visible == false));
        await WriteValueAsync(xml, "axPos", place.Position);
        await WriteGridlinesAsync(xml, axis, gridlines);

        if (title is not null)
        {
            await WriteTitleAsync(xml, title, place.Position is "l" or "r", null);
        }

        await WriteNumberFormatAsync(xml, axis.NumberFormat ?? "General", axis.NumberFormat is null);
        await WriteValueAsync(xml, "majorTickMark", "out");
        await WriteValueAsync(xml, "minorTickMark", "none");
        await WriteValueAsync(xml, "tickLblPos", "nextTo");
        await WriteAxisShapeAsync(xml, axis);
        await WriteTextPropertiesAsync(xml, axis.Font);
        await WriteValueAsync(xml, "crossAx", place.Crossing);

        if (place.CrossesAtMaximum == true)
        {
            await WriteValueAsync(xml, "crosses", "max");
        }
        else
        {
            await WriteValueAsync(xml, "crosses", "autoZero");
        }

        await WriteValueAsync(xml, "crossBetween", place.CrossBetween);

        if (axis.MajorUnit is double unit)
        {
            await WriteValueAsync(xml, "majorUnit", DrawingValues.Text(unit));
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteScalingAsync(XmlWriter xml, ChartAxis axis, bool values)
    {
        await xml.WriteStartElementAsync("c", "scaling", ChartUri);
        await WriteValueAsync(xml, "orientation", "minMax");

        if (values == true && axis.Maximum is double maximum)
        {
            await WriteValueAsync(xml, "max", DrawingValues.Text(maximum));
        }

        if (values == true && axis.Minimum is double minimum)
        {
            await WriteValueAsync(xml, "min", DrawingValues.Text(minimum));
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteGridlinesAsync(XmlWriter xml, ChartAxis axis, bool byDefault)
    {
        if ((axis.Gridlines ?? byDefault) == false)
        {
            return;
        }

        await xml.WriteStartElementAsync("c", "majorGridlines", ChartUri);

        if (axis.GridlineOutline is not null)
        {
            await xml.WriteStartElementAsync("c", "spPr", ChartUri);
            await DrawingWriter.WriteOutlineAsync(xml, axis.GridlineOutline);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteAxisShapeAsync(XmlWriter xml, ChartAxis axis)
    {
        if (axis.Line is null)
        {
            return;
        }

        await xml.WriteStartElementAsync("c", "spPr", ChartUri);
        await DrawingWriter.WriteOutlineAsync(xml, axis.Line);
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteNumberFormatAsync(XmlWriter xml, string format, bool sourceLinked)
    {
        await xml.WriteStartElementAsync("c", "numFmt", ChartUri);
        await xml.WriteAttributeStringAsync(null, "formatCode", null, format);
        await xml.WriteAttributeStringAsync(null, "sourceLinked", null, Flag(sourceLinked));
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteAxisIdsAsync(XmlWriter xml, bool secondary)
    {
        if (secondary == true)
        {
            await WriteValueAsync(xml, "axId", SecondaryCategoryAxisId);
            await WriteValueAsync(xml, "axId", SecondaryValueAxisId);

            return;
        }

        await WriteValueAsync(xml, "axId", CategoryAxisId);
        await WriteValueAsync(xml, "axId", ValueAxisId);
    }

    /// <summary>A title with one run of text. <paramref name="vertical"/> turns it to read upwards, as Excel does beside a vertical axis.</summary>
    private static async Task WriteTitleAsync(XmlWriter xml, string text, bool vertical, TextFont? font)
    {
        await xml.WriteStartElementAsync("c", "title", ChartUri);
        await xml.WriteStartElementAsync("c", "tx", ChartUri);
        await xml.WriteStartElementAsync("c", "rich", ChartUri);
        await xml.WriteStartElementAsync("a", "bodyPr", DrawingUri);

        if (vertical == true)
        {
            await xml.WriteAttributeStringAsync(null, "rot", null, "-5400000");
            await xml.WriteAttributeStringAsync(null, "vert", null, "horz");
        }

        await xml.WriteEndElementAsync();
        await xml.WriteElementStringAsync("a", "lstStyle", DrawingUri, string.Empty);
        await xml.WriteStartElementAsync("a", "p", DrawingUri);

        if (font is not null)
        {
            await xml.WriteStartElementAsync("a", "pPr", DrawingUri);
            await DrawingWriter.WriteRunPropertiesAsync(xml, "defRPr", font);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteStartElementAsync("a", "r", DrawingUri);

        if (font is not null)
        {
            await DrawingWriter.WriteRunPropertiesAsync(xml, "rPr", font);
        }

        await xml.WriteElementStringAsync("a", "t", DrawingUri, text);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await WriteValueAsync(xml, "overlay", "0");
        await xml.WriteEndElementAsync();
    }

    /// <summary>A c:txPr: the default look of the text of the element it is in.</summary>
    private static async Task WriteTextPropertiesAsync(XmlWriter xml, TextFont? font)
    {
        if (font is null)
        {
            return;
        }

        await xml.WriteStartElementAsync("c", "txPr", ChartUri);
        await xml.WriteElementStringAsync("a", "bodyPr", DrawingUri, string.Empty);
        await xml.WriteElementStringAsync("a", "lstStyle", DrawingUri, string.Empty);
        await xml.WriteStartElementAsync("a", "p", DrawingUri);
        await xml.WriteStartElementAsync("a", "pPr", DrawingUri);
        await DrawingWriter.WriteRunPropertiesAsync(xml, "defRPr", font);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "endParaRPr", DrawingUri);
        await xml.WriteAttributeStringAsync(null, "lang", null, "en-US");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    /// <summary>
    /// A cat, val, xVal or yVal: a numRef or strRef with the formula and the stored values, or a
    /// numLit or strLit with values alone. <paramref name="numbers"/> is the kind the element
    /// takes when the data does not say; stored texts or numbers decide otherwise.
    /// </summary>
    private static async Task WriteReferenceAsync(XmlWriter xml, string element, bool numbers, ChartData data)
    {
        if (data.Texts is not null && data.Numbers is null)
        {
            numbers = false;
        }
        else if (data.Numbers is not null)
        {
            numbers = true;
        }

        string kind = (numbers, data.Formula is not null) switch
        {
            (true, true) => "numRef",
            (true, false) => "numLit",
            (false, true) => "strRef",
            (false, false) => "strLit",
        };

        await xml.WriteStartElementAsync("c", element, ChartUri);
        await xml.WriteStartElementAsync("c", kind, ChartUri);

        if (data.Formula is not null)
        {
            await xml.WriteElementStringAsync("c", "f", ChartUri, data.Formula);
        }

        if (numbers == true && data.Numbers is not null)
        {
            await WritePointsAsync(xml, data.Formula is not null, "numCache", data.FormatCode ?? "General", data.Numbers.Count, Numbers(data.Numbers));
        }
        else if (numbers == false && data.Texts is not null)
        {
            await WritePointsAsync(xml, data.Formula is not null, "strCache", null, data.Texts.Count, data.Texts);
        }
        else if (data.Formula is null)
        {
            // A literal without values still says it has none.
            await WriteValueAsync(xml, "ptCount", "0");
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    /// <summary>A cache (or, for a literal, the points alone): the format code of numbers, the count, and a c:pt per point that has a value.</summary>
    private static async Task WritePointsAsync(XmlWriter xml, bool cache, string cacheName, string? formatCode, int count, IEnumerable<string?> points)
    {
        if (cache == true)
        {
            await xml.WriteStartElementAsync("c", cacheName, ChartUri);
        }

        if (formatCode is not null)
        {
            await xml.WriteElementStringAsync("c", "formatCode", ChartUri, formatCode);
        }

        await WriteValueAsync(xml, "ptCount", DrawingValues.Text(count));
        int index = 0;

        foreach (string? point in points)
        {
            if (point is not null)
            {
                await xml.WriteStartElementAsync("c", "pt", ChartUri);
                await xml.WriteAttributeStringAsync(null, "idx", null, DrawingValues.Text(index));
                await xml.WriteElementStringAsync("c", "v", ChartUri, point);
                await xml.WriteEndElementAsync();
            }

            index++;
        }

        if (cache == true)
        {
            await xml.WriteEndElementAsync();
        }
    }

    private static List<string?> Numbers(IReadOnlyList<double?> values)
    {
        List<string?> texts = [];

        foreach (double? value in values)
        {
            string? text = null;

            if (value is double number)
            {
                text = DrawingValues.Text(number);
            }

            texts.Add(text);
        }

        return texts;
    }

    private static async Task WriteValueAsync(XmlWriter xml, string name, string value)
    {
        await xml.WriteStartElementAsync("c", name, ChartUri);
        await xml.WriteAttributeStringAsync(null, "val", null, value);
        await xml.WriteEndElementAsync();
    }

    private static string Flag(bool value)
    {
        if (value == true)
        {
            return "1";
        }

        return "0";
    }

    private static string LegendName(LegendPosition position)
    {
        return position switch
        {
            LegendPosition.Bottom => "b",
            LegendPosition.Top => "t",
            LegendPosition.Left => "l",
            _ => "r",
        };
    }
}
