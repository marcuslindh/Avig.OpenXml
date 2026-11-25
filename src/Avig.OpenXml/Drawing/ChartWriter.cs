using System.Xml;

namespace Avig.OpenXml.Drawing;

/// <summary>
/// Writes a chart part (c:chartSpace). Every element is written in the order the schema gives,
/// which Office checks: a misplaced one makes it repair the file. The series' formulas are written
/// as they are given; what they point to is up to the package that owns the chart.
/// </summary>
public static partial class ChartWriter
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
    public static void WriteChartSpace(XmlWriter xml, ChartSpace chart)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(chart);
        xml.WriteStartDocument(true);
        xml.WriteStartElement("c", "chartSpace", ChartUri);
        xml.WriteAttributeString("xmlns", "a", null, DrawingUri);
        xml.WriteAttributeString("xmlns", "r", null, OpcNamespaces.Relationships);

        // Written always: the schema's default (rounded) is the opposite of Excel's.
        WriteValue(xml, "roundedCorners", Flag(chart.RoundedCorners));
        xml.WriteStartElement("c", "chart", ChartUri);

        if (chart.Title is not null)
        {
            WriteTitle(xml, chart.Title, false, chart.TitleFont);
            WriteValue(xml, "autoTitleDeleted", "0");
        }
        else
        {
            // Otherwise Excel shows the name of a lone series as the title.
            WriteValue(xml, "autoTitleDeleted", "1");
        }

        xml.WriteStartElement("c", "plotArea", ChartUri);
        xml.WriteElementString("c", "layout", ChartUri, string.Empty);
        List<SeriesGroup> groups = Groups(chart);

        foreach (SeriesGroup group in groups)
        {
            WriteGroup(xml, chart, group);
        }

        WriteAxes(xml, chart, groups.Exists(group => group.Secondary == true));

        // Excel paints a plot area without its own fill white, which would hide a background.
        Fill? plotFill = chart.PlotBackground;

        if (plotFill is null && chart.Background is not null)
        {
            plotFill = Fill.None;
        }

        if (plotFill is not null)
        {
            xml.WriteStartElement("c", "spPr", ChartUri);
            DrawingWriter.WriteFill(xml, plotFill);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();

        if (chart.Legend != LegendPosition.None)
        {
            xml.WriteStartElement("c", "legend", ChartUri);
            WriteValue(xml, "legendPos", LegendName(chart.Legend));
            WriteValue(xml, "overlay", "0");
            WriteTextProperties(xml, chart.LegendFont);
            xml.WriteEndElement();
        }

        WriteValue(xml, "plotVisOnly", "1");
        WriteValue(xml, "dispBlanksAs", "gap");
        xml.WriteEndElement();

        if (chart.Background is not null || chart.Border is not null)
        {
            xml.WriteStartElement("c", "spPr", ChartUri);

            if (chart.Background is not null)
            {
                DrawingWriter.WriteFill(xml, chart.Background);
            }

            if (chart.Border is not null)
            {
                DrawingWriter.WriteOutline(xml, chart.Border);
            }

            xml.WriteEndElement();
        }

        WriteTextProperties(xml, chart.Font);
        xml.WriteEndElement();
        xml.WriteEndDocument();
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

    private static void WriteGroup(XmlWriter xml, ChartSpace chart, SeriesGroup group)
    {
        // Stacking applies to the chart's own type; series drawn as another type stand alone.
        bool stacked = chart.Stacked == true && group.Type == chart.Type;
        bool pie = group.Type is ChartType.Pie or ChartType.Doughnut;
        string varyColors = Flag(chart.VaryColors ?? pie);

        switch (group.Type)
        {
            case ChartType.Column or ChartType.Bar:
                xml.WriteStartElement("c", "barChart", ChartUri);

                if (group.Type == ChartType.Bar)
                {
                    WriteValue(xml, "barDir", "bar");
                }
                else
                {
                    WriteValue(xml, "barDir", "col");
                }

                if (stacked == true)
                {
                    WriteValue(xml, "grouping", "stacked");
                }
                else
                {
                    WriteValue(xml, "grouping", "clustered");
                }

                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteValue(xml, "gapWidth", DrawingValues.Text(chart.GapWidth ?? 150));

                if (chart.Overlap is int overlap)
                {
                    WriteValue(xml, "overlap", DrawingValues.Text(overlap));
                }
                else if (stacked == true)
                {
                    WriteValue(xml, "overlap", "100");
                }

                WriteAxisIds(xml, group.Secondary);
                xml.WriteEndElement();

                break;

            case ChartType.Line:
                xml.WriteStartElement("c", "lineChart", ChartUri);
                WriteValue(xml, "grouping", Grouping(stacked));
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteValue(xml, "marker", "1");
                WriteAxisIds(xml, group.Secondary);
                xml.WriteEndElement();

                break;

            case ChartType.Area:
                xml.WriteStartElement("c", "areaChart", ChartUri);
                WriteValue(xml, "grouping", Grouping(stacked));
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteAxisIds(xml, group.Secondary);
                xml.WriteEndElement();

                break;

            case ChartType.Pie:
                xml.WriteStartElement("c", "pieChart", ChartUri);
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteValue(xml, "firstSliceAng", DrawingValues.Text(chart.FirstSliceAngle));
                xml.WriteEndElement();

                break;

            case ChartType.Doughnut:
                xml.WriteStartElement("c", "doughnutChart", ChartUri);
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteValue(xml, "firstSliceAng", DrawingValues.Text(chart.FirstSliceAngle));
                WriteValue(xml, "holeSize", DrawingValues.Text(chart.HoleSize ?? 50));
                xml.WriteEndElement();

                break;

            default:
                xml.WriteStartElement("c", "scatterChart", ChartUri);
                WriteValue(xml, "scatterStyle", "lineMarker");
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteAxisIds(xml, group.Secondary);
                xml.WriteEndElement();

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
    private static void WriteSeries(XmlWriter xml, ChartSpace chart, SeriesGroup group)
    {
        ChartType type = group.Type;
        bool pie = type is ChartType.Pie or ChartType.Doughnut;
        bool lines = type is ChartType.Line or ChartType.Scatter;
        bool bars = type is ChartType.Column or ChartType.Bar;

        foreach ((int index, ChartSpaceSeries series) in group.Members)
        {
            xml.WriteStartElement("c", "ser", ChartUri);
            WriteValue(xml, "idx", DrawingValues.Text(index));
            WriteValue(xml, "order", DrawingValues.Text(index));

            if (series.NameFormula is not null)
            {
                xml.WriteStartElement("c", "tx", ChartUri);
                xml.WriteStartElement("c", "strRef", ChartUri);
                xml.WriteElementString("c", "f", ChartUri, series.NameFormula);

                if (series.Name is not null)
                {
                    WritePoints(xml, true, "strCache", null, 1, [series.Name]);
                }

                xml.WriteEndElement();
                xml.WriteEndElement();
            }
            else if (series.Name is not null)
            {
                xml.WriteStartElement("c", "tx", ChartUri);
                xml.WriteElementString("c", "v", ChartUri, series.Name);
                xml.WriteEndElement();
            }

            // A pie colours each slice, so one colour for the whole series would hide them,
            // unless the chart turns that off.
            string? color = series.Color;

            if (pie == true && (chart.VaryColors ?? true) == true)
            {
                color = null;
            }

            WriteSeriesShape(xml, type, series, color);

            if (lines == true)
            {
                WriteMarker(xml, series, color ?? series.Outline?.Color);
            }

            if (bars == true)
            {
                WriteValue(xml, "invertIfNegative", "0");
            }

            for (int point = 0; point < series.PointColors.Count; point++)
            {
                if (series.PointColors[point] is string pointColor)
                {
                    WritePoint(xml, type, point, pointColor, series);
                }
            }

            if ((series.DataLabels ?? chart.DataLabels) is DataLabels labels)
            {
                WriteDataLabels(xml, labels);
            }

            if (type == ChartType.Scatter)
            {
                if (series.Categories is not null)
                {
                    WriteReference(xml, "xVal", true, series.Categories);
                }

                WriteReference(xml, "yVal", true, series.Values);
            }
            else
            {
                if (series.Categories is not null)
                {
                    WriteReference(xml, "cat", false, series.Categories);
                }

                WriteReference(xml, "val", true, series.Values);
            }

            if (lines == true)
            {
                WriteValue(xml, "smooth", Flag(series.Smooth));
            }

            xml.WriteEndElement();
        }
    }

    /// <summary>The series' c:spPr: the fill, line and shadow its type uses, when any is set.</summary>
    private static void WriteSeriesShape(XmlWriter xml, ChartType type, ChartSpaceSeries series, string? color)
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

        xml.WriteStartElement("c", "spPr", ChartUri);

        if (fill is not null)
        {
            DrawingWriter.WriteFill(xml, fill);
        }

        if (outline is not null)
        {
            if (type == ChartType.Scatter && outline.Color is null)
            {
                // As Excel writes a scatter series without lines.
                xml.WriteStartElement("a", "ln", DrawingUri);
                xml.WriteAttributeString(null, "w", null, "19050");
                xml.WriteElementString("a", "noFill", DrawingUri, string.Empty);
                xml.WriteEndElement();
            }
            else
            {
                DrawingWriter.WriteOutline(xml, outline);
            }
        }

        if (series.Shadow is not null)
        {
            DrawingWriter.WriteEffects(xml, series.Shadow, null, null);
        }

        xml.WriteEndElement();
    }

    private static void WriteMarker(XmlWriter xml, ChartSpaceSeries series, string? color)
    {
        if (series.Marker is null && series.MarkerSize is null && color is null)
        {
            return;
        }

        xml.WriteStartElement("c", "marker", ChartUri);
        MarkerStyle style = series.Marker ?? MarkerStyle.Circle;
        WriteValue(xml, "symbol", MarkerName(style));

        if (style != MarkerStyle.None)
        {
            WriteValue(xml, "size", DrawingValues.Text(series.MarkerSize ?? 5));

            if (color is not null)
            {
                xml.WriteStartElement("c", "spPr", ChartUri);
                DrawingWriter.WriteFill(xml, Fill.Solid(color));
                DrawingWriter.WriteOutline(xml, new Outline(color, 0.75));
                xml.WriteEndElement();
            }
        }

        xml.WriteEndElement();
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
    private static void WritePoint(XmlWriter xml, ChartType type, int point, string color, ChartSpaceSeries series)
    {
        xml.WriteStartElement("c", "dPt", ChartUri);
        WriteValue(xml, "idx", DrawingValues.Text(point));

        if (type is ChartType.Column or ChartType.Bar)
        {
            WriteValue(xml, "invertIfNegative", "0");
        }

        WriteValue(xml, "bubble3D", "0");
        xml.WriteStartElement("c", "spPr", ChartUri);

        if (type is ChartType.Line or ChartType.Scatter)
        {
            double width = series.Outline?.Width ?? 2.25;
            DrawingWriter.WriteOutline(xml, new Outline(color, width));
        }
        else
        {
            DrawingWriter.WriteFill(xml, Fill.Solid(color));

            if (series.Outline is not null)
            {
                DrawingWriter.WriteOutline(xml, series.Outline);
            }
        }

        if (series.Shadow is not null)
        {
            DrawingWriter.WriteEffects(xml, series.Shadow, null, null);
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    private static void WriteDataLabels(XmlWriter xml, DataLabels labels)
    {
        xml.WriteStartElement("c", "dLbls", ChartUri);

        if (labels.NumberFormat is not null)
        {
            xml.WriteStartElement("c", "numFmt", ChartUri);
            xml.WriteAttributeString(null, "formatCode", null, labels.NumberFormat);
            xml.WriteAttributeString(null, "sourceLinked", null, "0");
            xml.WriteEndElement();
        }

        // Excel's own labels have no box around them.
        xml.WriteStartElement("c", "spPr", ChartUri);
        xml.WriteElementString("a", "noFill", DrawingUri, string.Empty);
        xml.WriteStartElement("a", "ln", DrawingUri);
        xml.WriteElementString("a", "noFill", DrawingUri, string.Empty);
        xml.WriteEndElement();
        xml.WriteEndElement();
        WriteTextProperties(xml, labels.Font);

        if (labels.Position != DataLabelPosition.Default)
        {
            WriteValue(xml, "dLblPos", PositionName(labels.Position));
        }

        WriteValue(xml, "showLegendKey", "0");
        WriteValue(xml, "showVal", Flag(labels.ShowValue));
        WriteValue(xml, "showCatName", Flag(labels.ShowCategory));
        WriteValue(xml, "showSerName", Flag(labels.ShowSeriesName));
        WriteValue(xml, "showPercent", Flag(labels.ShowPercent));
        WriteValue(xml, "showBubbleSize", "0");
        xml.WriteEndElement();
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

    private static void WriteAxes(XmlWriter xml, ChartSpace chart, bool secondary)
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
            WriteValueAxis(xml, new AxisPlace(CategoryAxisId, ValueAxisId, "b", "midCat", false), chart.CategoryAxisTitle, categoryAxis, false);
            WriteValueAxis(xml, new AxisPlace(ValueAxisId, CategoryAxisId, "l", "midCat", false), chart.ValueAxisTitle, valueAxis, true);

            return;
        }

        WriteCategoryAxis(xml, CategoryAxisId, ValueAxisId, categoryPosition, chart.CategoryAxisTitle, categoryAxis);
        WriteValueAxis(xml, new AxisPlace(ValueAxisId, CategoryAxisId, valuePosition, "between", false), chart.ValueAxisTitle, valueAxis, true);

        if (secondary == true)
        {
            // The secondary series share the categories, so their own category axis is hidden;
            // their value axis sits on the right, crossing at the far end.
            WriteCategoryAxis(xml, SecondaryCategoryAxisId, SecondaryValueAxisId, "b", null, new ChartAxis { Visible = false });
            WriteValueAxis(xml, new AxisPlace(SecondaryValueAxisId, SecondaryCategoryAxisId, "r", "between", true), null, chart.SecondaryValueAxis ?? new ChartAxis(), false);
        }
    }

    private static void WriteCategoryAxis(XmlWriter xml, string id, string crossing, string position, string? title, ChartAxis axis)
    {
        xml.WriteStartElement("c", "catAx", ChartUri);
        WriteValue(xml, "axId", id);
        WriteScaling(xml, axis, false);
        WriteValue(xml, "delete", Flag(axis.Visible == false));
        WriteValue(xml, "axPos", position);
        WriteGridlines(xml, axis, false);

        if (title is not null)
        {
            WriteTitle(xml, title, false, null);
        }

        if (axis.NumberFormat is not null)
        {
            WriteNumberFormat(xml, axis.NumberFormat, false);
        }

        WriteValue(xml, "majorTickMark", "out");
        WriteValue(xml, "minorTickMark", "none");
        WriteValue(xml, "tickLblPos", "nextTo");
        WriteAxisShape(xml, axis);
        WriteTextProperties(xml, axis.Font);
        WriteValue(xml, "crossAx", crossing);
        WriteValue(xml, "crosses", "autoZero");
        WriteValue(xml, "auto", "1");
        WriteValue(xml, "lblAlgn", "ctr");
        WriteValue(xml, "lblOffset", "100");
        WriteValue(xml, "noMultiLvlLbl", "0");
        xml.WriteEndElement();
    }

    /// <summary>Where a value axis sits: its id, the axis it crosses, its side, how it meets the categories, and whether it crosses at the far end.</summary>
    private readonly record struct AxisPlace(string Id, string Crossing, string Position, string CrossBetween, bool CrossesAtMaximum);

    private static void WriteValueAxis(XmlWriter xml, AxisPlace place, string? title, ChartAxis axis, bool gridlines)
    {
        xml.WriteStartElement("c", "valAx", ChartUri);
        WriteValue(xml, "axId", place.Id);
        WriteScaling(xml, axis, true);
        WriteValue(xml, "delete", Flag(axis.Visible == false));
        WriteValue(xml, "axPos", place.Position);
        WriteGridlines(xml, axis, gridlines);

        if (title is not null)
        {
            WriteTitle(xml, title, place.Position is "l" or "r", null);
        }

        WriteNumberFormat(xml, axis.NumberFormat ?? "General", axis.NumberFormat is null);
        WriteValue(xml, "majorTickMark", "out");
        WriteValue(xml, "minorTickMark", "none");
        WriteValue(xml, "tickLblPos", "nextTo");
        WriteAxisShape(xml, axis);
        WriteTextProperties(xml, axis.Font);
        WriteValue(xml, "crossAx", place.Crossing);

        if (place.CrossesAtMaximum == true)
        {
            WriteValue(xml, "crosses", "max");
        }
        else
        {
            WriteValue(xml, "crosses", "autoZero");
        }

        WriteValue(xml, "crossBetween", place.CrossBetween);

        if (axis.MajorUnit is double unit)
        {
            WriteValue(xml, "majorUnit", DrawingValues.Text(unit));
        }

        xml.WriteEndElement();
    }

    private static void WriteScaling(XmlWriter xml, ChartAxis axis, bool values)
    {
        xml.WriteStartElement("c", "scaling", ChartUri);
        WriteValue(xml, "orientation", "minMax");

        if (values == true && axis.Maximum is double maximum)
        {
            WriteValue(xml, "max", DrawingValues.Text(maximum));
        }

        if (values == true && axis.Minimum is double minimum)
        {
            WriteValue(xml, "min", DrawingValues.Text(minimum));
        }

        xml.WriteEndElement();
    }

    private static void WriteGridlines(XmlWriter xml, ChartAxis axis, bool byDefault)
    {
        if ((axis.Gridlines ?? byDefault) == false)
        {
            return;
        }

        xml.WriteStartElement("c", "majorGridlines", ChartUri);

        if (axis.GridlineOutline is not null)
        {
            xml.WriteStartElement("c", "spPr", ChartUri);
            DrawingWriter.WriteOutline(xml, axis.GridlineOutline);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private static void WriteAxisShape(XmlWriter xml, ChartAxis axis)
    {
        if (axis.Line is null)
        {
            return;
        }

        xml.WriteStartElement("c", "spPr", ChartUri);
        DrawingWriter.WriteOutline(xml, axis.Line);
        xml.WriteEndElement();
    }

    private static void WriteNumberFormat(XmlWriter xml, string format, bool sourceLinked)
    {
        xml.WriteStartElement("c", "numFmt", ChartUri);
        xml.WriteAttributeString(null, "formatCode", null, format);
        xml.WriteAttributeString(null, "sourceLinked", null, Flag(sourceLinked));
        xml.WriteEndElement();
    }

    private static void WriteAxisIds(XmlWriter xml, bool secondary)
    {
        if (secondary == true)
        {
            WriteValue(xml, "axId", SecondaryCategoryAxisId);
            WriteValue(xml, "axId", SecondaryValueAxisId);

            return;
        }

        WriteValue(xml, "axId", CategoryAxisId);
        WriteValue(xml, "axId", ValueAxisId);
    }

    /// <summary>A title with one run of text. <paramref name="vertical"/> turns it to read upwards, as Excel does beside a vertical axis.</summary>
    private static void WriteTitle(XmlWriter xml, string text, bool vertical, TextFont? font)
    {
        xml.WriteStartElement("c", "title", ChartUri);
        xml.WriteStartElement("c", "tx", ChartUri);
        xml.WriteStartElement("c", "rich", ChartUri);
        xml.WriteStartElement("a", "bodyPr", DrawingUri);

        if (vertical == true)
        {
            xml.WriteAttributeString(null, "rot", null, "-5400000");
            xml.WriteAttributeString(null, "vert", null, "horz");
        }

        xml.WriteEndElement();
        xml.WriteElementString("a", "lstStyle", DrawingUri, string.Empty);
        xml.WriteStartElement("a", "p", DrawingUri);

        if (font is not null)
        {
            xml.WriteStartElement("a", "pPr", DrawingUri);
            DrawingWriter.WriteRunProperties(xml, "defRPr", font);
            xml.WriteEndElement();
        }

        xml.WriteStartElement("a", "r", DrawingUri);

        if (font is not null)
        {
            DrawingWriter.WriteRunProperties(xml, "rPr", font);
        }

        xml.WriteElementString("a", "t", DrawingUri, text);
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
        WriteValue(xml, "overlay", "0");
        xml.WriteEndElement();
    }

    /// <summary>A c:txPr: the default look of the text of the element it is in.</summary>
    private static void WriteTextProperties(XmlWriter xml, TextFont? font)
    {
        if (font is null)
        {
            return;
        }

        xml.WriteStartElement("c", "txPr", ChartUri);
        xml.WriteElementString("a", "bodyPr", DrawingUri, string.Empty);
        xml.WriteElementString("a", "lstStyle", DrawingUri, string.Empty);
        xml.WriteStartElement("a", "p", DrawingUri);
        xml.WriteStartElement("a", "pPr", DrawingUri);
        DrawingWriter.WriteRunProperties(xml, "defRPr", font);
        xml.WriteEndElement();
        xml.WriteStartElement("a", "endParaRPr", DrawingUri);
        xml.WriteAttributeString(null, "lang", null, "en-US");
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    /// <summary>
    /// A cat, val, xVal or yVal: a numRef or strRef with the formula and the stored values, or a
    /// numLit or strLit with values alone. <paramref name="numbers"/> is the kind the element
    /// takes when the data does not say; stored texts or numbers decide otherwise.
    /// </summary>
    private static void WriteReference(XmlWriter xml, string element, bool numbers, ChartData data)
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

        xml.WriteStartElement("c", element, ChartUri);
        xml.WriteStartElement("c", kind, ChartUri);

        if (data.Formula is not null)
        {
            xml.WriteElementString("c", "f", ChartUri, data.Formula);
        }

        if (numbers == true && data.Numbers is not null)
        {
            WritePoints(xml, data.Formula is not null, "numCache", data.FormatCode ?? "General", data.Numbers.Count, Numbers(data.Numbers));
        }
        else if (numbers == false && data.Texts is not null)
        {
            WritePoints(xml, data.Formula is not null, "strCache", null, data.Texts.Count, data.Texts);
        }
        else if (data.Formula is null)
        {
            // A literal without values still says it has none.
            WriteValue(xml, "ptCount", "0");
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    /// <summary>A cache (or, for a literal, the points alone): the format code of numbers, the count, and a c:pt per point that has a value.</summary>
    private static void WritePoints(XmlWriter xml, bool cache, string cacheName, string? formatCode, int count, IEnumerable<string?> points)
    {
        if (cache == true)
        {
            xml.WriteStartElement("c", cacheName, ChartUri);
        }

        if (formatCode is not null)
        {
            xml.WriteElementString("c", "formatCode", ChartUri, formatCode);
        }

        WriteValue(xml, "ptCount", DrawingValues.Text(count));
        int index = 0;

        foreach (string? point in points)
        {
            if (point is not null)
            {
                xml.WriteStartElement("c", "pt", ChartUri);
                xml.WriteAttributeString(null, "idx", null, DrawingValues.Text(index));
                xml.WriteElementString("c", "v", ChartUri, point);
                xml.WriteEndElement();
            }

            index++;
        }

        if (cache == true)
        {
            xml.WriteEndElement();
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

    private static void WriteValue(XmlWriter xml, string name, string value)
    {
        xml.WriteStartElement("c", name, ChartUri);
        xml.WriteAttributeString(null, "val", null, value);
        xml.WriteEndElement();
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
