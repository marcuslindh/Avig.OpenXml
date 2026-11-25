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
    private const string SeriesAxisId = "500000005";

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

        if (chart.Style is int style)
        {
            WriteValue(xml, "style", DrawingValues.Text(style));
        }

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

        WriteView3D(xml, chart);
        xml.WriteStartElement("c", "plotArea", ChartUri);
        xml.WriteElementString("c", "layout", ChartUri, string.Empty);
        List<SeriesGroup> groups = Groups(chart);

        foreach (SeriesGroup group in groups)
        {
            WriteGroup(xml, chart, group);
        }

        WriteAxes(xml, chart, groups.Exists(group => group.Secondary == true));

        if (chart.DataTable is ChartDataTable table && chart.HasAxes == true)
        {
            WriteDataTable(xml, table);
        }

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

        if (chart.ExternalData is string externalData)
        {
            xml.WriteStartElement("c", "externalData", ChartUri);
            xml.WriteAttributeString("r", "id", OpcNamespaces.Relationships, externalData);
            WriteValue(xml, "autoUpdate", "0");
            xml.WriteEndElement();
        }

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

    /// <summary>The grouping a plot is written with: the chart's for its own type, standard (or clustered) for series drawn as another type.</summary>
    private static ChartGrouping GroupingOf(ChartSpace chart, SeriesGroup group)
    {
        if (group.Type == chart.Type)
        {
            return chart.Grouping;
        }

        return ChartGrouping.Default;
    }

    private static string GroupingName(ChartGrouping grouping, bool bars)
    {
        return grouping switch
        {
            ChartGrouping.Stacked => "stacked",
            ChartGrouping.PercentStacked => "percentStacked",
            ChartGrouping.Depth => "standard",
            _ when bars == true => "clustered",
            _ => "standard",
        };
    }

    private static void WriteGroup(XmlWriter xml, ChartSpace chart, SeriesGroup group)
    {
        ChartGrouping grouping = GroupingOf(chart, group);
        bool stacked = grouping is ChartGrouping.Stacked or ChartGrouping.PercentStacked;
        bool pie = group.Type is ChartType.Pie or ChartType.Doughnut or ChartType.PieOfPie or ChartType.BarOfPie;
        bool threeD = chart.ThreeD is not null;
        string varyColors = Flag(chart.VaryColors ?? pie);

        switch (group.Type)
        {
            case ChartType.Column or ChartType.Bar:
                xml.WriteStartElement("c", Pick(threeD, "bar3DChart", "barChart"), ChartUri);
                WriteValue(xml, "barDir", Pick(group.Type == ChartType.Bar, "bar", "col"));
                WriteValue(xml, "grouping", GroupingName(grouping, true));
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteValue(xml, "gapWidth", DrawingValues.Text(chart.GapWidth ?? 150));

                if (threeD == true)
                {
                    if (grouping == ChartGrouping.Depth)
                    {
                        WriteGapDepth(xml, chart);
                    }

                    WriteValue(xml, "shape", BarShapeName(chart.BarShape));
                    WriteAxisIds(xml, false, grouping == ChartGrouping.Depth);
                    xml.WriteEndElement();

                    break;
                }

                if (chart.Overlap is int overlap)
                {
                    WriteValue(xml, "overlap", DrawingValues.Text(overlap));
                }
                else if (stacked == true)
                {
                    WriteValue(xml, "overlap", "100");
                }

                if (stacked == true && chart.SeriesLines == true)
                {
                    xml.WriteElementString("c", "serLines", ChartUri, string.Empty);
                }

                WriteAxisIds(xml, group.Secondary);
                xml.WriteEndElement();

                break;

            case ChartType.Line when threeD == true:
                xml.WriteStartElement("c", "line3DChart", ChartUri);
                WriteValue(xml, "grouping", "standard");
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteLines(xml, chart, group);
                WriteGapDepth(xml, chart);
                WriteAxisIds(xml, false, true);
                xml.WriteEndElement();

                break;

            case ChartType.Line:
                xml.WriteStartElement("c", "lineChart", ChartUri);
                WriteValue(xml, "grouping", GroupingName(grouping, false));
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteLines(xml, chart, group);
                WriteValue(xml, "marker", "1");
                WriteAxisIds(xml, group.Secondary);
                xml.WriteEndElement();

                break;

            case ChartType.Stock:
                xml.WriteStartElement("c", "stockChart", ChartUri);
                WriteSeries(xml, chart, group);
                WriteLines(xml, chart, group);
                WriteAxisIds(xml, group.Secondary);
                xml.WriteEndElement();

                break;

            case ChartType.Area:
                xml.WriteStartElement("c", Pick(threeD, "area3DChart", "areaChart"), ChartUri);
                WriteValue(xml, "grouping", GroupingName(grouping, false));
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);

                if (chart.DropLines == true)
                {
                    xml.WriteElementString("c", "dropLines", ChartUri, string.Empty);
                }

                if (threeD == true)
                {
                    WriteGapDepth(xml, chart);
                    WriteAxisIds(xml, false, grouping == ChartGrouping.Depth);
                }
                else
                {
                    WriteAxisIds(xml, group.Secondary);
                }

                xml.WriteEndElement();

                break;

            case ChartType.Pie when threeD == true:
                xml.WriteStartElement("c", "pie3DChart", ChartUri);
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
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

            case ChartType.PieOfPie or ChartType.BarOfPie:
                xml.WriteStartElement("c", "ofPieChart", ChartUri);
                WriteValue(xml, "ofPieType", Pick(group.Type == ChartType.BarOfPie, "bar", "pie"));
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteValue(xml, "gapWidth", DrawingValues.Text(chart.GapWidth ?? 100));

                if (chart.Split != ChartSplit.Auto)
                {
                    WriteValue(xml, "splitType", SplitName(chart.Split));
                }

                if (chart.SplitValue is double split)
                {
                    WriteValue(xml, "splitPos", DrawingValues.Text(split));
                }

                WriteValue(xml, "secondPieSize", DrawingValues.Text(chart.SecondPlotSize ?? 75));

                if ((chart.SeriesLines ?? true) == true)
                {
                    xml.WriteElementString("c", "serLines", ChartUri, string.Empty);
                }

                xml.WriteEndElement();

                break;

            case ChartType.Radar:
                xml.WriteStartElement("c", "radarChart", ChartUri);
                WriteValue(xml, "radarStyle", Pick(chart.RadarStyle == RadarStyle.Filled, "filled", "marker"));
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteAxisIds(xml, group.Secondary);
                xml.WriteEndElement();

                break;

            case ChartType.Bubble:
                xml.WriteStartElement("c", "bubbleChart", ChartUri);
                WriteValue(xml, "varyColors", varyColors);
                WriteSeries(xml, chart, group);
                WriteValue(xml, "bubbleScale", DrawingValues.Text(chart.BubbleScale ?? 100));
                WriteValue(xml, "showNegBubbles", Flag(chart.ShowNegativeBubbles));
                WriteAxisIds(xml, group.Secondary);
                xml.WriteEndElement();

                break;

            case ChartType.Surface:
                xml.WriteStartElement("c", Pick(threeD, "surface3DChart", "surfaceChart"), ChartUri);
                WriteValue(xml, "wireframe", Flag(chart.Wireframe));
                WriteSeries(xml, chart, group);
                WriteAxisIds(xml, false, true);
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

    private static string Pick(bool condition, string whenTrue, string whenFalse)
    {
        if (condition == true)
        {
            return whenTrue;
        }

        return whenFalse;
    }

    /// <summary>The drop lines, high-low lines and up-down bars of a line or stock plot, in the schema's order.</summary>
    private static void WriteLines(XmlWriter xml, ChartSpace chart, SeriesGroup group)
    {
        bool stock = group.Type == ChartType.Stock;

        if (chart.DropLines == true)
        {
            xml.WriteElementString("c", "dropLines", ChartUri, string.Empty);
        }

        if (chart.ThreeD is not null)
        {
            return;
        }

        if ((chart.HighLowLines ?? stock) == true)
        {
            xml.WriteElementString("c", "hiLowLines", ChartUri, string.Empty);
        }

        if ((chart.UpDownBars ?? (stock == true && group.Members.Count == 4)) == true)
        {
            xml.WriteStartElement("c", "upDownBars", ChartUri);
            WriteValue(xml, "gapWidth", "150");
            xml.WriteElementString("c", "upBars", ChartUri, string.Empty);
            xml.WriteElementString("c", "downBars", ChartUri, string.Empty);
            xml.WriteEndElement();
        }
    }

    private static void WriteGapDepth(XmlWriter xml, ChartSpace chart)
    {
        if (chart.ThreeD is ChartView3D view && view.GapDepth != 150)
        {
            WriteValue(xml, "gapDepth", DrawingValues.Text(view.GapDepth));
        }
    }

    /// <summary>The c:view3D of a 3-D chart, and of a contour (a surface seen from straight above).</summary>
    private static void WriteView3D(XmlWriter xml, ChartSpace chart)
    {
        ChartView3D? view = chart.ThreeD;

        if (view is null && chart.Type == ChartType.Surface)
        {
            view = new ChartView3D { RotationX = 90, RotationY = 0, RightAngleAxes = false, Perspective = 0 };
        }

        if (view is null)
        {
            return;
        }

        xml.WriteStartElement("c", "view3D", ChartUri);
        WriteValue(xml, "rotX", DrawingValues.Text(view.RotationX));
        WriteValue(xml, "rotY", DrawingValues.Text(view.RotationY));

        if (view.DepthPercent != 100)
        {
            WriteValue(xml, "depthPercent", DrawingValues.Text(view.DepthPercent));
        }

        WriteValue(xml, "rAngAx", Flag(view.RightAngleAxes));

        if (view.RightAngleAxes == false)
        {
            WriteValue(xml, "perspective", DrawingValues.Text(view.Perspective));
        }

        xml.WriteEndElement();
    }

    private static void WriteDataTable(XmlWriter xml, ChartDataTable table)
    {
        xml.WriteStartElement("c", "dTable", ChartUri);
        WriteValue(xml, "showHorzBorder", Flag(table.HorizontalBorders));
        WriteValue(xml, "showVertBorder", Flag(table.VerticalBorders));
        WriteValue(xml, "showOutline", Flag(table.Outline));
        WriteValue(xml, "showKeys", Flag(table.ShowLegendKeys));
        xml.WriteEndElement();
    }

    private static string BarShapeName(BarShape shape)
    {
        return shape switch
        {
            BarShape.Cylinder => "cylinder",
            BarShape.Cone => "cone",
            BarShape.ConeToMax => "coneToMax",
            BarShape.Pyramid => "pyramid",
            BarShape.PyramidToMax => "pyramidToMax",
            _ => "box",
        };
    }

    private static string SplitName(ChartSplit split)
    {
        return split switch
        {
            ChartSplit.Position => "pos",
            ChartSplit.Value => "val",
            ChartSplit.Percent => "percent",
            _ => "auto",
        };
    }

    /// <summary>The series, each with its children in the order its type's schema gives.</summary>
    private static void WriteSeries(XmlWriter xml, ChartSpace chart, SeriesGroup group)
    {
        ChartType type = group.Type;
        bool threeD = chart.ThreeD is not null;
        bool pie = type is ChartType.Pie or ChartType.Doughnut or ChartType.PieOfPie or ChartType.BarOfPie;
        bool lines = type is ChartType.Line or ChartType.Scatter or ChartType.Stock;
        bool markers = lines == true || (type == ChartType.Radar && chart.RadarStyle == RadarStyle.Marker);
        bool bars = type is ChartType.Column or ChartType.Bar or ChartType.Bubble;
        bool xy = type is ChartType.Scatter or ChartType.Bubble;
        bool extras = threeD == false && type is ChartType.Column or ChartType.Bar or ChartType.Line or ChartType.Area or ChartType.Scatter or ChartType.Bubble;
        bool points = type != ChartType.Surface;

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

            WriteSeriesShape(xml, chart, type, series, color);

            if (markers == true)
            {
                WriteMarker(xml, series, color ?? series.Outline?.Color, type == ChartType.Stock);
            }

            if (bars == true)
            {
                WriteValue(xml, "invertIfNegative", "0");
            }

            if (pie == true && series.Explosion is int explosion)
            {
                WriteValue(xml, "explosion", DrawingValues.Text(explosion));
            }

            for (int point = 0; point < series.PointColors.Count && points == true; point++)
            {
                if (series.PointColors[point] is string pointColor)
                {
                    WritePoint(xml, chart, type, point, pointColor, series);
                }
            }

            if ((series.DataLabels ?? chart.DataLabels) is DataLabels labels && points == true)
            {
                WriteDataLabels(xml, labels);
            }

            if (extras == true)
            {
                foreach (Trendline trendline in series.Trendlines)
                {
                    WriteTrendline(xml, trendline);
                }

                foreach (ErrorBars errorBars in series.ErrorBars.OrderBy(item => item.Direction == ErrorBarDirection.Y))
                {
                    WriteErrorBars(xml, errorBars, xy);
                }
            }

            if (xy == true)
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

            if (type == ChartType.Bubble)
            {
                if (series.BubbleSizes is not null)
                {
                    WriteReference(xml, "bubbleSize", true, series.BubbleSizes);
                }

                WriteValue(xml, "bubble3D", Flag(chart.Bubble3D));
            }

            if (lines == true)
            {
                WriteValue(xml, "smooth", Flag(series.Smooth));
            }

            xml.WriteEndElement();
        }
    }

    /// <summary>True for the types whose series are drawn as lines rather than filled: line, scatter, stock and radar with markers.</summary>
    private static bool IsLineType(ChartSpace chart, ChartType type)
    {
        return type is ChartType.Line or ChartType.Scatter or ChartType.Stock || (type == ChartType.Radar && chart.RadarStyle == RadarStyle.Marker);
    }

    /// <summary>The series' c:spPr: the fill, line and shadow its type uses, when any is set.</summary>
    private static void WriteSeriesShape(XmlWriter xml, ChartSpace chart, ChartType type, ChartSpaceSeries series, string? color)
    {
        Fill? fill = null;
        Outline? outline = series.Outline;
        bool pointsAlone = type == ChartType.Stock || (type == ChartType.Scatter && chart.ScatterLines == false);

        if (IsLineType(chart, type) == true)
        {
            if (outline is null && color is not null)
            {
                outline = new Outline(color, 2.25);
            }

            // A plain scatter chart (and a stock chart) shows points without lines between them, unless asked.
            if (pointsAlone == true && series.Outline is null && series.Smooth == false)
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

        // Lines in the application's colour: a line without a fill of its own.
        bool automaticLine = type == ChartType.Scatter && chart.ScatterLines == true && outline is null;

        if (fill is null && outline is null && series.Shadow is null && automaticLine == false)
        {
            return;
        }

        xml.WriteStartElement("c", "spPr", ChartUri);

        if (fill is not null)
        {
            DrawingWriter.WriteFill(xml, fill);
        }

        if (automaticLine == true)
        {
            xml.WriteStartElement("a", "ln", DrawingUri);
            xml.WriteAttributeString(null, "w", null, "28575");
            xml.WriteAttributeString(null, "cap", null, "rnd");
            xml.WriteEndElement();
        }

        if (outline is not null)
        {
            if (pointsAlone == true && outline.Color is null)
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

    private static void WriteMarker(XmlWriter xml, ChartSpaceSeries series, string? color, bool stock)
    {
        if (series.Marker is null && series.MarkerSize is null && color is null && stock == false)
        {
            return;
        }

        xml.WriteStartElement("c", "marker", ChartUri);
        MarkerStyle style = series.Marker ?? MarkerStyle.Circle;

        // A stock chart's points are not marked unless asked: the high-low lines show them.
        if (stock == true && series.Marker is null)
        {
            style = MarkerStyle.None;
        }

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
    private static void WritePoint(XmlWriter xml, ChartSpace chart, ChartType type, int point, string color, ChartSpaceSeries series)
    {
        xml.WriteStartElement("c", "dPt", ChartUri);
        WriteValue(xml, "idx", DrawingValues.Text(point));

        if (type is ChartType.Column or ChartType.Bar or ChartType.Bubble)
        {
            WriteValue(xml, "invertIfNegative", "0");
        }

        WriteValue(xml, "bubble3D", "0");
        xml.WriteStartElement("c", "spPr", ChartUri);

        if (IsLineType(chart, type) == true)
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

    private static void WriteTrendline(XmlWriter xml, Trendline trendline)
    {
        xml.WriteStartElement("c", "trendline", ChartUri);

        if (trendline.Name is not null)
        {
            xml.WriteElementString("c", "name", ChartUri, trendline.Name);
        }

        if (trendline.Outline is not null)
        {
            xml.WriteStartElement("c", "spPr", ChartUri);
            DrawingWriter.WriteOutline(xml, trendline.Outline);
            xml.WriteEndElement();
        }

        WriteValue(xml, "trendlineType", TrendlineName(trendline.Kind));

        if (trendline.Kind == TrendlineKind.Polynomial)
        {
            WriteValue(xml, "order", DrawingValues.Text(trendline.Order));
        }

        if (trendline.Kind == TrendlineKind.MovingAverage)
        {
            WriteValue(xml, "period", DrawingValues.Text(trendline.Period));
        }

        if (trendline.Forward > 0)
        {
            WriteValue(xml, "forward", DrawingValues.Text(trendline.Forward));
        }

        if (trendline.Backward > 0)
        {
            WriteValue(xml, "backward", DrawingValues.Text(trendline.Backward));
        }

        if (trendline.Intercept is double intercept)
        {
            WriteValue(xml, "intercept", DrawingValues.Text(intercept));
        }

        WriteValue(xml, "dispRSqr", Flag(trendline.DisplayRSquared));
        WriteValue(xml, "dispEq", Flag(trendline.DisplayEquation));
        xml.WriteEndElement();
    }

    private static string TrendlineName(TrendlineKind kind)
    {
        return kind switch
        {
            TrendlineKind.Exponential => "exp",
            TrendlineKind.Logarithmic => "log",
            TrendlineKind.MovingAverage => "movingAvg",
            TrendlineKind.Polynomial => "poly",
            TrendlineKind.Power => "power",
            _ => "linear",
        };
    }

    private static void WriteErrorBars(XmlWriter xml, ErrorBars bars, bool xy)
    {
        xml.WriteStartElement("c", "errBars", ChartUri);

        // The direction is only written where there are two.
        if (xy == true)
        {
            WriteValue(xml, "errDir", Pick(bars.Direction == ErrorBarDirection.X, "x", "y"));
        }

        string side = bars.Side switch
        {
            ErrorBarSide.Plus => "plus",
            ErrorBarSide.Minus => "minus",
            _ => "both",
        };

        string amount = bars.Amount switch
        {
            ErrorBarAmount.Percentage => "percentage",
            ErrorBarAmount.StandardDeviation => "stdDev",
            ErrorBarAmount.StandardError => "stdErr",
            _ => "fixedVal",
        };

        WriteValue(xml, "errBarType", side);
        WriteValue(xml, "errValType", amount);
        WriteValue(xml, "noEndCap", Flag(bars.NoEndCap));

        if (bars.Amount != ErrorBarAmount.StandardError)
        {
            WriteValue(xml, "val", DrawingValues.Text(bars.Value));
        }

        if (bars.Outline is not null)
        {
            xml.WriteStartElement("c", "spPr", ChartUri);
            DrawingWriter.WriteOutline(xml, bars.Outline);
            xml.WriteEndElement();
        }

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

    /// <summary>True when the chart has a depth axis (c:serAx): a surface, a 3-D line chart, and a 3-D chart with its series in depth.</summary>
    internal static bool HasSeriesAxis(ChartSpace chart)
    {
        return chart.Type == ChartType.Surface
            || (chart.ThreeD is not null && (chart.Type == ChartType.Line || (chart.Grouping == ChartGrouping.Depth && chart.Type is ChartType.Column or ChartType.Bar or ChartType.Area)));
    }

    private static void WriteAxes(XmlWriter xml, ChartSpace chart, bool secondary)
    {
        if (chart.HasAxes == false)
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

        if (chart.Type is ChartType.Scatter or ChartType.Bubble)
        {
            // Two value axes: x along the bottom, y up the side.
            WriteValueAxis(xml, new AxisPlace(CategoryAxisId, ValueAxisId, "b", "midCat", false), chart.CategoryAxisTitle, categoryAxis, false);
            WriteValueAxis(xml, new AxisPlace(ValueAxisId, CategoryAxisId, "l", "midCat", false), chart.ValueAxisTitle, valueAxis, true);

            return;
        }

        string crossBetween = Pick(chart.Type == ChartType.Surface, "midCat", "between");
        WriteCategoryAxis(xml, CategoryAxisId, ValueAxisId, categoryPosition, chart.CategoryAxisTitle, categoryAxis);
        WriteValueAxis(xml, new AxisPlace(ValueAxisId, CategoryAxisId, valuePosition, crossBetween, false), chart.ValueAxisTitle, valueAxis, true);

        if (secondary == true)
        {
            // The secondary series share the categories, so their own category axis is hidden;
            // their value axis sits on the right, crossing at the far end.
            WriteCategoryAxis(xml, SecondaryCategoryAxisId, SecondaryValueAxisId, "b", null, new ChartAxis { Visible = false });
            WriteValueAxis(xml, new AxisPlace(SecondaryValueAxisId, SecondaryCategoryAxisId, "r", "between", true), null, chart.SecondaryValueAxis ?? new ChartAxis(), false);
        }

        if (HasSeriesAxis(chart) == true)
        {
            WriteSeriesAxis(xml, chart.SeriesAxis ?? new ChartAxis());
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

    /// <summary>The depth axis of a 3-D chart, along the series.</summary>
    private static void WriteSeriesAxis(XmlWriter xml, ChartAxis axis)
    {
        xml.WriteStartElement("c", "serAx", ChartUri);
        WriteValue(xml, "axId", SeriesAxisId);
        WriteScaling(xml, axis, false);
        WriteValue(xml, "delete", Flag(axis.Visible == false));
        WriteValue(xml, "axPos", "b");
        WriteGridlines(xml, axis, false);
        WriteValue(xml, "majorTickMark", "out");
        WriteValue(xml, "minorTickMark", "none");
        WriteValue(xml, "tickLblPos", "nextTo");
        WriteAxisShape(xml, axis);
        WriteTextProperties(xml, axis.Font);
        WriteValue(xml, "crossAx", ValueAxisId);
        WriteValue(xml, "crosses", "autoZero");
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

        if (axis.MinorUnit is double minor)
        {
            WriteValue(xml, "minorUnit", DrawingValues.Text(minor));
        }

        xml.WriteEndElement();
    }

    private static void WriteScaling(XmlWriter xml, ChartAxis axis, bool values)
    {
        xml.WriteStartElement("c", "scaling", ChartUri);

        if (values == true && axis.LogBase is double logBase)
        {
            WriteValue(xml, "logBase", DrawingValues.Text(logBase));
        }

        WriteValue(xml, "orientation", Pick(axis.Reversed, "maxMin", "minMax"));

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
        if ((axis.Gridlines ?? byDefault) == true)
        {
            xml.WriteStartElement("c", "majorGridlines", ChartUri);

            if (axis.GridlineOutline is not null)
            {
                xml.WriteStartElement("c", "spPr", ChartUri);
                DrawingWriter.WriteOutline(xml, axis.GridlineOutline);
                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }

        if (axis.MinorGridlines == true)
        {
            xml.WriteStartElement("c", "minorGridlines", ChartUri);

            if (axis.MinorGridlineOutline is not null)
            {
                xml.WriteStartElement("c", "spPr", ChartUri);
                DrawingWriter.WriteOutline(xml, axis.MinorGridlineOutline);
                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }
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

    /// <summary>The three axis ids of a 3-D plot: the depth axis', or 0 when its series stand side by side, as PowerPoint writes them.</summary>
    private static void WriteAxisIds(XmlWriter xml, bool secondary, bool depth)
    {
        WriteAxisIds(xml, secondary);
        WriteValue(xml, "axId", Pick(depth, SeriesAxisId, "0"));
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
