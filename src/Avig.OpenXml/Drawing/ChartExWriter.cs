using System.Xml;

namespace Avig.OpenXml.Drawing;

/// <summary>
/// Writes a chart part of the newer chart engine (cx:chartSpace), in the order the schema
/// ([MS-ODRAWXML] 2.24) gives and as PowerPoint 365 writes it: the data first (cx:chartData,
/// one cx:data per series with the categories and the values), then the chart. cx:axisId,
/// cx:binSize and cx:binCount carry their value in a val attribute, as PowerPoint 365 reads them
/// (it does not open a part with the value as text, which is what the Open XML SDK's schema says).
/// </summary>
public static class ChartExWriter
{
    /// <summary>The namespace of the part (cx:).</summary>
    public const string Namespace = "http://schemas.microsoft.com/office/drawing/2014/chartex";

    /// <summary>The content type of the part.</summary>
    public const string ContentType = "application/vnd.ms-office.chartex+xml";

    /// <summary>The relationship from a slide, a drawing or a document to the part.</summary>
    public const string Relationship = "http://schemas.microsoft.com/office/2014/relationships/chartEx";

    /// <summary>The namespace an mc:Choice requires (as <c>cx1</c>) for a graphic frame that shows the part; region maps require <see cref="RegionMapNamespace"/>.</summary>
    public const string ChoiceNamespace = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";

    /// <summary>The namespace an mc:Choice requires (as <c>cx4</c>) for a filled map.</summary>
    public const string RegionMapNamespace = "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex";

    private const string DrawingUri = OpcNamespaces.DrawingMain;

    /// <summary>A complete chart part, from the XML declaration on. Call <see cref="ChartExSpace.Validate"/> first.</summary>
    /// <param name="xml">The writer of the part.</param>
    /// <param name="chart">The chart.</param>
    public static void WriteChartSpace(XmlWriter xml, ChartExSpace chart)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(chart);
        xml.WriteStartDocument(true);
        xml.WriteStartElement("cx", "chartSpace", Namespace);
        xml.WriteAttributeString("xmlns", "a", null, DrawingUri);
        xml.WriteAttributeString("xmlns", "r", null, OpcNamespaces.Relationships);
        xml.WriteStartElement("cx", "chartData", Namespace);

        if (chart.ExternalData is string externalData)
        {
            xml.WriteStartElement("cx", "externalData", Namespace);
            xml.WriteAttributeString("r", "id", OpcNamespaces.Relationships, externalData);
            xml.WriteAttributeString("cx", "autoUpdate", Namespace, "0");
            xml.WriteEndElement();
        }

        for (int index = 0; index < chart.Series.Count; index++)
        {
            WriteData(xml, chart, index);
        }

        xml.WriteEndElement();
        xml.WriteStartElement("cx", "chart", Namespace);

        if (chart.Title is not null)
        {
            xml.WriteStartElement("cx", "title", Namespace);
            xml.WriteAttributeString(null, "pos", null, "t");
            xml.WriteAttributeString(null, "align", null, "ctr");
            xml.WriteAttributeString(null, "overlay", null, "0");
            WriteText(xml, null, chart.Title);
            xml.WriteEndElement();
        }

        xml.WriteStartElement("cx", "plotArea", Namespace);
        xml.WriteStartElement("cx", "plotAreaRegion", Namespace);

        for (int index = 0; index < chart.Series.Count; index++)
        {
            WriteSeries(xml, chart, index);
        }

        if (chart.Type == ChartExType.Pareto)
        {
            // The cumulative percentage is a series of its own, owned by the columns, on a third axis.
            xml.WriteStartElement("cx", "series", Namespace);
            xml.WriteAttributeString(null, "layoutId", null, "paretoLine");
            xml.WriteAttributeString(null, "ownerIdx", null, "0");
            xml.WriteAttributeString(null, "uniqueId", null, UniqueId(255));
            Value(xml, "axisId", "2");
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        WriteAxes(xml, chart);
        xml.WriteEndElement();

        LegendPosition legend = chart.LegendOrDefault;

        if (legend != LegendPosition.None)
        {
            xml.WriteStartElement("cx", "legend", Namespace);
            xml.WriteAttributeString(null, "pos", null, DrawingNames.Legend(legend));
            xml.WriteAttributeString(null, "align", null, "ctr");
            xml.WriteAttributeString(null, "overlay", null, "0");
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndDocument();
    }

    /// <summary>The layout a type's series are written with.</summary>
    internal static string LayoutId(ChartExType type)
    {
        return type switch
        {
            ChartExType.Waterfall => "waterfall",
            ChartExType.Treemap => "treemap",
            ChartExType.Sunburst => "sunburst",
            ChartExType.BoxWhisker => "boxWhisker",
            ChartExType.Funnel => "funnel",
            ChartExType.RegionMap => "regionMap",
            _ => "clusteredColumn",
        };
    }

    /// <summary>The labels a type shows when none are given, as PowerPoint makes the chart; null for none.</summary>
    internal static DataLabels? DefaultLabels(ChartExType type)
    {
        DataLabels? labels = type switch
        {
            ChartExType.Treemap => new DataLabels { ShowValue = false, ShowCategory = true, Position = DataLabelPosition.InsideEnd },
            ChartExType.Sunburst => new DataLabels { ShowValue = false, ShowCategory = true, Position = DataLabelPosition.Center },
            ChartExType.Waterfall => new DataLabels { Position = DataLabelPosition.OutsideEnd },
            ChartExType.Funnel => new DataLabels { Position = DataLabelPosition.Center },
            _ => null,
        };

        return labels;
    }

    /// <summary>The kind of numbers a type's values are (cx:numDim type).</summary>
    private static string ValueKind(ChartExType type)
    {
        return type switch
        {
            ChartExType.Treemap or ChartExType.Sunburst => "size",
            ChartExType.RegionMap => "colorVal",
            _ => "val",
        };
    }

    private static void WriteData(XmlWriter xml, ChartExSpace chart, int index)
    {
        ChartExSeries series = chart.Series[index];
        xml.WriteStartElement("cx", "data", Namespace);
        xml.WriteAttributeString(null, "id", null, DrawingValues.Text(index));

        if (chart.Categories is ChartExCategories categories)
        {
            xml.WriteStartElement("cx", "strDim", Namespace);
            xml.WriteAttributeString(null, "type", null, "cat");

            if (categories.Formula is not null)
            {
                xml.WriteElementString("cx", "f", Namespace, categories.Formula);
            }

            // The leaves first, as the application writes the levels.
            for (int level = categories.Levels.Count - 1; level >= 0; level--)
            {
                IReadOnlyList<string?> names = categories.Levels[level];
                xml.WriteStartElement("cx", "lvl", Namespace);
                xml.WriteAttributeString(null, "ptCount", null, DrawingValues.Text(names.Count));

                for (int point = 0; point < names.Count; point++)
                {
                    if (names[point] is string name)
                    {
                        xml.WriteStartElement("cx", "pt", Namespace);
                        xml.WriteAttributeString(null, "idx", null, DrawingValues.Text(point));
                        xml.WriteString(name);
                        xml.WriteEndElement();
                    }
                }

                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }

        xml.WriteStartElement("cx", "numDim", Namespace);
        xml.WriteAttributeString(null, "type", null, ValueKind(chart.Type));

        if (series.Values.Formula is not null)
        {
            xml.WriteElementString("cx", "f", Namespace, series.Values.Formula);
        }

        IReadOnlyList<double?> numbers = series.Values.Numbers ?? [];
        xml.WriteStartElement("cx", "lvl", Namespace);
        xml.WriteAttributeString(null, "ptCount", null, DrawingValues.Text(numbers.Count));
        xml.WriteAttributeString(null, "formatCode", null, series.Values.FormatCode ?? "General");

        for (int point = 0; point < numbers.Count; point++)
        {
            if (numbers[point] is double number)
            {
                xml.WriteStartElement("cx", "pt", Namespace);
                xml.WriteAttributeString(null, "idx", null, DrawingValues.Text(point));
                xml.WriteString(DrawingValues.Text(number));
                xml.WriteEndElement();
            }
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    private static void WriteSeries(XmlWriter xml, ChartExSpace chart, int index)
    {
        ChartExSeries series = chart.Series[index];
        xml.WriteStartElement("cx", "series", Namespace);
        xml.WriteAttributeString(null, "layoutId", null, LayoutId(chart.Type));
        xml.WriteAttributeString(null, "uniqueId", null, UniqueId(index));

        if (series.Name is not null)
        {
            WriteText(xml, series.NameFormula, series.Name);
        }

        if (series.Color is string color)
        {
            xml.WriteStartElement("cx", "spPr", Namespace);
            DrawingWriter.WriteFill(xml, Fill.Solid(color));
            xml.WriteEndElement();
        }

        if ((series.DataLabels ?? chart.DataLabels ?? DefaultLabels(chart.Type)) is DataLabels labels)
        {
            WriteLabels(xml, labels);
        }

        Value(xml, "dataId", DrawingValues.Text(index));
        WriteLayout(xml, chart, series);

        if (chart.Type == ChartExType.Pareto)
        {
            Value(xml, "axisId", "1");
        }

        xml.WriteEndElement();
    }

    private static void WriteLayout(XmlWriter xml, ChartExSpace chart, ChartExSeries series)
    {
        switch (chart.Type)
        {
            case ChartExType.Treemap:
                xml.WriteStartElement("cx", "layoutPr", Namespace);
                Value(xml, "parentLabelLayout", chart.ParentLabels switch
                {
                    TreemapParentLabels.Banner => "banner",
                    TreemapParentLabels.None => "none",
                    _ => "overlapping",
                });
                xml.WriteEndElement();

                break;

            case ChartExType.Waterfall:
                xml.WriteStartElement("cx", "layoutPr", Namespace);
                xml.WriteStartElement("cx", "visibility", Namespace);
                xml.WriteAttributeString(null, "connectorLines", null, DrawingValues.Flag(chart.ConnectorLines));
                xml.WriteEndElement();

                if (series.Subtotals.Count > 0)
                {
                    xml.WriteStartElement("cx", "subtotals", Namespace);

                    foreach (int point in series.Subtotals)
                    {
                        Value(xml, "idx", DrawingValues.Text(point));
                    }

                    xml.WriteEndElement();
                }

                xml.WriteEndElement();

                break;

            case ChartExType.Histogram or ChartExType.Pareto:
                xml.WriteStartElement("cx", "layoutPr", Namespace);
                WriteBinning(xml, chart.Binning ?? new ChartExBinning());
                xml.WriteEndElement();

                break;

            case ChartExType.BoxWhisker:
                xml.WriteStartElement("cx", "layoutPr", Namespace);
                xml.WriteStartElement("cx", "visibility", Namespace);
                xml.WriteAttributeString(null, "meanLine", null, DrawingValues.Flag(chart.MeanLine));
                xml.WriteAttributeString(null, "meanMarker", null, DrawingValues.Flag(chart.MeanMarkers));
                xml.WriteAttributeString(null, "nonoutliers", null, DrawingValues.Flag(chart.InnerPoints));
                xml.WriteAttributeString(null, "outliers", null, DrawingValues.Flag(chart.Outliers));
                xml.WriteEndElement();
                xml.WriteStartElement("cx", "statistics", Namespace);
                xml.WriteAttributeString(null, "quartileMethod", null, Pick(chart.InclusiveQuartiles, "inclusive", "exclusive"));
                xml.WriteEndElement();
                xml.WriteEndElement();

                break;

            case ChartExType.RegionMap:
                xml.WriteStartElement("cx", "layoutPr", Namespace);
                xml.WriteStartElement("cx", "geography", Namespace);
                xml.WriteAttributeString(null, "cultureLanguage", null, "en-US");
                xml.WriteAttributeString(null, "cultureRegion", null, "US");
                xml.WriteAttributeString(null, "attribution", null, "Powered by Bing");
                xml.WriteEndElement();
                xml.WriteEndElement();

                break;

            default:
                break;
        }
    }

    private static void WriteBinning(XmlWriter xml, ChartExBinning binning)
    {
        xml.WriteStartElement("cx", "binning", Namespace);
        xml.WriteAttributeString(null, "intervalClosed", null, Pick(binning.ClosedLeft, "l", "r"));

        if (binning.Underflow is double underflow)
        {
            xml.WriteAttributeString(null, "underflow", null, DrawingValues.Text(underflow));
        }

        if (binning.Overflow is double overflow)
        {
            xml.WriteAttributeString(null, "overflow", null, DrawingValues.Text(overflow));
        }

        if (binning.BinWidth is double width)
        {
            Value(xml, "binSize", DrawingValues.Text(width));
        }
        else if (binning.BinCount is int count)
        {
            Value(xml, "binCount", DrawingValues.Text(count));
        }

        xml.WriteEndElement();
    }

    private static void WriteAxes(XmlWriter xml, ChartExSpace chart)
    {
        if (chart.HasAxes == false)
        {
            return;
        }

        string gap = chart.Type switch
        {
            ChartExType.Histogram or ChartExType.Pareto => "0",
            ChartExType.Funnel => "0.06",
            ChartExType.BoxWhisker => "1",
            _ => "0.5",
        };

        xml.WriteStartElement("cx", "axis", Namespace);
        xml.WriteAttributeString(null, "id", null, "0");

        if (chart.CategoryAxis == false)
        {
            xml.WriteAttributeString(null, "hidden", null, "1");
        }

        xml.WriteStartElement("cx", "catScaling", Namespace);
        xml.WriteAttributeString(null, "gapWidth", null, gap);
        xml.WriteEndElement();
        xml.WriteElementString("cx", "tickLabels", Namespace, string.Empty);
        xml.WriteEndElement();

        if (chart.Type == ChartExType.Funnel)
        {
            return;
        }

        xml.WriteStartElement("cx", "axis", Namespace);
        xml.WriteAttributeString(null, "id", null, "1");

        if (chart.ValueAxis == false)
        {
            xml.WriteAttributeString(null, "hidden", null, "1");
        }

        xml.WriteElementString("cx", "valScaling", Namespace, string.Empty);

        if ((chart.Gridlines ?? true) == true)
        {
            xml.WriteElementString("cx", "majorGridlines", Namespace, string.Empty);
        }

        xml.WriteElementString("cx", "tickLabels", Namespace, string.Empty);
        xml.WriteEndElement();

        if (chart.Type == ChartExType.Pareto)
        {
            xml.WriteStartElement("cx", "axis", Namespace);
            xml.WriteAttributeString(null, "id", null, "2");
            xml.WriteStartElement("cx", "valScaling", Namespace);
            xml.WriteAttributeString(null, "max", null, "1");
            xml.WriteAttributeString(null, "min", null, "0");
            xml.WriteEndElement();
            xml.WriteStartElement("cx", "units", Namespace);
            xml.WriteAttributeString(null, "unit", null, "percentage");
            xml.WriteEndElement();
            xml.WriteElementString("cx", "tickLabels", Namespace, string.Empty);
            xml.WriteEndElement();
        }
    }

    private static void WriteLabels(XmlWriter xml, DataLabels labels)
    {
        xml.WriteStartElement("cx", "dataLabels", Namespace);

        if (labels.Position != DataLabelPosition.Default)
        {
            xml.WriteAttributeString(null, "pos", null, DrawingNames.Position(labels.Position));
        }

        if (labels.NumberFormat is not null)
        {
            xml.WriteStartElement("cx", "numFmt", Namespace);
            xml.WriteAttributeString(null, "formatCode", null, labels.NumberFormat);
            xml.WriteAttributeString(null, "sourceLinked", null, "0");
            xml.WriteEndElement();
        }

        xml.WriteStartElement("cx", "visibility", Namespace);
        xml.WriteAttributeString(null, "seriesName", null, DrawingValues.Flag(labels.ShowSeriesName));
        xml.WriteAttributeString(null, "categoryName", null, DrawingValues.Flag(labels.ShowCategory));
        xml.WriteAttributeString(null, "value", null, DrawingValues.Flag(labels.ShowValue));
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    /// <summary>A cx:tx with a text and the reference it was read from.</summary>
    private static void WriteText(XmlWriter xml, string? formula, string text)
    {
        xml.WriteStartElement("cx", "tx", Namespace);
        xml.WriteStartElement("cx", "txData", Namespace);

        if (formula is not null)
        {
            xml.WriteElementString("cx", "f", Namespace, formula);
        }

        xml.WriteElementString("cx", "v", Namespace, text);
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    /// <summary>The series' id: the same for the same place, so that a chart written twice is the same XML.</summary>
    private static string UniqueId(int index)
    {
        return "{" + DrawingValues.Text(index).PadLeft(8, '0') + "-A51C-4B7D-9E0F-5C3A2B1D0E4F}";
    }

    private static void Value(XmlWriter xml, string name, string value)
    {
        xml.WriteStartElement("cx", name, Namespace);
        xml.WriteAttributeString(null, "val", null, value);
        xml.WriteEndElement();
    }

    private static string Pick(bool condition, string whenTrue, string whenFalse)
    {
        if (condition == true)
        {
            return whenTrue;
        }

        return whenFalse;
    }
}
