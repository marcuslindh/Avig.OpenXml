using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Avig.OpenXml.Tests;

/// <summary>
/// Chart parts of the newer chart engine (cx:chartSpace) for every type, and the workbook a
/// chart keeps its data in: each passes Microsoft's schema validator, and a chart read back
/// writes the same XML again.
/// </summary>
public class ChartExTests
{
    public static TheoryData<ChartExType> Types => [.. Enum.GetValues<ChartExType>()];

    [Theory]
    [MemberData(nameof(Types))]
    public void Every_type_validates_and_round_trips(ChartExType type)
    {
        ChartExSpace chart = Sample(type) with { Title = type.ToString(), ExternalData = "rId1" };

        ChartExSpace read = RoundTrip(chart);

        Assert.Equal(type, read.Type);
        Assert.Equal(chart.Series.Count, read.Series.Count);
        Assert.Equal(chart.Series[0].Values.Numbers, read.Series[0].Values.Numbers);
        Assert.Equal(chart.Series[0].Values.Formula, read.Series[0].Values.Formula);
        Assert.Equal(chart.Categories?.Levels.Count, read.Categories?.Levels.Count);
        Assert.Equal("rId1", read.ExternalData);
        Assert.Equal(type.ToString(), read.Title);
    }

    [Fact]
    public void The_options_round_trip()
    {
        ChartExSpace waterfall = Sample(ChartExType.Waterfall) with { ConnectorLines = false, Legend = LegendPosition.Bottom, Gridlines = false, ValueAxis = false };

        ChartExSpace read = RoundTrip(waterfall);
        Assert.False(read.ConnectorLines);
        Assert.Equal([0, 3], read.Series[0].Subtotals);
        Assert.Equal(LegendPosition.Bottom, read.Legend);
        Assert.False(read.Gridlines);
        Assert.False(read.ValueAxis);

        ChartExSpace histogram = Sample(ChartExType.Histogram) with { Binning = new ChartExBinning { BinWidth = 2.5, Underflow = 1, Overflow = 9, ClosedLeft = true } };

        Assert.Equal(2.5, RoundTrip(histogram).Binning!.BinWidth);
        Assert.Equal(4, RoundTrip(Sample(ChartExType.Pareto) with { Binning = new ChartExBinning { BinCount = 4 } }).Binning!.BinCount);

        ChartExSpace box = Sample(ChartExType.BoxWhisker) with { InclusiveQuartiles = true, MeanLine = true, InnerPoints = true, Outliers = false, MeanMarkers = false };

        ChartExSpace readBox = RoundTrip(box);
        Assert.True(readBox.InclusiveQuartiles);
        Assert.True(readBox.MeanLine);
        Assert.False(readBox.MeanMarkers);

        ChartExSpace treemap = Sample(ChartExType.Treemap) with { ParentLabels = TreemapParentLabels.Banner, Series = [Sample(ChartExType.Treemap).Series[0] with { Color = "#2384F9", DataLabels = new DataLabels { ShowValue = true, ShowCategory = true, NumberFormat = "0.0" } }] };

        ChartExSpace readTreemap = RoundTrip(treemap);
        Assert.Equal(TreemapParentLabels.Banner, readTreemap.ParentLabels);
        Assert.Equal("#2384F9", readTreemap.Series[0].Color);
        Assert.Equal(["Branch 1", "Branch 1", "Branch 2", "Branch 2"], readTreemap.Categories!.Levels[0]);
    }

    [Fact]
    public void What_a_type_does_not_take_is_refused()
    {
        ChartExSeries series = new(ChartData.FromNumbers([1, 2]));

        Assert.Throws<ArgumentException>(() => new ChartExSpace(ChartExType.Funnel).Validate());
        Assert.Throws<ArgumentException>(() => new ChartExSpace(ChartExType.Funnel) { Series = [series, series] }.Validate());
        Assert.Throws<ArgumentException>(() => new ChartExSpace(ChartExType.Funnel) { Series = [series with { Subtotals = [1] }] }.Validate());
        Assert.Throws<ArgumentException>(() => new ChartExSpace(ChartExType.Funnel) { Series = [series], Categories = new ChartExCategories([["a"], ["b"]]) }.Validate());
        Assert.Throws<ArgumentException>(() => new ChartExSpace(ChartExType.RegionMap) { Series = [series] }.Validate());
        Assert.Throws<ArgumentOutOfRangeException>(() => new ChartExBinning { BinWidth = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => series with { Subtotals = [-1] });
        new ChartExSpace(ChartExType.BoxWhisker) { Series = [series, series] }.Validate();
    }

    [Fact]
    public void A_hostile_part_is_clamped_or_invalid_never_thrown()
    {
        XElement hostile = XElement.Parse($"""
            <cx:chartSpace xmlns:cx="{ChartExWriter.Namespace}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:lvl ptCount="999999999"><cx:pt idx="-1">x</cx:pt><cx:pt idx="2">c</cx:pt></cx:lvl></cx:strDim>
                <cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="0">NaN</cx:pt><cx:pt idx="1">4</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="clusteredColumn"><cx:dataId val="0"/><cx:layoutPr><cx:binning><cx:binCount val="99999"/><cx:binSize val="-2"/></cx:binning></cx:layoutPr></cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>
            """);

        Result<ChartExSpace> read = ChartExReader.Read(hostile);

        Assert.True(read.IsSuccess(out ChartExSpace? chart), read.Error);
        Assert.Equal(ChartExType.Histogram, chart.Type);
        Assert.Equal(1000, chart.Binning!.BinCount);
        Assert.Null(chart.Binning.BinWidth);
        Assert.Equal(1_048_576, chart.Categories!.Levels[0].Count);
        Assert.Equal([null, 4], chart.Series[0].Values.Numbers);

        Assert.True(ChartExReader.Read(XElement.Parse($"""<cx:chartSpace xmlns:cx="{ChartExWriter.Namespace}"><cx:chart/></cx:chartSpace>""")).IsFailure(out _, out _));
        Assert.True(ChartExReader.Read(XElement.Parse($"""<c:chartSpace xmlns:c="{OpcNamespaces.DrawingChart}"/>""")).IsFailure(out _, out _));
    }

    [Fact]
    public async Task A_chart_s_workbook_holds_its_data_and_opens_in_the_validator()
    {
        ChartSpace chart = new(ChartType.Column)
        {
            Series =
            [
                new ChartSpaceSeries(ChartData.FromNumbers([4.3, 2.5, null])) { Name = "Sales", Categories = ChartData.FromTexts(["North", "South", "East"]) },
                new ChartSpaceSeries(ChartData.FromNumbers([1, 2, 3])) { Categories = ChartData.FromTexts(["North", "South", "East"]) },
            ],
        };

        ChartWorkbook workbook = ChartWorkbook.From(chart);
        ChartSpace bound = workbook.Chart!;

        Assert.Equal("Sheet1!$B$2:$B$4", bound.Series[0].Values.Formula);
        Assert.Equal("Sheet1!$A$2:$A$4", bound.Series[0].Categories!.Formula);
        Assert.Equal("Sheet1!$C$1", bound.Series[1].NameFormula);
        Assert.Equal("Series 2", bound.Series[1].Name);
        Assert.Equal([4.3, 2.5, null], bound.Series[0].Values.Numbers);
        Assert.Equal("North", workbook.Rows[1][0].Text);
        Assert.Equal(2.5, workbook.Rows[2][1].Number);
        Assert.Empty(ValidateWorkbook(workbook.ToArray()));

        using MemoryStream streamed = new();
        await workbook.WriteAsync(streamed);
        Assert.Equal(workbook.ToArray().Length, streamed.ToArray().Length);

        // A scatter chart's x values differ per series: each gets a column of its own; a bubble's sizes one more.
        ChartSpace bubbles = new(ChartType.Bubble)
        {
            Series =
            [
                new ChartSpaceSeries(ChartData.FromNumbers([1, 2])) { Categories = ChartData.FromNumbers([5, 6]), BubbleSizes = ChartData.FromNumbers([3, 4]) },
                new ChartSpaceSeries(ChartData.FromNumbers([7, 8])) { Categories = ChartData.FromNumbers([9, 10]), BubbleSizes = ChartData.FromNumbers([1, 1]) },
            ],
        };

        ChartSpace boundBubbles = ChartWorkbook.From(bubbles).Chart!;
        Assert.Equal("Sheet1!$A$2:$A$3", boundBubbles.Series[0].Categories!.Formula);
        Assert.Equal("Sheet1!$C$2:$C$3", boundBubbles.Series[0].BubbleSizes!.Formula);
        Assert.Equal("Sheet1!$D$2:$D$3", boundBubbles.Series[1].Categories!.Formula);
        Assert.Equal("Sheet1!$E$2:$E$3", boundBubbles.Series[1].Values.Formula);

        ChartWorkbook treemap = ChartWorkbook.From(Sample(ChartExType.Treemap) with { Series = [new ChartExSeries(ChartData.FromNumbers([1, 2, 3, 4]))] });
        Assert.Equal("Sheet1!$A$2:$B$5", treemap.ExtendedChart!.Categories!.Formula);
        Assert.Equal("Sheet1!$C$2:$C$5", treemap.ExtendedChart.Series[0].Values.Formula);
        Assert.Equal("Sheet1!$C$1", treemap.ExtendedChart.Series[0].NameFormula);
        Assert.Empty(ValidateWorkbook(treemap.ToArray()));
    }

    internal static ChartExSpace Sample(ChartExType type)
    {
        ChartData values = ChartData.FromNumbers([22, 12, 18, 87], "Sheet1!$C$2:$C$5");
        ChartExCategories categories = ChartExCategories.Flat(["Stage 1", "Stage 2", "Stage 3", "Stage 4"]) with { Formula = "Sheet1!$A$2:$A$5" };

        List<ChartExSeries> series = [new ChartExSeries(values) { Name = "Values", NameFormula = "Sheet1!$C$1" }];

        switch (type)
        {
            case ChartExType.Treemap or ChartExType.Sunburst:
                categories = new ChartExCategories([["Branch 1", "Branch 1", "Branch 2", "Branch 2"], ["Leaf 1", "Leaf 2", "Leaf 3", "Leaf 4"]]) { Formula = "Sheet1!$A$2:$B$5" };

                break;

            case ChartExType.Waterfall:
                series[0] = series[0] with { Subtotals = [3, 0] };

                break;

            case ChartExType.BoxWhisker:
                series.Add(new ChartExSeries(ChartData.FromNumbers([3, 9, 2, 5], "Sheet1!$D$2:$D$5")) { Name = "More" });

                break;

            case ChartExType.RegionMap:
                categories = ChartExCategories.Flat(["Sweden", "Norway", "Finland", "Denmark"]) with { Formula = "Sheet1!$A$2:$A$5" };

                break;

            case ChartExType.Histogram or ChartExType.Pareto:
                return new ChartExSpace(type) { Series = series };

            default:
                break;
        }

        return new ChartExSpace(type) { Series = series, Categories = categories };
    }

    private static ChartExSpace RoundTrip(ChartExSpace chart)
    {
        chart.Validate();
        string first = Write(chart);
        List<string> problems = Validate(first);
        Assert.True(problems.Count == 0, string.Join(" | ", problems));

        Result<ChartExSpace> read = ChartExReader.Read(XElement.Parse(first));
        Assert.True(read.IsSuccess(out ChartExSpace? readChart), read.Error);
        readChart.Validate();
        Assert.Equal(first, Write(readChart));

        return readChart;
    }

    private static string Write(ChartExSpace chart)
    {
        StringBuilder text = new();
        XmlWriterSettings settings = new() { OmitXmlDeclaration = true };

        using (XmlWriter xml = XmlWriter.Create(text, settings))
        {
            ChartExWriter.WriteChartSpace(xml, chart);
        }

        return text.ToString();
    }

    private static List<string> Validate(string xml)
    {
        OpenXmlValidator validator = new(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);

        // PowerPoint 365 opens cx:axisId, cx:binSize and cx:binCount only with the value in a val
        // attribute; the SDK's schema has it as text, so those two findings are not held against the part.
        return [.. validator.Validate(new DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.ChartSpace(xml))
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .Where(problem => IsValueAttributeFinding(problem) == false)];
    }

    /// <summary>The validator's findings about the val attribute PowerPoint needs on cx:axisId, cx:binSize and cx:binCount.</summary>
    internal static bool IsValueAttributeFinding(string problem)
    {
        bool element = problem.Contains("cx:axisId", StringComparison.Ordinal) || problem.Contains("cx:binSize", StringComparison.Ordinal) || problem.Contains("cx:binCount", StringComparison.Ordinal);

        return element == true && (problem.Contains("The 'val' attribute is not declared", StringComparison.Ordinal) || problem.Contains("The text value cannot be empty", StringComparison.Ordinal));
    }

    private static List<string> ValidateWorkbook(byte[] bytes)
    {
        using MemoryStream stream = new(bytes);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
        OpenXmlValidator validator = new(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);

        return [.. validator.Validate(document).Select(error => $"{error.Path?.XPath}: {error.Description}")];
    }
}
