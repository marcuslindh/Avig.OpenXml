using System.Xml.Linq;

namespace Avig.OpenXml.Tests;

/// <summary>
/// Every chart type PowerPoint 365 and Excel offer (their XlChartType), as a <see cref="ChartSpace"/>:
/// each passes Microsoft's schema validator and writes the same XML after a round trip. The XML
/// follows what PowerPoint 365 wrote for each type.
/// </summary>
public class ChartKindsTests
{
    private static readonly ChartView3D Tilted = new();

    /// <summary>The XlChartType names and the chart that stands for each.</summary>
    public static readonly IReadOnlyDictionary<string, Func<ChartSpace>> Kinds = new Dictionary<string, Func<ChartSpace>>(StringComparer.Ordinal)
    {
        ["xlColumnClustered"] = () => Chart(ChartType.Column),
        ["xlColumnStacked"] = () => Chart(ChartType.Column) with { Grouping = ChartGrouping.Stacked },
        ["xlColumnStacked100"] = () => Chart(ChartType.Column) with { Grouping = ChartGrouping.PercentStacked },
        ["xl3DColumnClustered"] = () => Chart(ChartType.Column) with { ThreeD = Tilted },
        ["xl3DColumnStacked"] = () => Chart(ChartType.Column) with { ThreeD = Tilted, Grouping = ChartGrouping.Stacked },
        ["xl3DColumnStacked100"] = () => Chart(ChartType.Column) with { ThreeD = Tilted, Grouping = ChartGrouping.PercentStacked },
        ["xl3DColumn"] = () => Chart(ChartType.Column) with { ThreeD = Tilted, Grouping = ChartGrouping.Depth, SeriesAxis = new ChartAxis() },
        ["xlBarClustered"] = () => Chart(ChartType.Bar),
        ["xlBarStacked"] = () => Chart(ChartType.Bar) with { Grouping = ChartGrouping.Stacked, SeriesLines = true },
        ["xlBarStacked100"] = () => Chart(ChartType.Bar) with { Grouping = ChartGrouping.PercentStacked },
        ["xl3DBarClustered"] = () => Chart(ChartType.Bar) with { ThreeD = Tilted },
        ["xl3DBarStacked"] = () => Chart(ChartType.Bar) with { ThreeD = Tilted, Grouping = ChartGrouping.Stacked },
        ["xl3DBarStacked100"] = () => Chart(ChartType.Bar) with { ThreeD = Tilted, Grouping = ChartGrouping.PercentStacked },
        ["xlLine"] = () => Chart(ChartType.Line, marker: MarkerStyle.None),
        ["xlLineStacked"] = () => Chart(ChartType.Line, marker: MarkerStyle.None) with { Grouping = ChartGrouping.Stacked },
        ["xlLineStacked100"] = () => Chart(ChartType.Line, marker: MarkerStyle.None) with { Grouping = ChartGrouping.PercentStacked },
        ["xlLineMarkers"] = () => Chart(ChartType.Line, marker: MarkerStyle.Circle) with { DropLines = true, HighLowLines = true, UpDownBars = true },
        ["xlLineMarkersStacked"] = () => Chart(ChartType.Line, marker: MarkerStyle.Circle) with { Grouping = ChartGrouping.Stacked },
        ["xlLineMarkersStacked100"] = () => Chart(ChartType.Line, marker: MarkerStyle.Circle) with { Grouping = ChartGrouping.PercentStacked },
        ["xl3DLine"] = () => Chart(ChartType.Line) with { ThreeD = Tilted with { RightAngleAxes = false }, Grouping = ChartGrouping.Depth },
        ["xlPie"] = () => Chart(ChartType.Pie, 1),
        ["xlPieExploded"] = () => Chart(ChartType.Pie, 1, explosion: 25),
        ["xl3DPie"] = () => Chart(ChartType.Pie, 1) with { ThreeD = new ChartView3D { RotationX = 30, RotationY = 0, RightAngleAxes = false } },
        ["xl3DPieExploded"] = () => Chart(ChartType.Pie, 1, explosion: 25) with { ThreeD = new ChartView3D { RotationX = 30, RotationY = 0, RightAngleAxes = false } },
        ["xlPieOfPie"] = () => Chart(ChartType.PieOfPie, 1),
        ["xlBarOfPie"] = () => Chart(ChartType.BarOfPie, 1) with { Split = ChartSplit.Position, SplitValue = 2, SecondPlotSize = 60, GapWidth = 120 },
        ["xlDoughnut"] = () => Chart(ChartType.Doughnut) with { HoleSize = 75 },
        ["xlDoughnutExploded"] = () => Chart(ChartType.Doughnut, explosion: 25) with { HoleSize = 75 },
        ["xlXYScatter"] = () => Chart(ChartType.Scatter, marker: MarkerStyle.Circle),
        ["xlXYScatterSmooth"] = () => Chart(ChartType.Scatter, marker: MarkerStyle.Circle, smooth: true),
        ["xlXYScatterSmoothNoMarkers"] = () => Chart(ChartType.Scatter, marker: MarkerStyle.None, smooth: true),
        ["xlXYScatterLines"] = () => Chart(ChartType.Scatter, marker: MarkerStyle.Circle) with { ScatterLines = true },
        ["xlXYScatterLinesNoMarkers"] = () => Chart(ChartType.Scatter, marker: MarkerStyle.None) with { ScatterLines = true },
        ["xlArea"] = () => Chart(ChartType.Area) with { DropLines = true },
        ["xlAreaStacked"] = () => Chart(ChartType.Area) with { Grouping = ChartGrouping.Stacked },
        ["xlAreaStacked100"] = () => Chart(ChartType.Area) with { Grouping = ChartGrouping.PercentStacked },
        ["xl3DArea"] = () => Chart(ChartType.Area) with { ThreeD = Tilted with { RightAngleAxes = false }, Grouping = ChartGrouping.Depth },
        ["xl3DAreaStacked"] = () => Chart(ChartType.Area) with { ThreeD = Tilted with { RightAngleAxes = false }, Grouping = ChartGrouping.Stacked },
        ["xl3DAreaStacked100"] = () => Chart(ChartType.Area) with { ThreeD = Tilted with { RightAngleAxes = false }, Grouping = ChartGrouping.PercentStacked },
        ["xlRadar"] = () => Chart(ChartType.Radar, marker: MarkerStyle.None),
        ["xlRadarMarkers"] = () => Chart(ChartType.Radar, marker: MarkerStyle.Circle),
        ["xlRadarFilled"] = () => Chart(ChartType.Radar) with { RadarStyle = RadarStyle.Filled },
        ["xlSurface"] = () => Chart(ChartType.Surface, 3) with { ThreeD = Tilted with { RightAngleAxes = false } },
        ["xlSurfaceWireframe"] = () => Chart(ChartType.Surface, 3) with { ThreeD = Tilted with { RightAngleAxes = false }, Wireframe = true },
        ["xlSurfaceTopView"] = () => Chart(ChartType.Surface, 3),
        ["xlSurfaceTopViewWireframe"] = () => Chart(ChartType.Surface, 3) with { Wireframe = true },
        ["xlBubble"] = () => Chart(ChartType.Bubble),
        ["xlBubble3DEffect"] = () => Chart(ChartType.Bubble) with { Bubble3D = true, BubbleScale = 80, ShowNegativeBubbles = true },
        ["xlStockHLC"] = () => Chart(ChartType.Stock, 3),
        ["xlStockOHLC"] = () => Chart(ChartType.Stock, 4),
        ["xlStockVHLC"] = () => Volume(3),
        ["xlStockVOHLC"] = () => Volume(4),
        ["xlCylinderColClustered"] = () => Shaped(ChartType.Column, BarShape.Cylinder, ChartGrouping.Default),
        ["xlCylinderColStacked"] = () => Shaped(ChartType.Column, BarShape.Cylinder, ChartGrouping.Stacked),
        ["xlCylinderColStacked100"] = () => Shaped(ChartType.Column, BarShape.Cylinder, ChartGrouping.PercentStacked),
        ["xlCylinderBarClustered"] = () => Shaped(ChartType.Bar, BarShape.Cylinder, ChartGrouping.Default),
        ["xlCylinderBarStacked"] = () => Shaped(ChartType.Bar, BarShape.Cylinder, ChartGrouping.Stacked),
        ["xlCylinderBarStacked100"] = () => Shaped(ChartType.Bar, BarShape.Cylinder, ChartGrouping.PercentStacked),
        ["xlCylinderCol"] = () => Shaped(ChartType.Column, BarShape.Cylinder, ChartGrouping.Depth),
        ["xlConeColClustered"] = () => Shaped(ChartType.Column, BarShape.Cone, ChartGrouping.Default),
        ["xlConeColStacked"] = () => Shaped(ChartType.Column, BarShape.Cone, ChartGrouping.Stacked),
        ["xlConeColStacked100"] = () => Shaped(ChartType.Column, BarShape.Cone, ChartGrouping.PercentStacked),
        ["xlConeBarClustered"] = () => Shaped(ChartType.Bar, BarShape.Cone, ChartGrouping.Default),
        ["xlConeBarStacked"] = () => Shaped(ChartType.Bar, BarShape.Cone, ChartGrouping.Stacked),
        ["xlConeBarStacked100"] = () => Shaped(ChartType.Bar, BarShape.Cone, ChartGrouping.PercentStacked),
        ["xlConeCol"] = () => Shaped(ChartType.Column, BarShape.Cone, ChartGrouping.Depth),
        ["xlPyramidColClustered"] = () => Shaped(ChartType.Column, BarShape.Pyramid, ChartGrouping.Default),
        ["xlPyramidColStacked"] = () => Shaped(ChartType.Column, BarShape.Pyramid, ChartGrouping.Stacked),
        ["xlPyramidColStacked100"] = () => Shaped(ChartType.Column, BarShape.Pyramid, ChartGrouping.PercentStacked),
        ["xlPyramidBarClustered"] = () => Shaped(ChartType.Bar, BarShape.Pyramid, ChartGrouping.Default),
        ["xlPyramidBarStacked"] = () => Shaped(ChartType.Bar, BarShape.Pyramid, ChartGrouping.Stacked),
        ["xlPyramidBarStacked100"] = () => Shaped(ChartType.Bar, BarShape.Pyramid, ChartGrouping.PercentStacked),
        ["xlPyramidCol"] = () => Shaped(ChartType.Column, BarShape.Pyramid, ChartGrouping.Depth),
        ["xlComboColumnClusteredLine"] = () => Combo(ChartType.Column, ChartType.Line, false),
        ["xlComboColumnClusteredLineSecondaryAxis"] = () => Combo(ChartType.Column, ChartType.Line, true),
        ["xlComboAreaStackedColumnClustered"] = () => Combo(ChartType.Area, ChartType.Column, false) with { Grouping = ChartGrouping.Stacked },
    };

    public static TheoryData<string> Names => [.. Kinds.Keys];

    [Theory]
    [MemberData(nameof(Names))]
    public async Task Every_PowerPoint_chart_type_validates_and_round_trips(string name)
    {
        ChartSpace chart = Kinds[name]() with { Title = name, Style = 2, ExternalData = "rId3" };

        ChartSpace read = await RoundTripAsync(chart);

        Assert.Equal(chart.Type, read.Type);
        Assert.Equal(chart.Grouping, read.Grouping);
        Assert.Equal(chart.ThreeD is null, read.ThreeD is null);
        Assert.Equal(chart.BarShape, read.BarShape);
        Assert.Equal(chart.Series.Count, read.Series.Count);
        Assert.Equal("rId3", read.ExternalData);
        Assert.Equal(2, read.Style);
    }

    [Fact]
    public async Task Trendlines_error_bars_a_data_table_and_axes_round_trip()
    {
        ChartSpaceSeries series = Series(0, ChartType.Column) with
        {
            Trendlines =
            [
                new Trendline { Kind = TrendlineKind.Polynomial, Order = 3, Name = "Trend", DisplayEquation = true, DisplayRSquared = true, Forward = 1, Outline = new Outline("#C00000", 1.5) },
                new Trendline { Kind = TrendlineKind.MovingAverage, Period = 3 },
            ],
            ErrorBars = [new ErrorBars { Amount = ErrorBarAmount.Percentage, Value = 5, Side = ErrorBarSide.Plus, NoEndCap = true }],
        };

        ChartSpace chart = new(ChartType.Column)
        {
            Series = [series, Series(1, ChartType.Column)],
            DataTable = new ChartDataTable { VerticalBorders = false },
            ValueAxis = new ChartAxis { MinorGridlines = true, MinorUnit = 0.5, LogBase = 10, MinorGridlineOutline = new Outline("#EEEEEE", 0.25) },
            CategoryAxis = new ChartAxis { Reversed = true },
        };

        ChartSpace read = await RoundTripAsync(chart);

        Assert.Equal(2, read.Series[0].Trendlines.Count);
        Assert.Equal(TrendlineKind.Polynomial, read.Series[0].Trendlines[0].Kind);
        Assert.Equal(3, read.Series[0].Trendlines[0].Order);
        Assert.Equal("Trend", read.Series[0].Trendlines[0].Name);
        Assert.Equal(ErrorBarAmount.Percentage, read.Series[0].ErrorBars[0].Amount);
        Assert.False(read.DataTable!.VerticalBorders);
        Assert.True(read.ValueAxis!.MinorGridlines);
        Assert.Equal(10, read.ValueAxis.LogBase);
        Assert.True(read.CategoryAxis!.Reversed);

        ChartSpace scatter = Chart(ChartType.Scatter, marker: MarkerStyle.Circle);
        scatter = scatter with
        {
            Series = [scatter.Series[0] with { ErrorBars = [new ErrorBars { Direction = ErrorBarDirection.X, Amount = ErrorBarAmount.StandardError }, new ErrorBars { Amount = ErrorBarAmount.StandardDeviation, Value = 2 }] }],
        };

        ChartSpace readScatter = await RoundTripAsync(scatter);
        Assert.Equal(2, readScatter.Series[0].ErrorBars.Count);
    }

    [Fact]
    public void Options_a_type_does_not_take_are_refused()
    {
        Assert.Throws<ArgumentException>(() => Chart(ChartType.Scatter).With(chart => chart with { ThreeD = Tilted }).Validate());
        Assert.Throws<ArgumentException>(() => Chart(ChartType.Pie, 1).With(chart => chart with { Grouping = ChartGrouping.Stacked }).Validate());
        Assert.Throws<ArgumentException>(() => Chart(ChartType.Column).With(chart => chart with { Grouping = ChartGrouping.Depth }).Validate());
        Assert.Throws<ArgumentException>(() => Chart(ChartType.Column).With(chart => chart with { BarShape = BarShape.Cone }).Validate());
        Assert.Throws<ArgumentException>(() => Chart(ChartType.Stock, 2).Validate());
        Assert.Throws<ArgumentException>(() => Chart(ChartType.Pie, 1).With(chart => chart with { ThreeD = new ChartView3D { RotationX = -10 } }).Validate());
        Assert.Throws<ArgumentException>(() => Chart(ChartType.Surface, 3).With(chart => chart with { DataLabels = new DataLabels() }).Validate());
        Assert.Throws<ArgumentException>(() => Chart(ChartType.Column).With(chart => chart with { ThreeD = Tilted, Series = [chart.Series[0] with { Trendlines = [new Trendline()] }] }).Validate());
        Assert.Throws<ArgumentException>(() => Chart(ChartType.Column).With(chart => chart with { Series = [chart.Series[0] with { Explosion = 10 }] }).Validate());
        Assert.Throws<ArgumentException>(() => Chart(ChartType.Line).With(chart => chart with { Series = [chart.Series[0] with { ErrorBars = [new ErrorBars { Direction = ErrorBarDirection.X }] }] }).Validate());
        Assert.Throws<ArgumentException>(() => new ChartSpaceSeries(ChartData.Reference("A1")) { ErrorBars = [new ErrorBars(), new ErrorBars()] });
        Assert.Throws<ArgumentOutOfRangeException>(() => new ChartView3D { RotationX = 91 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new ChartSpace(ChartType.Bubble) { BubbleScale = 301 });

        // Stacked and Grouping are the same setting.
        ChartSpace stacked = new(ChartType.Column) { Stacked = true };

        Assert.Equal(ChartGrouping.Stacked, stacked.Grouping);
        Assert.True((stacked with { Grouping = ChartGrouping.PercentStacked }).Stacked);
        Assert.Equal(ChartGrouping.Default, (stacked with { Stacked = false }).Grouping);
    }

    [Fact]
    public void A_hostile_3D_chart_is_clamped()
    {
        string chartNamespace = OpcNamespaces.DrawingChart;
        XElement hostile = XElement.Parse($"""
            <c:chartSpace xmlns:c="{chartNamespace}" xmlns:r="{OpcNamespaces.Relationships}">
              <c:style val="9999"/>
              <c:chart>
                <c:view3D><c:rotX val="999"/><c:rotY val="-5"/><c:depthPercent val="1"/><c:perspective val="9999"/></c:view3D>
                <c:plotArea>
                  <c:bar3DChart><c:barDir val="bar"/><c:grouping val="standard"/><c:shape val="star"/>
                    <c:ser><c:idx val="0"/><c:explosion val="9999"/><c:trendline><c:trendlineType val="poly"/><c:order val="99"/></c:trendline>
                      <c:errBars><c:errDir val="x"/><c:val val="-3"/></c:errBars><c:errBars><c:errDir val="x"/></c:errBars><c:errBars/>
                      <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>
                  </c:bar3DChart>
                  <c:ofPieChart><c:secondPieSize val="0"/><c:splitPos val="-1"/></c:ofPieChart>
                </c:plotArea>
              </c:chart>
              <c:externalData r:id=""/>
            </c:chartSpace>
            """);

        Result<ChartSpace> read = ChartReader.Read(hostile);

        Assert.True(read.IsSuccess(out ChartSpace? chart), read.Error);
        Assert.Equal(ChartType.Bar, chart.Type);
        Assert.Equal(90, chart.ThreeD!.RotationX);
        Assert.Equal(0, chart.ThreeD.RotationY);
        Assert.Equal(20, chart.ThreeD.DepthPercent);
        Assert.Equal(ChartGrouping.Default, chart.Grouping);
        Assert.Equal(BarShape.Box, chart.BarShape);
        Assert.Equal(400, chart.Series[0].Explosion);
        Assert.Equal(6, chart.Series[0].Trendlines[0].Order);
        Assert.Equal(2, chart.Series[0].ErrorBars.Count);
        Assert.Null(chart.Style);
        Assert.Null(chart.ExternalData);
    }

    private static ChartSpace Chart(ChartType type, int count = 2, MarkerStyle? marker = null, bool smooth = false, int? explosion = null)
    {
        List<ChartSpaceSeries> series = [];

        for (int index = 0; index < count; index++)
        {
            ChartSpaceSeries item = Series(index, type) with { Marker = marker, Smooth = smooth, Explosion = explosion };

            series.Add(item);
        }

        return new ChartSpace(type) { Series = series };
    }

    private static ChartSpaceSeries Series(int index, ChartType type)
    {
        string column = ((char)('B' + index)).ToString();
        ChartData values = ChartData.FromNumbers([4.3 + index, 2.5, 3.5 - index, 4.5], $"Sheet1!${column}$2:${column}$5");
        ChartData categories = ChartData.FromTexts(["Category 1", "Category 2", "Category 3", "Category 4"], "Sheet1!$A$2:$A$5");

        if (type is ChartType.Scatter or ChartType.Bubble)
        {
            categories = ChartData.FromNumbers([0.7, 1.8, 2.6, 3.2], "Sheet1!$A$2:$A$5");
        }

        ChartSpaceSeries series = new(values) { Name = "Series " + (index + 1), NameFormula = $"Sheet1!${column}$1", Categories = categories };

        if (type == ChartType.Bubble)
        {
            series = series with { BubbleSizes = ChartData.FromNumbers([10, 4, 6, 8], $"Sheet1!$Z$2:$Z$5") };
        }

        return series;
    }

    private static ChartSpace Shaped(ChartType type, BarShape shape, ChartGrouping grouping)
    {
        return Chart(type) with { ThreeD = Tilted, BarShape = shape, Grouping = grouping };
    }

    /// <summary>A stock chart with the volume as columns on the primary axis and the prices on the secondary, as PowerPoint makes the volume-high-low-close charts.</summary>
    private static ChartSpace Volume(int prices)
    {
        ChartSpace stock = Chart(ChartType.Stock, prices + 1);
        List<ChartSpaceSeries> series = [stock.Series[0] with { Type = ChartType.Column }];
        series.AddRange(stock.Series.Skip(1).Select(item => item with { SecondaryAxis = true }));

        return stock with { Series = series };
    }

    private static ChartSpace Combo(ChartType first, ChartType second, bool secondary)
    {
        ChartSpace chart = Chart(first, 3);
        List<ChartSpaceSeries> series = [chart.Series[0], chart.Series[1], chart.Series[2] with { Type = second, SecondaryAxis = secondary }];

        return chart with { Series = series };
    }

    private static async Task<ChartSpace> RoundTripAsync(ChartSpace chart)
    {
        chart.Validate();
        string first = await DrawingKit.WriteChartAsync(chart);
        Assert.Empty(DrawingKit.ValidateChart(first));

        Result<ChartSpace> read = ChartReader.Read(XElement.Parse(first));
        Assert.True(read.IsSuccess(out ChartSpace? readChart), read.Error);
        readChart.Validate();
        string second = await DrawingKit.WriteChartAsync(readChart);
        Assert.Equal(first, second);

        return readChart;
    }
}

/// <summary>A fluent step for the tests.</summary>
internal static class ChartSpaceTestExtensions
{
    public static ChartSpace With(this ChartSpace chart, Func<ChartSpace, ChartSpace> change)
    {
        return change(chart);
    }
}
