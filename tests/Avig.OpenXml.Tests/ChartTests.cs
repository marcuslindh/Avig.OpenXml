using System.Xml.Linq;

namespace Avig.OpenXml.Tests;

/// <summary>
/// Chart parts: every type and every style option passes Microsoft's schema validator, and a
/// chart read back writes the same XML again. Formulas are opaque here; the stored values a
/// presentation needs are written and read too.
/// </summary>
public class ChartTests
{
    public static TheoryData<ChartType> Types => [.. Enum.GetValues<ChartType>()];

    [Theory]
    [MemberData(nameof(Types))]
    public async Task Every_chart_type_validates_and_round_trips(ChartType type)
    {
        ChartSpace chart = new(type)
        {
            Title = "Sales per month",
            Series =
            [
                new ChartSpaceSeries(ChartData.Reference("'Sales 2026'!$B$2:$B$13")) { Name = "2026", Categories = ChartData.Reference("'Sales 2026'!$A$2:$A$13"), Color = "#2384F9" },
                new ChartSpaceSeries(ChartData.Reference("'Sales 2026'!$C$2:$C$13")) { Name = "2025", Categories = ChartData.Reference("'Sales 2026'!$A$2:$A$13") },
            ],
            Legend = LegendPosition.Bottom,
            CategoryAxisTitle = "Month",
            ValueAxisTitle = "Amount",
        };

        await RoundTripAsync(chart);
    }

    [Fact]
    public async Task Every_style_option_validates_and_round_trips()
    {
        TextFont font = new() { Name = "Segoe UI", Size = 9, Color = "#3A4A6B" };

        ChartSpace chart = new(ChartType.Column)
        {
            Title = "Budget and actual",
            TitleFont = font with { Size = 14, Bold = true },
            Series =
            [
                new ChartSpaceSeries(ChartData.Reference("Data!$B$2:$B$7"))
                {
                    Name = "Budget",
                    Categories = ChartData.Reference("Data!$A$2:$A$7"),
                    Fill = Fill.Gradient("#2384F9", "#605776"),
                    Outline = new Outline("#1F2940", 0.5),
                    Shadow = new Shadow(),
                    PointColors = [null, "#E1541C"],
                    DataLabels = new DataLabels { Position = DataLabelPosition.InsideEnd, NumberFormat = "0%", Font = font },
                },
                new ChartSpaceSeries(ChartData.Reference("Data!$C$2:$C$7"))
                {
                    Name = "Actual",
                    Type = ChartType.Line,
                    Color = "#00C5AE",
                    Smooth = true,
                    Marker = MarkerStyle.Diamond,
                    MarkerSize = 7,
                },
                new ChartSpaceSeries(ChartData.Reference("Data!$D$2:$D$7"))
                {
                    Name = "Share",
                    Type = ChartType.Line,
                    SecondaryAxis = true,
                    Outline = new Outline("#E1541C", 1.5) { Dash = LineDash.Dash },
                },
            ],
            Legend = LegendPosition.Top,
            LegendFont = font,
            GapWidth = 60,
            Overlap = -20,
            VaryColors = false,
            RoundedCorners = true,
            Background = Fill.Solid("#F4F6FA"),
            PlotBackground = Fill.None,
            Border = Outline.None,
            Font = font,
            CategoryAxis = new ChartAxis { Line = Outline.None, Font = font, NumberFormat = "mmm" },
            ValueAxis = new ChartAxis { Minimum = 0, Maximum = 120, MajorUnit = 20, GridlineOutline = new Outline("#DDDDDD", 0.5) { Dash = LineDash.RoundDot }, NumberFormat = "# ##0" },
            SecondaryValueAxis = new ChartAxis { Visible = true, Gridlines = false, NumberFormat = "0%" },
            DataLabels = new DataLabels { ShowValue = false, ShowCategory = true },
        };

        ChartSpace read = await RoundTripAsync(chart);

        Assert.Equal(ChartType.Column, read.Type);
        Assert.Equal(3, read.Series.Count);
        Assert.Equal(ChartType.Line, read.Series[1].Type);
        Assert.True(read.Series[2].SecondaryAxis);
        Assert.Equal("#E1541C", read.Series[0].PointColors[1]);
        Assert.Equal(120, read.ValueAxis!.Maximum);
        Assert.Equal("0%", read.SecondaryValueAxis!.NumberFormat);
        Assert.Equal("Data!$B$2:$B$7", read.Series[0].Values.Formula);
    }

    [Fact]
    public async Task Pie_doughnut_and_scatter_options_round_trip()
    {
        await RoundTripAsync(new ChartSpace(ChartType.Doughnut)
        {
            Series = [new ChartSpaceSeries(ChartData.Reference("Data!$B$1:$B$2")) { PointColors = ["#00C5AE", "#2E3850"], Outline = Outline.None }],
            HoleSize = 75,
            FirstSliceAngle = 90,
            Legend = LegendPosition.None,
        });

        await RoundTripAsync(new ChartSpace(ChartType.Pie)
        {
            Series = [new ChartSpaceSeries(ChartData.Reference("Data!$B$1:$B$4")) { DataLabels = new DataLabels { ShowPercent = true, Position = DataLabelPosition.BestFit } }],
        });

        await RoundTripAsync(new ChartSpace(ChartType.Scatter)
        {
            Series = [new ChartSpaceSeries(ChartData.Reference("Data!$B$1:$B$9")) { Categories = ChartData.Reference("Data!$A$1:$A$9"), Marker = MarkerStyle.Circle }],
            Stacked = false,
        });

        await RoundTripAsync(new ChartSpace(ChartType.Area) { Stacked = true, Series = [new ChartSpaceSeries(ChartData.Reference("Data!$B$1:$B$4")), new ChartSpaceSeries(ChartData.Reference("Data!$C$1:$C$4"))] });
    }

    [Fact]
    public async Task Stored_values_and_literals_round_trip()
    {
        ChartSpace chart = new(ChartType.Bar)
        {
            Series =
            [
                new ChartSpaceSeries(ChartData.FromNumbers([1.5, null, 3], "Sheet1!$B$2:$B$4") with { FormatCode = "0.0" })
                {
                    Name = "Revenue",
                    NameFormula = "Sheet1!$B$1",
                    Categories = ChartData.FromTexts(["North", "South", "East"], "Sheet1!$A$2:$A$4"),
                },
                new ChartSpaceSeries(ChartData.FromNumbers([4, 5, 6]))
                {
                    Name = "Literal",
                    Categories = ChartData.FromTexts(["a", "b", "c"]),
                },
            ],
        };

        ChartSpace read = await RoundTripAsync(chart);

        Assert.Equal([1.5, null, 3], read.Series[0].Values.Numbers!);
        Assert.Equal("0.0", read.Series[0].Values.FormatCode);
        Assert.Equal("Sheet1!$B$1", read.Series[0].NameFormula);
        Assert.Equal("Revenue", read.Series[0].Name);
        Assert.Equal(["North", "South", "East"], read.Series[0].Categories!.Texts!);
        Assert.Null(read.Series[1].Values.Formula);
        Assert.Equal([4.0, 5, 6], read.Series[1].Values.Numbers!);
    }

    [Fact]
    public void What_Office_would_repair_is_an_argument_exception()
    {
        ChartSpaceSeries series = new(ChartData.Reference("A1:A3"));

        Assert.Throws<ArgumentException>(() => new ChartSpace(ChartType.Pie) { Series = [series, series with { Type = ChartType.Line }] }.Validate());
        Assert.Throws<ArgumentException>(() => new ChartSpace(ChartType.Column) { Series = [series with { SecondaryAxis = true }] }.Validate());
        Assert.Throws<ArgumentException>(() => new ChartSpace(ChartType.Area) { Series = [series], DataLabels = new DataLabels { Position = DataLabelPosition.Above } }.Validate());
        Assert.Throws<ArgumentException>(() => new ChartSpace(ChartType.Column) { Series = [series], ValueAxis = new ChartAxis { Minimum = 5, Maximum = 1 } }.Validate());
        Assert.Throws<ArgumentException>(() => new ChartSpace(ChartType.Column) { Series = [series with { Fill = Fill.Picture(DrawingKit.PngPicture()) }] }.Validate());
        new ChartSpace(ChartType.Column) { Series = [series, series with { Type = ChartType.Line, SecondaryAxis = true }] }.Validate();
    }

    [Fact]
    public void A_hostile_chart_part_is_clamped_or_invalid_never_thrown()
    {
        string chartNamespace = OpcNamespaces.DrawingChart;
        XElement hostile = XElement.Parse($"""
            <c:chartSpace xmlns:c="{chartNamespace}">
              <c:chart><c:plotArea>
                <c:barChart><c:barDir val="col"/><c:gapWidth val="99999"/><c:overlap val="-999"/>
                  <c:ser><c:idx val="x"/><c:tx><c:v>{new string('n', 500)}</c:v></c:tx>
                    <c:marker><c:size val="9999"/></c:marker>
                    <c:dPt><c:idx val="99999999"/></c:dPt>
                    <c:val><c:numRef><c:f>A1</c:f><c:numCache><c:ptCount val="999999999"/><c:pt idx="-4"><c:v>NaN</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
                  </c:ser>
                  <c:axId val="1"/><c:axId val="2"/>
                </c:barChart>
                <c:valAx><c:axId val="2"/><c:scaling><c:max val="1"/><c:min val="5"/></c:scaling></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        Result<ChartSpace> read = ChartReader.Read(hostile);

        Assert.True(read.IsSuccess(out ChartSpace? chart), read.Error);
        Assert.Equal(500, chart.GapWidth);
        Assert.Equal(-100, chart.Overlap);
        Assert.Equal(255, chart.Series[0].Name!.Length);
        Assert.Equal(72, chart.Series[0].MarkerSize);
        Assert.Null(chart.ValueAxis!.Minimum);
        Assert.Equal(1_048_576, chart.Series[0].Values.Numbers!.Count);

        Assert.True(ChartReader.Read(XElement.Parse($"""<c:chartSpace xmlns:c="{chartNamespace}"><c:chart/></c:chartSpace>""")).IsFailure(out _, out _));
        Assert.True(ChartReader.Read(new XElement("notAChart")).IsFailure(out _, out _));
    }

    /// <summary>Validates the chart, writes it, validates the XML, reads it back and writes that again: the same XML.</summary>
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
