using System.Xml.Linq;

namespace Avig.OpenXml.Tests;

/// <summary>
/// The DrawingML layer on its own: every shape, fill, line, effect, text, picture and group is
/// written for a spreadsheet drawing and for a slide, passes Microsoft's schema validator in both,
/// and reads back into what writes the same XML again.
/// </summary>
public class DrawingTests
{
    public static TheoryData<string> HostNames => ["xdr", "p"];

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Every_named_shape_kind_validates_and_round_trips(string host)
    {
        ShapeTreeNamespace tree = Host(host);

        foreach (ShapeKind kind in Enum.GetValues<ShapeKind>())
        {
            Shape shape = new(kind)
            {
                Fill = Fill.Solid("#2384F9", 0.25),
                Outline = new Outline("#1F2940", 1.5) { Dash = LineDash.DashDot, Transparency = 0.5 },
                Shadow = new Shadow { Angle = 90, Blur = 6 },
                InnerShadow = new Shadow { Color = "#FFFFFF", Distance = 2 },
                SoftEdges = 3,
                Rotation = 30,
                FlipVertical = true,
                Name = "Shape of kind " + kind,
                Description = "Alternative text",
            };

            if (kind != ShapeKind.Line)
            {
                shape = shape with { Text = new ShapeText(kind.ToString()) };
            }

            ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, shape);

            Assert.Equal(ShapeTreeItemKind.Shape, read.Kind);
            Assert.Equal(kind, read.Shape!.Kind);
            Assert.Equal("Shape of kind " + kind, read.Name);
            Assert.Equal(30, read.Shape.Rotation, 6);
            Assert.True(read.Shape.FlipVertical);
            Assert.Equal(DrawingKit.Frame, read.Bounds);
        }
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Any_preset_geometry_with_adjust_values_round_trips(string host)
    {
        ShapeTreeNamespace tree = Host(host);

        foreach (string preset in new[] { "star5", "cloud", "flowChartDecision", "wedgeRoundRectCallout", "gear9", "chartPlus" })
        {
            Shape shape = Shape.FromPreset(preset) with { Adjustments = [new ShapeAdjustment("adj", 12_500)] };

            ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, shape);

            Assert.Null(read.Shape!.Kind);
            Assert.Equal(preset, read.Shape.Preset);
            Assert.Equal(12_500, Assert.Single(read.Shape.Adjustments).Value);
        }

        Assert.Equal(ShapeKind.Ellipse, Shape.FromPreset("ellipse").Kind);
        Assert.Throws<ArgumentException>(() => Shape.FromPreset("notAShape"));
    }

    [Fact]
    public async Task A_rounded_rectangle_keeps_its_corner_radius()
    {
        ShapeTreeItem read = await DrawingKit.RoundTripAsync(DrawingKit.Slide, new Shape(ShapeKind.RoundedRectangle) { CornerRadius = 0.25 });

        Assert.Equal(0.25, read.Shape!.CornerRadius);
        Assert.Empty(read.Shape.Adjustments);
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Every_fill_validates_and_round_trips(string host)
    {
        ShapeTreeNamespace tree = Host(host);
        List<Fill> fills =
        [
            Fill.None,
            Fill.Solid("#00C5AE"),
            Fill.Solid("#00C5AE", 0.4),
            Fill.Gradient("#2384F9", "#605776", 45),
            Fill.Gradient(120, new GradientStop(0, "#FF0000"), new GradientStop(0.5, "#00FF00", 0.3), new GradientStop(1, "#0000FF")),
            Fill.Picture(DrawingKit.PngPicture()),
            Fill.Picture(DrawingKit.SvgPicture()),
        ];

        foreach (FillPattern pattern in Enum.GetValues<FillPattern>())
        {
            fills.Add(Fill.Pattern(pattern, "#1F2940", "#FFFFFF"));
        }

        foreach (Fill fill in fills)
        {
            ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, new Shape(ShapeKind.Rectangle) { Fill = fill });

            Assert.Equal(fill.Kind, read.Shape!.Fill!.Kind);
            Assert.Equal(fill.PatternType, read.Shape.Fill.PatternType);
            Assert.Equal(fill.Color, read.Shape.Fill.Color);
        }
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Text_with_paragraphs_runs_fonts_and_margins_round_trips(string host)
    {
        TextFont heading = new() { Name = "Segoe UI", Size = 14, Bold = true, Color = "#FFFFFF" };

        TextFont body = new() { Italic = true, Size = 9 };

        ShapeText text = new([new TextRun("Projects\n", heading), new TextRun("12 open, "), new TextRun("3 late\n\nlast line", body)])
        {
            Horizontal = TextAlignment.Right,
            Vertical = TextAnchor.Bottom,
            Wrap = false,
            MarginLeft = 4,
            MarginTop = 2,
            MarginRight = 0,
            MarginBottom = 7.2,
        };

        ShapeTreeItem read = await DrawingKit.RoundTripAsync(Host(host), new Shape(ShapeKind.TextBox) { Text = text });
        ShapeText readText = read.Shape!.Text!;

        Assert.Equal(ShapeKind.TextBox, read.Shape.Kind);
        Assert.Equal("Projects\n12 open, 3 late\n\nlast line", readText.PlainText);
        Assert.Equal(TextAlignment.Right, readText.Horizontal);
        Assert.Equal(TextAnchor.Bottom, readText.Vertical);
        Assert.False(readText.Wrap);
        Assert.Equal(4, readText.MarginLeft);
        Assert.Equal(heading, readText.Runs[0].Font);
        Assert.Equal(body, readText.Runs[^1].Font);
    }

    [Fact]
    public async Task A_shape_without_text_reads_back_without_text()
    {
        ShapeTreeItem read = await DrawingKit.RoundTripAsync(DrawingKit.Spreadsheet, new Shape(ShapeKind.Ellipse));

        Assert.Null(read.Shape!.Text);
        Assert.Null(read.Shape.Fill);
        Assert.Null(read.Shape.Outline);
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Links_on_shapes_share_a_relationship_and_round_trip(string host)
    {
        Hyperlink web = Hyperlink.ToUrl("https://example.com/report") with { Tooltip = "Open the report" };

        Hyperlink inside = Hyperlink.ToLocation("'Sales 2026'!A1");
        Hyperlink jump = Hyperlink.ToLocation("slide3.xml") with { Action = "ppaction://hlinksldjump", HighlightClick = true, StopsSound = true };

        (string xml, DrawingContext context) = await DrawingKit.WriteAsync(
            Host(host),
            (writer, written) => DrawingWriter.WriteShapeAsync(writer, new Shape(ShapeKind.Rectangle) { Link = web }, written, DrawingKit.Frame),
            (writer, written) => DrawingWriter.WriteShapeAsync(writer, new Shape(ShapeKind.Ellipse) { Link = web }, written, DrawingKit.Frame),
            (writer, written) => DrawingWriter.WriteShapeAsync(writer, new Shape(ShapeKind.Diamond) { Link = inside }, written, DrawingKit.Frame),
            (writer, written) => DrawingWriter.WriteShapeAsync(writer, new Shape(ShapeKind.Chevron) { Link = jump }, written, DrawingKit.Frame));

        Assert.Empty(DrawingKit.ValidatePart(Host(host), xml));
        Assert.Equal(3, context.Targets.Count);

        List<Hyperlink?> links = [];

        foreach (XElement element in DrawingKit.Items(xml))
        {
            Result<ShapeTreeItem> read = await DrawingReader.ReadAsync(element, DrawingKit.ReadContext(context));
            Assert.True(read.IsSuccess(out ShapeTreeItem? item), read.Error);
            links.Add(item.Shape!.Link);
        }

        Assert.Equal([web, web, inside, jump], links);
        Assert.True(links[0]!.IsExternal);
        Assert.False(links[2]!.IsExternal);
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Pictures_round_trip_with_svg_and_png_fallback(string host)
    {
        Picture svg = DrawingKit.SvgPicture() with
        {
            Description = "A blue box",
            Transparency = 0.25,
            Shadow = new Shadow(),
            Link = Hyperlink.ToUrl("https://example.com"),
        };

        ShapeTreeItem read = await DrawingKit.RoundTripAsync(Host(host), svg);
        Picture picture = read.Picture!;

        Assert.Equal(PictureFormat.Svg, picture.Format);
        Assert.Equal(48, picture.Width);
        Assert.Equal(24, picture.Height);
        Assert.Equal(PictureFormat.Png, picture.Fallback!.Format);
        Assert.Equal("A blue box", picture.Description);
        Assert.Equal(0.25, picture.Transparency, 6);
        Assert.NotNull(picture.Shadow);
        Assert.Equal(svg.Data.ToArray(), picture.Data.ToArray());

        ShapeTreeItem png = await DrawingKit.RoundTripAsync(Host(host), DrawingKit.PngPicture());

        Assert.Equal(PictureFormat.Png, png.Picture!.Format);
        Assert.Null(png.Picture.Fallback);
    }

    [Fact]
    public async Task The_same_image_twice_shares_one_relationship()
    {
        Picture picture = DrawingKit.PngPicture();

        (_, DrawingContext context) = await DrawingKit.WriteAsync(DrawingKit.Slide, async (writer, written) =>
        {
            await DrawingWriter.WritePictureAsync(writer, picture, written, DrawingKit.Frame);
            await DrawingWriter.WritePictureAsync(writer, picture with { Description = "Again" }, written, DrawingKit.Frame);
            await DrawingWriter.WriteShapeAsync(writer, new Shape(ShapeKind.Rectangle) { Fill = Fill.Picture(picture) }, written, DrawingKit.Frame);
        });

        Assert.Single(context.Targets);
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Groups_nest_and_round_trip(string host)
    {
        ShapeGroup inner = new([
            GroupItem.Of(new Shape(ShapeKind.Ellipse) { Fill = Fill.Solid("#E1541C") }, DrawingBounds.FromPoints(0, 0, 20, 20)),
            GroupItem.Of(DrawingKit.PngPicture(), DrawingBounds.FromPoints(25, 0, 20, 20)),
        ]);
        ShapeGroup outer = new([
            GroupItem.Of(new Shape(ShapeKind.RoundedRectangle) { Fill = Fill.Solid("#2E3850") }, DrawingBounds.FromPoints(0, 0, 145, 40)),
            GroupItem.Of(new Shape(ShapeKind.TextBox) { Text = new ShapeText("Data", new TextFont { Color = "#FFFFFF" }) { Horizontal = TextAlignment.Left } }, DrawingBounds.FromPoints(42, 0, 100, 40)),
            GroupItem.Of(inner, DrawingBounds.FromPoints(5, 10, 45, 20)),
        ])
        {
            Link = Hyperlink.ToLocation("'Data'!A1"),
            Name = "Navigation",
            Description = "Go to the data",
        };

        ShapeTreeItem read = await DrawingKit.RoundTripAsync(Host(host), outer);

        Assert.Equal(ShapeTreeItemKind.Group, read.Kind);
        Assert.Equal(3, read.Group!.Items.Count);
        Assert.Equal(2, read.Group.Items[2].Group!.Items.Count);
        Assert.Equal(DrawingBounds.FromPoints(42, 0, 100, 40), read.Group.Items[1].Bounds);
        Assert.Equal("Navigation", read.Name);
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task A_chart_frame_points_to_its_chart_part(string host)
    {
        object chart = new();
        (string xml, DrawingContext context) = await DrawingKit.WriteItemAsync(Host(host), chart);

        Assert.Empty(DrawingKit.ValidatePart(Host(host), xml));
        DrawingTarget target = Assert.Single(context.Targets);
        Assert.Same(chart, target.Chart);

        Result<ShapeTreeItem> read = await DrawingReader.ReadAsync(Assert.Single(DrawingKit.Items(xml)), DrawingKit.ReadContext(context));
        Assert.True(read.IsSuccess(out ShapeTreeItem? item), read.Error);
        Assert.Equal(ShapeTreeItemKind.GraphicFrame, item.Kind);
        Assert.Equal(OpcNamespaces.DrawingChart, item.GraphicDataUri);
        Assert.Equal(target.Id, item.ChartRelationshipId);
        Assert.Equal("A chart", item.Description);
        Assert.Equal(DrawingKit.Frame, item.Bounds);
    }

    [Fact]
    public async Task The_hosts_differ_only_in_the_shape_tree_namespace_and_its_two_details()
    {
        Shape shape = new(ShapeKind.Rectangle) { Text = new ShapeText("Same") };

        (string sheet, _) = await DrawingKit.WriteItemAsync(DrawingKit.Spreadsheet, shape);
        (string slide, _) = await DrawingKit.WriteItemAsync(DrawingKit.Slide, shape);

        XElement fromSheet = Assert.Single(DrawingKit.Items(sheet));
        XElement fromSlide = Assert.Single(DrawingKit.Items(slide));

        Assert.Equal("", fromSheet.Attribute("macro")?.Value);
        Assert.Null(fromSlide.Attribute("macro"));
        Assert.NotNull(fromSlide.Descendants().FirstOrDefault(element => element.Name.LocalName == "nvPr"));
        Assert.Equal(DrawingMl(fromSheet), DrawingMl(fromSlide));
    }

    /// <summary>The a: elements of a shape, as text: what is the same in every host.</summary>
    private static string DrawingMl(XElement shape)
    {
        return string.Concat(shape.Descendants().Where(element => element.Name.NamespaceName == OpcNamespaces.DrawingMain && element.Parent!.Name.NamespaceName != OpcNamespaces.DrawingMain)
            .Select(element => element.ToString(SaveOptions.DisableFormatting)));
    }

    [Fact]
    public async Task A_hostile_shape_is_clamped_not_thrown()
    {
        string main = OpcNamespaces.DrawingMain;
        XElement shape = XElement.Parse($"""
            <p:sp xmlns:p="{OpcNamespaces.PresentationMain}" xmlns:a="{main}">
              <p:nvSpPr><p:cNvPr id="x" name="{new string('n', 5000)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr>
                <a:xfrm rot="999999999999" flipH="maybe"><a:off x="-5" y="99999999999999999"/><a:ext cx="-1" cy="abc"/></a:xfrm>
                <a:prstGeom prst="noSuchShape"><a:avLst><a:gd name="adj" fmla="val 99999999999999"/><a:gd name="bad name" fmla="val 1"/></a:avLst></a:prstGeom>
                <a:gradFill><a:gsLst><a:gs pos="-7"><a:srgbClr val="ZZZZZZ"/></a:gs></a:gsLst></a:gradFill>
                <a:ln w="-100000000"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="-5"/></a:srgbClr></a:solidFill></a:ln>
                <a:effectLst><a:outerShdw blurRad="999999999999" dist="-3" dir="x"><a:srgbClr val="000000"/></a:outerShdw><a:softEdge rad="99999999999"/></a:effectLst>
              </p:spPr>
              <p:txBody><a:bodyPr lIns="-4" anchor="sideways"/><a:p><a:r><a:rPr sz="999999" b="yes"><a:latin typeface="{new string('f', 40)}"/></a:rPr><a:t>{new string('t', 40_000)}</a:t></a:r></a:p></p:txBody>
            </p:sp>
            """);

        Result<ShapeTreeItem> read = await DrawingReader.ReadAsync(shape, new DrawingReadContext([], (_, _) => Task.FromResult(Result<byte[]>.NotFound("none"))));

        Assert.True(read.IsSuccess(out ShapeTreeItem? item), read.Error);
        Shape result = item.Shape!;
        Assert.Equal("rect", result.Preset);
        Assert.Null(result.Fill);
        Assert.Equal(0, result.Outline!.Width);
        Assert.Equal(100, result.SoftEdges);
        Assert.Equal(400, result.Text!.Runs[0].Font!.Size);
        Assert.Equal(32_767, result.Text.PlainText.Length);
        Assert.Equal(1000, result.Name!.Length);
        Assert.Equal(-5, item.Bounds.Left.Emu);
        Assert.Equal(0, item.Bounds.Width.Emu);
    }

    [Fact]
    public async Task A_picture_whose_image_is_missing_is_invalid()
    {
        Picture picture = DrawingKit.PngPicture();
        (string xml, _) = await DrawingKit.WriteItemAsync(DrawingKit.Slide, picture);

        Result<ShapeTreeItem> read = await DrawingReader.ReadAsync(Assert.Single(DrawingKit.Items(xml)), new DrawingReadContext([], (_, _) => Task.FromResult(Result<byte[]>.NotFound("none"))));

        Assert.True(read.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
    }

    [Fact]
    public void Bounds_are_lengths()
    {
        DrawingBounds bounds = new(Length.Cm(2), Length.Pt(10), Length.Inch(1), Length.Mm(5));

        Assert.Equal(720_000, bounds.Left.Emu);
        Assert.Equal(Length.Pt(10), bounds.Top);
        Assert.Equal(Length.Cm(2) + Length.Inch(1), bounds.Right);
        Assert.Equal(DrawingBounds.FromEmu(720_000, 127_000, 914_400, 180_000), bounds);
        Assert.Equal(DrawingBounds.FromPoints(10, 10, 72, 36), new DrawingBounds(Length.Pt(10), Length.Pt(10), Length.Inch(1), Length.Inch(0.5)));
        Assert.Throws<ArgumentOutOfRangeException>(() => DrawingBounds.FromPoints(-1, 0, 1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new DrawingBounds(Length.Pt(-1), Length.Zero, Length.Zero, Length.Zero));
    }

    [Fact]
    public void Values_out_of_range_are_argument_exceptions()
    {
        Assert.Throws<ArgumentException>(() => Fill.Solid("blue"));
        Assert.Throws<ArgumentOutOfRangeException>(() => Fill.Solid("#000000", 2));
        Assert.Throws<ArgumentException>(() => Fill.Gradient(0, new GradientStop(0, "#000000")));
        Assert.Throws<ArgumentOutOfRangeException>(() => new Outline("#000000", 2000));
        Assert.Throws<ArgumentOutOfRangeException>(() => new Shadow { Blur = 101 });
        Assert.Throws<ArgumentException>(() => new TextFont { Name = new string('x', 32) });
        Assert.Throws<ArgumentException>(() => new TextRun("bad \u0001 text"));
        Assert.Throws<ArgumentException>(() => Hyperlink.ToUrl("ftp://example.com"));
        Assert.Throws<ArgumentException>(() => Hyperlink.ToLocation(string.Empty));
        Assert.Throws<ArgumentException>(() => new Shape(ShapeKind.Rectangle) { Adjustments = [new ShapeAdjustment("no space", 1)] });
        Assert.Throws<ArgumentOutOfRangeException>(() => new Shape((ShapeKind)99));
    }

    private static ShapeTreeNamespace Host(string prefix)
    {
        if (prefix == "xdr")
        {
            return DrawingKit.Spreadsheet;
        }

        return DrawingKit.Slide;
    }
}
