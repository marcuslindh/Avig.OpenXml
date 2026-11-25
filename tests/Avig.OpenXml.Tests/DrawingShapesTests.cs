namespace Avig.OpenXml.Tests;

/// <summary>
/// Everything a shape can be beside the basics: every preset, geometries of its own, connectors
/// glued to shapes, every gradient, tile, line, dash, arrowhead, effect, bevel, material, camera
/// and light, theme styles, locks and the non-visual flags. Each is written for a spreadsheet
/// drawing and a slide, passes Microsoft's schema validator, and reads back into what writes the
/// same XML again.
/// </summary>
public class DrawingShapesTests
{
    public static TheoryData<string> HostNames => ["xdr", "p"];

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Every_preset_geometry_validates_and_round_trips(string host)
    {
        ShapeTreeNamespace tree = Host(host);
        int count = 0;

        foreach (string preset in Shape.PresetNames)
        {
            ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, Shape.FromPreset(preset));

            Assert.Equal(preset, read.Shape!.Preset);
            count++;
        }

        Assert.Equal(187, count);
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Connectors_are_cxnSp_and_keep_their_glued_ends(string host)
    {
        ShapeTreeNamespace tree = Host(host);

        foreach (string preset in new[] { "line", "straightConnector1", "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5", "curvedConnector2", "curvedConnector3", "curvedConnector4", "curvedConnector5" })
        {
            Shape connector = Shape.FromPreset(preset) with
            {
                Outline = new Outline("#203864", 2) { Tail = new LineEnd(LineEndType.Triangle) },
                StartConnection = new ConnectionSite(4, 3),
                EndConnection = new ConnectionSite(5, 1),
                Locks = new ShapeLocks { NoChangeArrowheads = true },
            };

            (string xml, _) = await DrawingKit.WriteItemAsync(tree, connector);
            Assert.Contains("cxnSp", xml, StringComparison.Ordinal);
            ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, connector);

            Assert.True(read.Shape!.IsConnector);
            Assert.Equal(new ConnectionSite(4, 3), read.Shape.StartConnection);
            Assert.Equal(new ConnectionSite(5, 1), read.Shape.EndConnection);
            Assert.True(read.Shape.Locks!.NoChangeArrowheads);
        }
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task A_geometry_of_its_own_with_every_path_command_round_trips(string host)
    {
        CustomGeometry geometry = new(
        [
            new GeometryPath(1000, 800,
            [
                PathCommand.MoveTo(0, 400),
                PathCommand.LineTo(300, 0),
                PathCommand.CubicBezierTo(new PathPoint(500, -100), new PathPoint(700, 100), new PathPoint(1000, 400)),
                PathCommand.QuadraticBezierTo(new PathPoint(800, 800), new PathPoint(500, 800)),
                PathCommand.ArcTo(250, 200, 90, 180),
                PathCommand.Close(),
            ]),
            new GeometryPath(1000, 800, [PathCommand.MoveTo(100, 100), PathCommand.LineTo(900, 700)]) { Fill = PathFill.None, ExtrusionOk = false },
            new GeometryPath(1000, 800, [PathCommand.MoveTo(0, 0), PathCommand.LineTo(1000, 0), PathCommand.LineTo(1000, 800), PathCommand.Close()]) { Fill = PathFill.Darken, Stroke = false },
        ]);

        ShapeTreeItem read = await DrawingKit.RoundTripAsync(Host(host), Shape.FromGeometry(geometry) with { Fill = Fill.Solid("#2384F9"), Text = new ShapeText("Freeform") });

        Assert.NotNull(read.Shape!.Geometry);
        Assert.Equal(3, read.Shape.Geometry.Paths.Count);
        Assert.Equal(6, read.Shape.Geometry.Paths[0].Commands.Count);
        Assert.Equal(PathCommandKind.ArcTo, read.Shape.Geometry.Paths[0].Commands[4].Kind);
        Assert.Equal(180, read.Shape.Geometry.Paths[0].Commands[4].SweepAngle);
        Assert.Equal(PathFill.Darken, read.Shape.Geometry.Paths[2].Fill);
        Assert.False(read.Shape.Geometry.Paths[2].Stroke);
        Assert.StartsWith("Freeform", read.Name, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Every_gradient_tile_and_group_fill_round_trips(string host)
    {
        ShapeTreeNamespace tree = Host(host);
        GradientStop[] stops = [new(0, "#FFFFFF"), new(0.4, "#00C5AE", 0.2), new(1, "#203864")];
        List<Fill> fills =
        [
            Fill.PathGradient(GradientShape.Circle, FillRectangle.Center, stops),
            Fill.PathGradient(GradientShape.Rectangle, new FillRectangle(1, 1, 0, 0), stops),
            Fill.PathGradient(GradientShape.Shape, new FillRectangle(0.5, 0.5, 0.5, 0.5), stops),
            Fill.PathGradient(GradientShape.Circle, null, stops).WithoutRotation(),
            Fill.Gradient(30, stops).WithScaledAngle().WithTileRectangle(new FillRectangle(-0.5, 0, 0, -0.5)),
            Fill.TiledPicture(DrawingKit.PngPicture()),
            Fill.TiledPicture(DrawingKit.PngPicture(), new PictureTiling { OffsetX = Length.Pt(3), OffsetY = Length.Pt(-2), ScaleX = 0.5, ScaleY = 2, Flip = TileFlip.Both, Alignment = RectangleAlignment.Center }),
            Fill.Picture(DrawingKit.PngPicture()).WithoutRotation(),
            Fill.Group,
        ];

        foreach (Fill fill in fills)
        {
            ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, new Shape(ShapeKind.Ellipse) { Fill = fill });

            Assert.Equal(fill.Kind, read.Shape!.Fill!.Kind);
            Assert.Equal(fill.GradientShape, read.Shape.Fill.GradientShape);
            Assert.Equal(fill.Focus, read.Shape.Fill.Focus);
            Assert.Equal(fill.Tiling, read.Shape.Fill.Tiling);
            Assert.Equal(fill.RotateWithShape, read.Shape.Fill.RotateWithShape);
        }
    }

    [Fact]
    public async Task A_background_fill_is_useBgFill_on_a_slide_and_no_fill_on_a_sheet()
    {
        Shape shape = new(ShapeKind.Rectangle) { Fill = Fill.Background };

        ShapeTreeItem read = await DrawingKit.RoundTripAsync(DrawingKit.Slide, shape);
        (string sheet, _) = await DrawingKit.WriteItemAsync(DrawingKit.Spreadsheet, shape);

        Assert.Equal(FillKind.Background, read.Shape!.Fill!.Kind);
        Assert.DoesNotContain("useBgFill", sheet, StringComparison.Ordinal);
        Assert.Empty(DrawingKit.ValidatePart(DrawingKit.Spreadsheet, sheet));
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Every_dash_compound_cap_join_and_arrowhead_round_trips(string host)
    {
        ShapeTreeNamespace tree = Host(host);
        List<Outline> outlines = [];

        foreach (LineDash dash in Enum.GetValues<LineDash>())
        {
            outlines.Add(new Outline("#1F2940", 2) { Dash = dash });
        }

        foreach (LineCompound compound in Enum.GetValues<LineCompound>())
        {
            outlines.Add(new Outline("#1F2940", 6) { Compound = compound });
        }

        foreach (LineCap cap in Enum.GetValues<LineCap>())
        {
            outlines.Add(new Outline("#1F2940", 6) { Cap = cap, Alignment = LineAlignment.Inset });
        }

        foreach (LineJoin join in Enum.GetValues<LineJoin>())
        {
            outlines.Add(new Outline("#1F2940", 6) { Join = join });
        }

        outlines.Add(new Outline("#1F2940", 6) { Join = LineJoin.Miter, MiterLimit = 8 });

        foreach (LineEndType type in Enum.GetValues<LineEndType>())
        {
            foreach (LineEndSize size in Enum.GetValues<LineEndSize>())
            {
                outlines.Add(new Outline("#1F2940", 2) { Head = new LineEnd(type, size, LineEndSize.Large), Tail = new LineEnd(type, LineEndSize.Small, size) });
            }
        }

        outlines.Add(new Outline("#1F2940", 2) { CustomDash = [new DashStop(4, 1), new DashStop(1, 1)] });
        outlines.Add(new Outline(Fill.Gradient("#FF0000", "#0000FF", 0), 8));
        outlines.Add(new Outline(Fill.Pattern(FillPattern.Weave, "#FF0000", "#FFFFFF"), 8));

        foreach (Outline outline in outlines)
        {
            ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, Shape.FromPreset("straightConnector1") with { Outline = outline });
            Outline back = read.Shape!.Outline!;

            Assert.Equal(outline.Dash, back.Dash);
            Assert.Equal(outline.Compound, back.Compound);
            Assert.Equal(outline.Cap, back.Cap);
            Assert.Equal(outline.Join, back.Join);
            Assert.Equal(outline.Head, back.Head);
            Assert.Equal(outline.Tail, back.Tail);
            Assert.Equal(outline.CustomDash, back.CustomDash);
            Assert.Equal(outline.LineFill?.Kind, back.LineFill?.Kind);
        }
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Every_effect_round_trips(string host)
    {
        Shape shape = new(ShapeKind.RoundedRectangle)
        {
            Shadow = new Shadow { ScaleX = 0.9, ScaleY = -0.23, SkewX = -20, SkewY = 5, Alignment = RectangleAlignment.BottomLeft, RotateWithShape = true, Blur = 10, Distance = 4 },
            InnerShadow = new Shadow { Color = "#FF0000" },
            Glow = new GlowEffect("#FFC000", 11, 0.4),
            Reflection = new ReflectionEffect { Size = 0.9, Distance = 4, Blur = 6, StartTransparency = 0.38 },
            Blur = new BlurEffect(3, false),
            PresetShadow = new PresetShadow { Preset = 13, Color = "#808080", Distance = 5, Angle = 120 },
            FillOverlay = new FillOverlayEffect(Fill.Solid("#00FF00", 0.5), BlendMode.Multiply),
            SoftEdges = 2.5,
        };

        ShapeTreeItem read = await DrawingKit.RoundTripAsync(Host(host), shape);

        Assert.Equal(shape.Glow, read.Shape!.Glow);
        Assert.Equal(shape.Blur, read.Shape.Blur);
        Assert.Equal(shape.PresetShadow, read.Shape.PresetShadow);
        Assert.Equal(0.9, read.Shape.Reflection!.Size, 4);
        Assert.Equal(-0.23, read.Shape.Shadow!.ScaleY, 4);
        Assert.Equal(-20, read.Shape.Shadow.SkewX, 4);
        Assert.Equal(BlendMode.Multiply, read.Shape.FillOverlay!.Blend);
        Assert.Equal(2.5, read.Shape.SoftEdges);
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Every_bevel_material_camera_and_light_round_trips(string host)
    {
        ShapeTreeNamespace tree = Host(host);

        foreach (BevelType bevel in Enum.GetValues<BevelType>())
        {
            Shape3D format = new() { TopBevel = new Bevel { Type = bevel, Width = 8, Height = 4 }, BottomBevel = new Bevel { Type = bevel } };

            Shape shape = new(ShapeKind.Rectangle) { Shape3D = format };

            ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, shape);

            Assert.Equal(bevel, read.Shape!.Shape3D!.TopBevel!.Type);
            Assert.Equal(8, read.Shape.Shape3D.TopBevel.Width);
            Assert.Equal(CameraPreset.OrthographicFront, read.Shape.Scene3D!.Camera);
        }

        foreach (MaterialType material in Enum.GetValues<MaterialType>())
        {
            ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, new Shape(ShapeKind.Ellipse) { Shape3D = new Shape3D { Material = material, Depth = 20, DepthColor = "#FF0000", Contour = 2, ContourColor = "#00FF00", Distance = -5 } });

            Assert.Equal(material, read.Shape!.Shape3D!.Material);
            Assert.Equal("#FF0000", read.Shape.Shape3D.DepthColor);
            Assert.Equal(-5, read.Shape.Shape3D.Distance);
        }

        foreach (CameraPreset camera in Enum.GetValues<CameraPreset>())
        {
            ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, new Shape(ShapeKind.Rectangle) { Scene3D = new Scene3D { Camera = camera }, Shape3D = new Shape3D { Depth = 10 } });

            Assert.Equal(camera, read.Shape!.Scene3D!.Camera);
        }

        foreach (LightRigType light in Enum.GetValues<LightRigType>())
        {
            foreach (LightDirection direction in Enum.GetValues<LightDirection>())
            {
                ShapeTreeItem read = await DrawingKit.RoundTripAsync(tree, new Shape(ShapeKind.Rectangle) { Scene3D = new Scene3D { Light = light, LightDirection = direction } });

                Assert.Equal(light, read.Shape!.Scene3D!.Light);
                Assert.Equal(direction, read.Shape.Scene3D.LightDirection);
            }
        }

        Scene3D turned = new()
        {
            Camera = CameraPreset.PerspectiveFront,
            FieldOfView = 45,
            Zoom = 1.5,
            CameraRotation = new Rotation3D(20, 300, 10),
            LightRotation = new Rotation3D(0, 0, 45),
            Backdrop = new Backdrop(new Vector3D(0, 0, -10000), new Vector3D(0, 0, 10000), new Vector3D(0, 10000, 0)),
            KeepTextFlat = true,
        };

        ShapeTreeItem back = await DrawingKit.RoundTripAsync(tree, new Shape(ShapeKind.Rectangle) { Scene3D = turned, Text = new ShapeText("3-D") });

        Assert.Equal(turned, back.Shape!.Scene3D);
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task Styles_locks_and_the_non_visual_flags_round_trip(string host)
    {
        Shape shape = new(ShapeKind.Hexagon)
        {
            Style = new ShapeStyle { Line = new StyleReference(1, "accent2", 0.5), Fill = new StyleReference(3, "accent2"), Effect = new StyleReference(2, "accent2"), Font = ThemeFont.Major, FontColor = "lt1" },
            Locks = new ShapeLocks { NoMove = true, NoResize = true, NoRotate = true, NoTextEdit = true, NoGroup = true },
            IsHidden = true,
            IsDecorative = true,
            Title = "A title",
            Description = "Alternative text",
        };

        ShapeTreeItem read = await DrawingKit.RoundTripAsync(Host(host), shape);

        Assert.Equal(shape.Style, read.Shape!.Style);
        Assert.Equal(shape.Locks, read.Shape.Locks);
        Assert.True(read.Shape.IsHidden);
        Assert.True(read.Shape.IsDecorative);
        Assert.Equal("A title", read.Shape.Title);

        ShapeTreeItem plain = await DrawingKit.RoundTripAsync(Host(host), new Shape(ShapeKind.Hexagon));
        Assert.Null(plain.Shape!.Style);
    }

    [Theory]
    [MemberData(nameof(HostNames))]
    public async Task A_group_keeps_its_turn_flips_locks_and_hidden_flag(string host)
    {
        Shape filled = new(ShapeKind.Rectangle) { Fill = Fill.Group };

        ShapeGroup group = new([GroupItem.Of(filled, DrawingBounds.FromPoints(0, 0, 50, 50)), GroupItem.Of(new Shape(ShapeKind.Ellipse), DrawingBounds.FromPoints(60, 0, 50, 50))])
        {
            Rotation = 15,
            FlipHorizontal = true,
            Locks = new ShapeLocks { NoUngroup = true, NoMove = true },
            IsHidden = true,
        };

        ShapeTreeItem read = await DrawingKit.RoundTripAsync(Host(host), group);

        Assert.Equal(15, read.Group!.Rotation, 6);
        Assert.True(read.Group.FlipHorizontal);
        Assert.True(read.Group.Locks!.NoUngroup);
        Assert.True(read.Group.IsHidden);
        Assert.Equal(FillKind.Group, read.Group.Items[0].Shape!.Fill!.Kind);
    }

    [Fact]
    public void Values_out_of_range_are_argument_exceptions()
    {
        Assert.Throws<ArgumentException>(() => new CustomGeometry([]));
        Assert.Throws<ArgumentException>(() => new GeometryPath(10, 10, []));
        Assert.Throws<ArgumentOutOfRangeException>(() => PathCommand.ArcTo(-1, 1, 0, 90));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ConnectionSite(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new GlowEffect("#FFFFFF", 200));
        Assert.Throws<ArgumentOutOfRangeException>(() => new Bevel { Width = -1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new Rotation3D(400, 0, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new StyleReference(5));
        Assert.Throws<ArgumentException>(() => new StyleReference(1, "purple"));
        Assert.Throws<ArgumentException>(() => new Outline(Fill.None));
        Assert.Throws<ArgumentException>(() => new FillOverlayEffect(Fill.Group));
        Assert.Throws<ArgumentOutOfRangeException>(() => Fill.PathGradient(GradientShape.Linear, null, new GradientStop(0, "#000000"), new GradientStop(1, "#FFFFFF")));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PresetShadow { Preset = 21 });
        Assert.Throws<ArgumentException>(() => new Shape3D { DepthColor = "red" });
    }

    [Fact]
    public async Task A_hostile_custom_geometry_and_style_are_skipped_not_thrown()
    {
        string xml = """
            <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:nvSpPr><p:cNvPr id="2" name="x"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr>
                <a:custGeom><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="w" y="0"/></a:moveTo></a:path></a:pathLst></a:custGeom>
                <a:scene3d><a:camera prst="nonsense" fov="-5" zoom="99999999"/><a:lightRig rig="nope" dir="up"/></a:scene3d>
                <a:sp3d extrusionH="-100" prstMaterial="gold"><a:bevelT w="999999999999" prst="x"/></a:sp3d>
              </p:spPr>
              <p:style><a:lnRef idx="9"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:srgbClr val="FF0000"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"/></p:style>
            </p:sp>
            """;

        Result<ShapeTreeItem> read = await DrawingReader.ReadAsync(System.Xml.Linq.XElement.Parse(xml), DrawingKit.ReadContext(new DrawingContext(DrawingKit.Slide, [], 2)));

        Assert.True(read.IsSuccess(out ShapeTreeItem? item), read.Error);
        Assert.Null(item.Shape!.Geometry);
        Assert.Null(item.Shape.Style);
        Assert.Equal(CameraPreset.OrthographicFront, item.Shape.Scene3D!.Camera);
        Assert.Equal(0, item.Shape.Shape3D!.Depth);
        Assert.Equal(1584, item.Shape.Shape3D.TopBevel!.Width);
    }

    private static ShapeTreeNamespace Host(string name)
    {
        if (name == "xdr")
        {
            return DrawingKit.Spreadsheet;
        }

        return DrawingKit.Slide;
    }
}
