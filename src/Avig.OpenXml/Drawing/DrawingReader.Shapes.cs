using System.Xml.Linq;

namespace Avig.OpenXml.Drawing;

public static partial class DrawingReader
{
    /// <summary>The most path commands read from one geometry, so that a hostile file cannot fill the memory.</summary>
    private const int MaxPathCommands = 100_000;

    /// <summary>The geometry, line, effect and 3-D reading that shapes have beside the basics.</summary>
    private sealed partial class Reader
    {
        /// <summary>
        /// An a:custGeom whose paths hold numbers only (as Office writes freeforms); a path that
        /// refers to guides by name cannot be held, and the geometry is then null.
        /// </summary>
        public static bool TryCustomGeometry(XElement? element, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out CustomGeometry? geometry)
        {
            geometry = null;

            if (element is null)
            {
                return false;
            }

            List<GeometryPath> paths = [];
            int total = 0;

            foreach (XElement path in Children(Child(element, "pathLst"), "path"))
            {
                if (paths.Count >= CustomGeometry.MaxPaths)
                {
                    break;
                }

                List<PathCommand> commands = [];

                foreach (XElement command in path.Elements())
                {
                    if (total >= MaxPathCommands)
                    {
                        break;
                    }

                    if (TryCommand(command, out PathCommand? read) == false)
                    {
                        return false;
                    }

                    commands.Add(read);
                    total++;
                }

                if (commands.Count == 0)
                {
                    continue;
                }

                long width = Math.Clamp(Long(path.Attribute("w")?.Value), 0, DrawingValues.MaxCoordinate);
                long height = Math.Clamp(Long(path.Attribute("h")?.Value), 0, DrawingValues.MaxCoordinate);
                paths.Add(new GeometryPath(width, height, commands)
                {
                    Fill = ShapeNames.PathFillOf(path.Attribute("fill")?.Value),
                    Stroke = DrawingValues.Boolean(path.Attribute("stroke")?.Value, true),
                    ExtrusionOk = DrawingValues.Boolean(path.Attribute("extrusionOk")?.Value, true),
                });
            }

            if (paths.Count == 0)
            {
                return false;
            }

            geometry = new CustomGeometry(paths);

            return true;
        }

        private static bool TryCommand(XElement element, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out PathCommand? command)
        {
            command = null;
            List<PathPoint> points = [];

            foreach (XElement point in Children(element, "pt"))
            {
                if (points.Count >= 3 || TryPathCoordinate(point.Attribute("x")?.Value, out long left) == false || TryPathCoordinate(point.Attribute("y")?.Value, out long top) == false)
                {
                    return false;
                }

                points.Add(new PathPoint(left, top));
            }

            switch (element.Name.LocalName)
            {
                case "moveTo" when points.Count == 1:
                    command = PathCommand.MoveTo(points[0].Left, points[0].Top);

                    return true;

                case "lnTo" when points.Count == 1:
                    command = PathCommand.LineTo(points[0].Left, points[0].Top);

                    return true;

                case "cubicBezTo" when points.Count == 3:
                    command = PathCommand.CubicBezierTo(points[0], points[1], points[2]);

                    return true;

                case "quadBezTo" when points.Count == 2:
                    command = PathCommand.QuadraticBezierTo(points[0], points[1]);

                    return true;

                case "close":
                    command = PathCommand.Close();

                    return true;

                case "arcTo":
                    if (TryPathCoordinate(element.Attribute("wR")?.Value, out long widthRadius) == false || TryPathCoordinate(element.Attribute("hR")?.Value, out long heightRadius) == false
                        || DrawingValues.TryLong(element.Attribute("stAng")?.Value, out long start) == false || DrawingValues.TryLong(element.Attribute("swAng")?.Value, out long sweep) == false)
                    {
                        return false;
                    }

                    command = PathCommand.ArcTo(Math.Max(widthRadius, 0), Math.Max(heightRadius, 0), Math.Clamp(start / 60_000.0, -360, 360), Math.Clamp(sweep / 60_000.0, -360, 360));

                    return true;

                default:
                    return false;
            }
        }

        private static bool TryPathCoordinate(string? text, out long value)
        {
            if (DrawingValues.TryLong(text, out value) == false)
            {
                return false;
            }

            value = Math.Clamp(value, -DrawingValues.MaxCoordinate, DrawingValues.MaxCoordinate);

            return true;
        }

        /// <summary>The cap, compound, alignment, dash, join and arrowheads of an a:ln, onto a line read without them.</summary>
        public static Outline LineDetails(XElement line, Outline outline)
        {
            LineAlignment alignment = LineAlignment.Center;

            if (line.Attribute("algn")?.Value == "in")
            {
                alignment = LineAlignment.Inset;
            }

            Outline read = outline with
            {
                Compound = ShapeNames.CompoundOf(line.Attribute("cmpd")?.Value),
                Alignment = alignment,
                Head = ReadLineEnd(Child(line, "headEnd")),
                Tail = ReadLineEnd(Child(line, "tailEnd")),
            };

            if (ShapeNames.TryCapOf(line.Attribute("cap")?.Value, out LineCap cap) == true)
            {
                read = read with { Cap = cap };
            }

            if (Child(line, "round") is not null)
            {
                read = read with { Join = LineJoin.Round };
            }
            else if (Child(line, "bevel") is not null)
            {
                read = read with { Join = LineJoin.Bevel };
            }
            else if (Child(line, "miter") is XElement miter)
            {
                read = read with { Join = LineJoin.Miter };

                if (DrawingValues.TryLong(miter.Attribute("lim")?.Value, out long limit) == true)
                {
                    read = read with { MiterLimit = Math.Clamp(limit / 100_000.0, 0, 1000) };
                }
            }

            List<DashStop> stops = [];

            foreach (XElement stop in Children(Child(line, "custDash"), "ds"))
            {
                if (stops.Count >= 100)
                {
                    break;
                }

                stops.Add(new DashStop(Math.Clamp(Long(stop.Attribute("d")?.Value) / 100_000.0, 0, 1000), Math.Clamp(Long(stop.Attribute("sp")?.Value) / 100_000.0, 0, 1000)));
            }

            if (stops.Count > 0)
            {
                read = read with { CustomDash = stops };
            }

            return read;
        }

        private static LineEnd? ReadLineEnd(XElement? element)
        {
            LineEnd? end = null;

            if (element is not null)
            {
                end = new LineEnd(ShapeNames.EndTypeOf(element.Attribute("type")?.Value), ShapeNames.EndSizeOf(element.Attribute("w")?.Value), ShapeNames.EndSizeOf(element.Attribute("len")?.Value));
            }

            return end;
        }

        /// <summary>The effects beside the shadows and soft edges.</summary>
        public void MoreEffects(XElement? effects, out GlowEffect? glow, out ReflectionEffect? reflection, out BlurEffect? blur, out PresetShadow? presetShadow, out FillOverlayEffect? overlay)
        {
            glow = null;
            reflection = null;
            blur = null;
            presetShadow = null;
            overlay = null;

            if (Child(effects, "glow") is XElement glowElement && TryColor(glowElement, out string? glowColor, out double glowTransparency) == true)
            {
                glow = new GlowEffect(glowColor, Math.Clamp(DrawingValues.Points(Long(glowElement.Attribute("rad")?.Value)), 0, 150), glowTransparency);
            }

            if (Child(effects, "reflection") is XElement mirror)
            {
                reflection = new ReflectionEffect
                {
                    Blur = Math.Clamp(DrawingValues.Points(Long(mirror.Attribute("blurRad")?.Value)), 0, 100),
                    StartTransparency = 1 - Fraction(mirror.Attribute("stA")?.Value, 1),
                    StartPosition = Fraction(mirror.Attribute("stPos")?.Value, 0),
                    EndTransparency = 1 - Fraction(mirror.Attribute("endA")?.Value, 0),
                    Size = Fraction(mirror.Attribute("endPos")?.Value, 1),
                    Distance = Math.Clamp(DrawingValues.Points(Long(mirror.Attribute("dist")?.Value)), 0, 200),
                    Direction = Degrees(mirror.Attribute("dir")?.Value),
                    FadeDirection = Degrees(mirror.Attribute("fadeDir")?.Value ?? "5400000"),
                    ScaleX = Scale(mirror.Attribute("sx")?.Value),
                    ScaleY = Scale(mirror.Attribute("sy")?.Value),
                    Alignment = ShapeNames.AlignmentOf(mirror.Attribute("algn")?.Value, RectangleAlignment.BottomLeft),
                    RotateWithShape = DrawingValues.Boolean(mirror.Attribute("rotWithShape")?.Value, true),
                };
            }

            if (Child(effects, "blur") is XElement blurElement)
            {
                blur = new BlurEffect(Math.Clamp(DrawingValues.Points(Long(blurElement.Attribute("rad")?.Value)), 0, 100), DrawingValues.Boolean(blurElement.Attribute("grow")?.Value, true));
            }

            if (Child(effects, "prstShdw") is XElement preset && TryColor(preset, out string? presetColor, out double presetTransparency) == true)
            {
                string name = preset.Attribute("prst")?.Value ?? "shdw1";
                int number = 1;

                if (name.StartsWith("shdw", StringComparison.Ordinal) == true && int.TryParse(name[4..], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int parsed) == true)
                {
                    number = Math.Clamp(parsed, 1, 20);
                }

                presetShadow = new PresetShadow
                {
                    Preset = number,
                    Color = presetColor,
                    Transparency = presetTransparency,
                    Distance = Math.Clamp(DrawingValues.Points(Long(preset.Attribute("dist")?.Value)), 0, 200),
                    Angle = Degrees(preset.Attribute("dir")?.Value),
                };
            }

            if (Child(effects, "fillOverlay") is XElement fillOverlay && TryFill(fillOverlay, out Fill? overlayFill) == true && (overlayFill.Kind is FillKind.Picture or FillKind.Group or FillKind.Background) == false)
            {
                overlay = new FillOverlayEffect(overlayFill, ShapeNames.BlendOf(fillOverlay.Attribute("blend")?.Value));
            }
        }

        /// <summary>The scale, skew, alignment and turn of an a:outerShdw, onto a shadow read without them.</summary>
        public static Shadow ShadowDetails(XElement element, Shadow shadow)
        {
            return shadow with
            {
                ScaleX = Scale(element.Attribute("sx")?.Value),
                ScaleY = Scale(element.Attribute("sy")?.Value),
                SkewX = Math.Clamp(Long(element.Attribute("kx")?.Value) / 60_000.0, -90, 90),
                SkewY = Math.Clamp(Long(element.Attribute("ky")?.Value) / 60_000.0, -90, 90),
                Alignment = ShapeNames.AlignmentOf(element.Attribute("algn")?.Value, RectangleAlignment.Bottom),
                RotateWithShape = DrawingValues.Boolean(element.Attribute("rotWithShape")?.Value, true),
            };
        }

        /// <summary>The a:scene3d and a:sp3d of shape properties.</summary>
        public void ThreeD(XElement? properties, out Scene3D? scene, out Shape3D? shape)
        {
            scene = null;
            shape = null;

            if (Child(properties, "scene3d") is XElement sceneElement)
            {
                XElement? camera = Child(sceneElement, "camera");
                XElement? light = Child(sceneElement, "lightRig");
                scene = new Scene3D
                {
                    Camera = ShapeNames.CameraOf(camera?.Attribute("prst")?.Value),
                    Zoom = Math.Clamp(Fraction(camera?.Attribute("zoom")?.Value, 1, 10), 0, 10),
                    CameraRotation = ReadRotation(Child(camera, "rot")),
                    Light = ShapeNames.LightOf(light?.Attribute("rig")?.Value),
                    LightDirection = ShapeNames.DirectionOf(light?.Attribute("dir")?.Value),
                    LightRotation = ReadRotation(Child(light, "rot")),
                    Backdrop = ReadBackdrop(Child(sceneElement, "backdrop")),
                };

                if (DrawingValues.TryLong(camera?.Attribute("fov")?.Value, out long fieldOfView) == true)
                {
                    scene = scene with { FieldOfView = Math.Clamp(fieldOfView / 60_000.0, 0, 180) };
                }
            }

            if (Child(properties, "sp3d") is XElement shapeElement)
            {
                shape = new Shape3D
                {
                    Distance = Math.Clamp(DrawingValues.Points(Long(shapeElement.Attribute("z")?.Value)), -1584, 1584),
                    Depth = Math.Clamp(DrawingValues.Points(Long(shapeElement.Attribute("extrusionH")?.Value)), 0, 1584),
                    Contour = Math.Clamp(DrawingValues.Points(Long(shapeElement.Attribute("contourW")?.Value)), 0, 1584),
                    Material = ShapeNames.MaterialOf(shapeElement.Attribute("prstMaterial")?.Value),
                    TopBevel = ReadBevel(Child(shapeElement, "bevelT")),
                    BottomBevel = ReadBevel(Child(shapeElement, "bevelB")),
                };

                if (TryColor(Child(shapeElement, "extrusionClr"), out string? depthColor, out _) == true)
                {
                    shape = shape with { DepthColor = depthColor };
                }

                if (TryColor(Child(shapeElement, "contourClr"), out string? contourColor, out _) == true)
                {
                    shape = shape with { ContourColor = contourColor };
                }
            }
        }

        /// <summary>The p:style of a shape when it is one a <see cref="ShapeStyle"/> holds (theme colours only).</summary>
        public static bool TryStyle(XElement? element, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out ShapeStyle? style)
        {
            style = null;

            if (element is null || TryReference(Child(element, "lnRef"), out StyleReference? line) == false || TryReference(Child(element, "fillRef"), out StyleReference? fill) == false
                || TryReference(Child(element, "effectRef"), out StyleReference? effect) == false)
            {
                return false;
            }

            XElement? font = Child(element, "fontRef");
            string fontColor = Child(font, "schemeClr")?.Attribute("val")?.Value ?? "tx1";

            try
            {
                style = new ShapeStyle
                {
                    Line = line,
                    Fill = fill,
                    Effect = effect,
                    Font = ShapeNames.FontOf(font?.Attribute("idx")?.Value),
                    FontColor = fontColor,
                };
            }
            catch (ArgumentException)
            {
                return false;
            }

            return true;
        }

        private static bool TryReference(XElement? element, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out StyleReference? reference)
        {
            reference = null;
            XElement? color = Child(element, "schemeClr");

            if (element is null || color is null || int.TryParse(element.Attribute("idx")?.Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int index) == false)
            {
                return false;
            }

            if (((index >= 0 && index <= 3) || (index >= 1001 && index <= 1003)) == false)
            {
                return false;
            }

            string name = color.Attribute("val")?.Value ?? string.Empty;

            if (color.Elements().Any(child => child.Name.LocalName != "shade") == true)
            {
                return false;
            }

            double? shade = null;

            if (Child(color, "shade") is XElement shadeElement)
            {
                shade = Math.Clamp(Long(shadeElement.Attribute("val")?.Value) / 100_000.0, 0, 1);
            }

            try
            {
                reference = new StyleReference(index, name, shade);
            }
            catch (ArgumentException)
            {
                return false;
            }

            return true;
        }

        /// <summary>The a:spLocks, a:cxnSpLocks or a:grpSpLocks inside a cNv…Pr, or null when it locks nothing.</summary>
        public static ShapeLocks? Locks(XElement? element)
        {
            ShapeLocks? found = null;

            if (element is not null)
            {
                ShapeLocks locks = new()
                {
                    NoGroup = Flag(element, "noGrp"),
                    NoUngroup = Flag(element, "noUngrp"),
                    NoSelect = Flag(element, "noSelect"),
                    NoRotate = Flag(element, "noRot"),
                    NoChangeAspect = Flag(element, "noChangeAspect"),
                    NoMove = Flag(element, "noMove"),
                    NoResize = Flag(element, "noResize"),
                    NoEditPoints = Flag(element, "noEditPoints"),
                    NoAdjustHandles = Flag(element, "noAdjustHandles"),
                    NoChangeArrowheads = Flag(element, "noChangeArrowheads"),
                    NoChangeShapeType = Flag(element, "noChangeShapeType"),
                    NoTextEdit = Flag(element, "noTextEdit"),
                };

                if (locks.Any == true)
                {
                    found = locks;
                }
            }

            return found;
        }

        /// <summary>A connector's a:stCxn or a:endCxn, or null.</summary>
        public static ConnectionSite? Connection(XElement? element)
        {
            ConnectionSite? found = null;

            if (element is not null
                && int.TryParse(element.Attribute("id")?.Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int id) == true
                && int.TryParse(element.Attribute("idx")?.Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int site) == true
                && id >= 1 && site <= 1000)
            {
                found = new ConnectionSite(id, site);
            }

            return found;
        }

        /// <summary>True when the cNvPr marks the drawing decorative (adec:decorative val="1").</summary>
        public static bool IsDecorative(XElement? nonVisual)
        {
            XElement? decorative = Child(nonVisual, "extLst")?.Descendants().FirstOrDefault(child => child.Name.LocalName == "decorative");

            return decorative is not null && DrawingValues.Boolean(decorative.Attribute("val")?.Value, false) == true;
        }

        /// <summary>The path, focus, tile rectangle, turn and scaling of an a:gradFill, onto a linear gradient read without them.</summary>
        public static Fill GradientDetails(XElement element, Fill gradient)
        {
            Fill read = gradient;

            if (Child(element, "path") is XElement path)
            {
                GradientStop[] stops = [.. gradient.Stops];
                read = Fill.PathGradient(ShapeNames.GradientPathOf(path.Attribute("path")?.Value), ReadRectangle(Child(path, "fillToRect")), stops);
            }
            else if (DrawingValues.Boolean(Child(element, "lin")?.Attribute("scaled")?.Value, false) == true)
            {
                read = read.WithScaledAngle();
            }

            if (ReadRectangle(Child(element, "tileRect")) is FillRectangle tile)
            {
                read = read.WithTileRectangle(tile);
            }

            if (DrawingValues.Boolean(element.Attribute("rotWithShape")?.Value, true) == false)
            {
                read = read.WithoutRotation();
            }

            return read;
        }

        /// <summary>The tiling and turn of an a:blipFill, onto a picture fill read without them.</summary>
        public static Fill PictureDetails(XElement element, Fill picture)
        {
            Fill read = picture;

            if (Child(element, "tile") is XElement tile)
            {
                read = Fill.TiledPicture(picture.Image!, new PictureTiling
                {
                    OffsetX = Length.FromEmu(Coordinate(tile.Attribute("tx")?.Value)),
                    OffsetY = Length.FromEmu(Coordinate(tile.Attribute("ty")?.Value)),
                    ScaleX = Math.Clamp(Fraction(tile.Attribute("sx")?.Value, 1, 100), 0, 100),
                    ScaleY = Math.Clamp(Fraction(tile.Attribute("sy")?.Value, 1, 100), 0, 100),
                    Flip = ShapeNames.FlipOf(tile.Attribute("flip")?.Value),
                    Alignment = ShapeNames.AlignmentOf(tile.Attribute("algn")?.Value, RectangleAlignment.TopLeft),
                });
            }

            if (DrawingValues.Boolean(element.Attribute("rotWithShape")?.Value, true) == false)
            {
                read = read.WithoutRotation();
            }

            return read;
        }

        private static FillRectangle? ReadRectangle(XElement? element)
        {
            FillRectangle? found = null;

            if (element is not null)
            {
                found = new FillRectangle(RectangleInset(element, "l"), RectangleInset(element, "t"), RectangleInset(element, "r"), RectangleInset(element, "b"));
            }

            return found;
        }

        private static double RectangleInset(XElement element, string name)
        {
            return Math.Clamp(Long(element.Attribute(name)?.Value) / 100_000.0, -100, 100);
        }

        private static bool Flag(XElement element, string name)
        {
            return DrawingValues.Boolean(element.Attribute(name)?.Value, false);
        }

        private static Bevel? ReadBevel(XElement? element)
        {
            Bevel? found = null;

            if (element is not null)
            {
                found = new Bevel
                {
                    Width = Math.Clamp(DrawingValues.Points(Long(element.Attribute("w")?.Value ?? "76200")), 0, 1584),
                    Height = Math.Clamp(DrawingValues.Points(Long(element.Attribute("h")?.Value ?? "76200")), 0, 1584),
                    Type = ShapeNames.BevelOf(element.Attribute("prst")?.Value),
                };
            }

            return found;
        }

        private static Rotation3D? ReadRotation(XElement? element)
        {
            Rotation3D? found = null;

            if (element is not null)
            {
                found = new Rotation3D(Turn(element.Attribute("lat")?.Value), Turn(element.Attribute("lon")?.Value), Turn(element.Attribute("rev")?.Value));
            }

            return found;
        }

        private static Backdrop? ReadBackdrop(XElement? element)
        {
            Backdrop? found = null;

            if (element is not null)
            {
                found = new Backdrop(Vector(Child(element, "anchor"), "x", "y", "z"), Vector(Child(element, "norm"), "dx", "dy", "dz"), Vector(Child(element, "up"), "dx", "dy", "dz"));
            }

            return found;
        }

        private static Vector3D Vector(XElement? element, string across, string down, string toward)
        {
            return new Vector3D(Coordinate(element?.Attribute(across)?.Value), Coordinate(element?.Attribute(down)?.Value), Coordinate(element?.Attribute(toward)?.Value));
        }

        /// <summary>An angle in 60 000ths of a degree, turned into 0 to 360.</summary>
        private static double Turn(string? text)
        {
            double degrees = (Long(text) % 21_600_000) / 60_000.0;

            if (degrees < 0)
            {
                degrees += 360;
            }

            return Math.Clamp(degrees, 0, 360);
        }

        private static double Degrees(string? text)
        {
            return Math.Clamp(Long(text) / 60_000.0, -360, 360);
        }

        /// <summary>A percentage in 100 000ths as a fraction, clamped to 0 and <paramref name="maximum"/>; the fallback when missing.</summary>
        private static double Fraction(string? text, double fallback, double maximum = 1)
        {
            if (DrawingValues.TryLong(text, out long value) == false)
            {
                return fallback;
            }

            return Math.Clamp(value / 100_000.0, 0, maximum);
        }

        private static double Scale(string? text)
        {
            if (DrawingValues.TryLong(text, out long value) == false)
            {
                return 1;
            }

            return Math.Clamp(value / 100_000.0, -2, 2);
        }
    }
}
