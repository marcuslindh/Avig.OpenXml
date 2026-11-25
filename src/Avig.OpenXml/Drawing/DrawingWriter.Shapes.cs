using System.Xml;

namespace Avig.OpenXml.Drawing;

/// <summary>Every effect a shape can have, as one value for the effect writer.</summary>
internal readonly record struct EffectSet(
    Shadow? Outer,
    Shadow? Inner,
    double? SoftEdges,
    GlowEffect? Glow = null,
    ReflectionEffect? Reflection = null,
    BlurEffect? Blur = null,
    PresetShadow? PresetShadow = null,
    FillOverlayEffect? FillOverlay = null)
{
    public bool Any => Outer is not null || Inner is not null || SoftEdges is not null || Glow is not null || Reflection is not null || Blur is not null || PresetShadow is not null || FillOverlay is not null;

    public static EffectSet Of(Shape shape)
    {
        return new EffectSet(shape.Shadow, shape.InnerShadow, shape.SoftEdges, shape.Glow, shape.Reflection, shape.Blur, shape.PresetShadow, shape.FillOverlay);
    }
}

/// <summary>What a cNvPr holds beside its id, name, alternative text and link.</summary>
internal readonly record struct NonVisualExtras(bool Hidden, string? Title, bool Decorative)
{
    public static NonVisualExtras None => default;
}

public static partial class DrawingWriter
{
    private const string DecorativeExtensionUri = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";
    private const string DecorativeNamespace = "http://schemas.microsoft.com/office/drawing/2017/decorative";

    /// <summary>An a:custGeom: the paths, with the guides and the text box DrawingML requires around them.</summary>
    internal static void WriteCustomGeometry(XmlWriter xml, CustomGeometry geometry)
    {
        xml.WriteStartElement("a", "custGeom", Main);
        xml.WriteElementString("a", "avLst", Main, string.Empty);
        xml.WriteElementString("a", "gdLst", Main, string.Empty);
        xml.WriteElementString("a", "ahLst", Main, string.Empty);
        xml.WriteElementString("a", "cxnLst", Main, string.Empty);
        xml.WriteStartElement("a", "rect", Main);
        xml.WriteAttributeString(null, "l", null, "l");
        xml.WriteAttributeString(null, "t", null, "t");
        xml.WriteAttributeString(null, "r", null, "r");
        xml.WriteAttributeString(null, "b", null, "b");
        xml.WriteEndElement();
        xml.WriteStartElement("a", "pathLst", Main);

        foreach (GeometryPath path in geometry.Paths)
        {
            xml.WriteStartElement("a", "path", Main);

            if (path.Width != 0)
            {
                xml.WriteAttributeString(null, "w", null, DrawingValues.Text(path.Width));
            }

            if (path.Height != 0)
            {
                xml.WriteAttributeString(null, "h", null, DrawingValues.Text(path.Height));
            }

            if (path.Fill != PathFill.Normal)
            {
                xml.WriteAttributeString(null, "fill", null, ShapeNames.PathFill(path.Fill));
            }

            if (path.Stroke == false)
            {
                xml.WriteAttributeString(null, "stroke", null, "0");
            }

            if (path.ExtrusionOk == false)
            {
                xml.WriteAttributeString(null, "extrusionOk", null, "0");
            }

            foreach (PathCommand command in path.Commands)
            {
                WritePathCommand(xml, command);
            }

            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    private static void WritePathCommand(XmlWriter xml, PathCommand command)
    {
        string name = command.Kind switch
        {
            PathCommandKind.MoveTo => "moveTo",
            PathCommandKind.LineTo => "lnTo",
            PathCommandKind.CubicBezierTo => "cubicBezTo",
            PathCommandKind.QuadraticBezierTo => "quadBezTo",
            PathCommandKind.ArcTo => "arcTo",
            _ => "close",
        };

        xml.WriteStartElement("a", name, Main);

        if (command.Kind == PathCommandKind.ArcTo)
        {
            xml.WriteAttributeString(null, "wR", null, DrawingValues.Text(command.WidthRadius));
            xml.WriteAttributeString(null, "hR", null, DrawingValues.Text(command.HeightRadius));
            xml.WriteAttributeString(null, "stAng", null, DrawingValues.Text((long)Math.Round(command.StartAngle * 60_000)));
            xml.WriteAttributeString(null, "swAng", null, DrawingValues.Text((long)Math.Round(command.SweepAngle * 60_000)));
        }

        foreach (PathPoint point in command.Points)
        {
            xml.WriteStartElement("a", "pt", Main);
            xml.WriteAttributeString(null, "x", null, DrawingValues.Text(point.Left));
            xml.WriteAttributeString(null, "y", null, DrawingValues.Text(point.Top));
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    /// <summary>The a:effectLst of every effect given, in the schema's order.</summary>
    internal static void WriteEffectList(XmlWriter xml, EffectSet effects)
    {
        xml.WriteStartElement("a", "effectLst", Main);

        if (effects.Blur is BlurEffect blur)
        {
            xml.WriteStartElement("a", "blur", Main);
            xml.WriteAttributeString(null, "rad", null, DrawingValues.Text(DrawingValues.Emu(blur.Radius)));

            if (blur.Grow == false)
            {
                xml.WriteAttributeString(null, "grow", null, "0");
            }

            xml.WriteEndElement();
        }

        if (effects.FillOverlay is FillOverlayEffect overlay)
        {
            xml.WriteStartElement("a", "fillOverlay", Main);
            xml.WriteAttributeString(null, "blend", null, ShapeNames.Blend(overlay.Blend));
            WriteFill(xml, overlay.Fill, null);
            xml.WriteEndElement();
        }

        if (effects.Glow is GlowEffect glow)
        {
            xml.WriteStartElement("a", "glow", Main);
            xml.WriteAttributeString(null, "rad", null, DrawingValues.Text(DrawingValues.Emu(glow.Radius)));
            WriteColor(xml, glow.Color, glow.Transparency);
            xml.WriteEndElement();
        }

        if (effects.Inner is Shadow inner)
        {
            xml.WriteStartElement("a", "innerShdw", Main);
            WriteShadowAttributes(xml, inner);
            WriteColor(xml, inner.Color, inner.Transparency);
            xml.WriteEndElement();
        }

        if (effects.Outer is Shadow outer)
        {
            xml.WriteStartElement("a", "outerShdw", Main);
            WriteShadowAttributes(xml, outer);
            WriteScale(xml, "sx", outer.ScaleX);
            WriteScale(xml, "sy", outer.ScaleY);
            WriteSkew(xml, "kx", outer.SkewX);
            WriteSkew(xml, "ky", outer.SkewY);
            xml.WriteAttributeString(null, "algn", null, ShapeNames.Alignment(outer.Alignment));
            xml.WriteAttributeString(null, "rotWithShape", null, DrawingValues.Flag(outer.RotateWithShape));
            WriteColor(xml, outer.Color, outer.Transparency);
            xml.WriteEndElement();
        }

        if (effects.PresetShadow is PresetShadow preset)
        {
            xml.WriteStartElement("a", "prstShdw", Main);
            xml.WriteAttributeString(null, "prst", null, "shdw" + DrawingValues.Text(preset.Preset));
            xml.WriteAttributeString(null, "dist", null, DrawingValues.Text(DrawingValues.Emu(preset.Distance)));
            xml.WriteAttributeString(null, "dir", null, DrawingValues.Text(DrawingValues.Angle(preset.Angle)));
            WriteColor(xml, preset.Color, preset.Transparency);
            xml.WriteEndElement();
        }

        if (effects.Reflection is ReflectionEffect reflection)
        {
            xml.WriteStartElement("a", "reflection", Main);
            xml.WriteAttributeString(null, "blurRad", null, DrawingValues.Text(DrawingValues.Emu(reflection.Blur)));
            xml.WriteAttributeString(null, "stA", null, DrawingValues.Text(DrawingValues.Percent(1 - reflection.StartTransparency)));
            xml.WriteAttributeString(null, "stPos", null, DrawingValues.Text(DrawingValues.Percent(reflection.StartPosition)));
            xml.WriteAttributeString(null, "endA", null, DrawingValues.Text(DrawingValues.Percent(1 - reflection.EndTransparency)));
            xml.WriteAttributeString(null, "endPos", null, DrawingValues.Text(DrawingValues.Percent(reflection.Size)));
            xml.WriteAttributeString(null, "dist", null, DrawingValues.Text(DrawingValues.Emu(reflection.Distance)));
            xml.WriteAttributeString(null, "dir", null, DrawingValues.Text(DrawingValues.Angle(reflection.Direction)));
            xml.WriteAttributeString(null, "fadeDir", null, DrawingValues.Text(DrawingValues.Angle(reflection.FadeDirection)));
            WriteScale(xml, "sx", reflection.ScaleX);
            WriteScale(xml, "sy", reflection.ScaleY);
            xml.WriteAttributeString(null, "algn", null, ShapeNames.Alignment(reflection.Alignment));
            xml.WriteAttributeString(null, "rotWithShape", null, DrawingValues.Flag(reflection.RotateWithShape));
            xml.WriteEndElement();
        }

        if (effects.SoftEdges is double points)
        {
            xml.WriteStartElement("a", "softEdge", Main);
            xml.WriteAttributeString(null, "rad", null, DrawingValues.Text(DrawingValues.Emu(points)));
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    /// <summary>The a:scene3d and a:sp3d of a shape; a 3-D format without a scene gets Office's flat front view, as Office writes it.</summary>
    internal static void WriteThreeD(XmlWriter xml, Scene3D? scene, Shape3D? shape)
    {
        if (scene is null && shape is null)
        {
            return;
        }

        scene ??= new Scene3D();
        xml.WriteStartElement("a", "scene3d", Main);
        xml.WriteStartElement("a", "camera", Main);
        xml.WriteAttributeString(null, "prst", null, ShapeNames.Camera(scene.Camera));

        if (scene.FieldOfView is double fieldOfView)
        {
            xml.WriteAttributeString(null, "fov", null, DrawingValues.Text((long)Math.Round(fieldOfView * 60_000)));
        }

        if (scene.Zoom != 1)
        {
            xml.WriteAttributeString(null, "zoom", null, DrawingValues.Text(DrawingValues.Percent(scene.Zoom)));
        }

        WriteRotation3D(xml, scene.CameraRotation);
        xml.WriteEndElement();
        xml.WriteStartElement("a", "lightRig", Main);
        xml.WriteAttributeString(null, "rig", null, ShapeNames.Light(scene.Light));
        xml.WriteAttributeString(null, "dir", null, ShapeNames.Direction(scene.LightDirection));
        WriteRotation3D(xml, scene.LightRotation);
        xml.WriteEndElement();

        if (scene.Backdrop is Backdrop backdrop)
        {
            xml.WriteStartElement("a", "backdrop", Main);
            WriteVector(xml, "anchor", "x", "y", "z", backdrop.Anchor);
            WriteVector(xml, "norm", "dx", "dy", "dz", backdrop.Normal);
            WriteVector(xml, "up", "dx", "dy", "dz", backdrop.Up);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();

        if (shape is null)
        {
            return;
        }

        xml.WriteStartElement("a", "sp3d", Main);

        if (shape.Distance != 0)
        {
            xml.WriteAttributeString(null, "z", null, DrawingValues.Text(DrawingValues.Emu(shape.Distance)));
        }

        if (shape.Depth != 0)
        {
            xml.WriteAttributeString(null, "extrusionH", null, DrawingValues.Text(DrawingValues.Emu(shape.Depth)));
        }

        if (shape.Contour != 0)
        {
            xml.WriteAttributeString(null, "contourW", null, DrawingValues.Text(DrawingValues.Emu(shape.Contour)));
        }

        if (shape.Material != MaterialType.WarmMatte)
        {
            xml.WriteAttributeString(null, "prstMaterial", null, ShapeNames.Material(shape.Material));
        }

        WriteBevel(xml, "bevelT", shape.TopBevel);
        WriteBevel(xml, "bevelB", shape.BottomBevel);

        if (shape.DepthColor is string depthColor)
        {
            xml.WriteStartElement("a", "extrusionClr", Main);
            WriteColor(xml, depthColor, 0);
            xml.WriteEndElement();
        }

        if (shape.ContourColor is string contourColor)
        {
            xml.WriteStartElement("a", "contourClr", Main);
            WriteColor(xml, contourColor, 0);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    /// <summary>The p:style of a shape: its own, or the one Office gives a new shape of its kind.</summary>
    internal static void WriteShapeStyle(XmlWriter xml, ShapeTreeNamespace tree, ShapeStyle style)
    {
        xml.WriteStartElement(tree.Prefix, "style", tree.Uri);
        WriteReference(xml, "lnRef", style.Line);
        WriteReference(xml, "fillRef", style.Fill);
        WriteReference(xml, "effectRef", style.Effect);
        xml.WriteStartElement("a", "fontRef", Main);
        xml.WriteAttributeString(null, "idx", null, ShapeNames.Font(style.Font));
        xml.WriteStartElement("a", "schemeClr", Main);
        xml.WriteAttributeString(null, "val", null, style.FontColor);
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    /// <summary>The style Office gives a new shape: a connector's line in accent1, a text box nothing, any other shape the accent fill.</summary>
    internal static ShapeStyle DefaultStyle(Shape shape)
    {
        if (shape.IsConnector == true)
        {
            return new ShapeStyle { Line = new StyleReference(1), Fill = new StyleReference(0), Effect = new StyleReference(0), FontColor = "tx1" };
        }

        if (shape.Kind == ShapeKind.TextBox)
        {
            return new ShapeStyle { Line = new StyleReference(0), Fill = new StyleReference(0), Effect = new StyleReference(0), FontColor = "dk1" };
        }

        return new ShapeStyle();
    }

    /// <summary>An a:spLocks, a:cxnSpLocks or a:grpSpLocks with the locks that element takes.</summary>
    internal static void WriteLocks(XmlWriter xml, string element, ShapeLocks? locks)
    {
        if (locks is null || locks.Any == false)
        {
            return;
        }

        xml.WriteStartElement("a", element, Main);
        WriteLock(xml, "noGrp", locks.NoGroup);

        if (element == "grpSpLocks")
        {
            WriteLock(xml, "noUngrp", locks.NoUngroup);
        }

        WriteLock(xml, "noSelect", locks.NoSelect);
        WriteLock(xml, "noRot", locks.NoRotate);
        WriteLock(xml, "noChangeAspect", locks.NoChangeAspect);
        WriteLock(xml, "noMove", locks.NoMove);
        WriteLock(xml, "noResize", locks.NoResize);

        if (element != "grpSpLocks")
        {
            WriteLock(xml, "noEditPoints", locks.NoEditPoints);
            WriteLock(xml, "noAdjustHandles", locks.NoAdjustHandles);
            WriteLock(xml, "noChangeArrowheads", locks.NoChangeArrowheads);
            WriteLock(xml, "noChangeShapeType", locks.NoChangeShapeType);
        }

        if (element == "spLocks")
        {
            WriteLock(xml, "noTextEdit", locks.NoTextEdit);
        }

        xml.WriteEndElement();
    }

    /// <summary>A connector's a:stCxn or a:endCxn.</summary>
    internal static void WriteConnection(XmlWriter xml, string element, ConnectionSite? site)
    {
        if (site is null)
        {
            return;
        }

        xml.WriteStartElement("a", element, Main);
        xml.WriteAttributeString(null, "id", null, DrawingValues.Text(site.ShapeId));
        xml.WriteAttributeString(null, "idx", null, DrawingValues.Text(site.Site));
        xml.WriteEndElement();
    }

    /// <summary>The extension that marks a drawing decorative, inside its cNvPr.</summary>
    internal static void WriteDecorative(XmlWriter xml)
    {
        xml.WriteStartElement("a", "extLst", Main);
        xml.WriteStartElement("a", "ext", Main);
        xml.WriteAttributeString(null, "uri", null, DecorativeExtensionUri);
        xml.WriteStartElement("adec", "decorative", DecorativeNamespace);
        xml.WriteAttributeString(null, "val", null, "1");
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    /// <summary>What an a:ln holds beside its fill: the dash, the join and the arrowheads.</summary>
    internal static void WriteLineDetails(XmlWriter xml, Outline outline)
    {
        if (outline.CustomDash.Count > 0)
        {
            xml.WriteStartElement("a", "custDash", Main);

            foreach (DashStop stop in outline.CustomDash)
            {
                xml.WriteStartElement("a", "ds", Main);
                xml.WriteAttributeString(null, "d", null, DrawingValues.Text(DrawingValues.Percent(stop.Dash)));
                xml.WriteAttributeString(null, "sp", null, DrawingValues.Text(DrawingValues.Percent(stop.Space)));
                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }
        else if (outline.Dash != LineDash.Solid)
        {
            xml.WriteStartElement("a", "prstDash", Main);
            xml.WriteAttributeString(null, "val", null, DrawingNames.Dash(outline.Dash));
            xml.WriteEndElement();
        }

        switch (outline.Join)
        {
            case LineJoin.Round:
                xml.WriteElementString("a", "round", Main, string.Empty);

                break;

            case LineJoin.Bevel:
                xml.WriteElementString("a", "bevel", Main, string.Empty);

                break;

            case LineJoin.Miter:
                xml.WriteStartElement("a", "miter", Main);

                if (outline.MiterLimit is double limit)
                {
                    xml.WriteAttributeString(null, "lim", null, DrawingValues.Text(DrawingValues.Percent(limit)));
                }

                xml.WriteEndElement();

                break;
        }

        WriteLineEnd(xml, "headEnd", outline.Head);
        WriteLineEnd(xml, "tailEnd", outline.Tail);
    }

    /// <summary>The attributes of an a:ln beside its width.</summary>
    internal static void WriteLineAttributes(XmlWriter xml, Outline outline)
    {
        if (outline.Cap is LineCap cap)
        {
            xml.WriteAttributeString(null, "cap", null, ShapeNames.Cap(cap));
        }

        if (outline.Compound != LineCompound.SingleLine)
        {
            xml.WriteAttributeString(null, "cmpd", null, ShapeNames.Compound(outline.Compound));
        }

        if (outline.Alignment == LineAlignment.Inset)
        {
            xml.WriteAttributeString(null, "algn", null, "in");
        }
    }

    /// <summary>The path or line of a gradient and its tile rectangle, after its stops.</summary>
    internal static void WriteGradientShape(XmlWriter xml, Fill fill)
    {
        if (fill.GradientShape == GradientShape.Linear)
        {
            xml.WriteStartElement("a", "lin", Main);
            xml.WriteAttributeString(null, "ang", null, DrawingValues.Text(DrawingValues.Angle(fill.Angle)));
            xml.WriteAttributeString(null, "scaled", null, DrawingValues.Flag(fill.Scaled));
            xml.WriteEndElement();
        }
        else
        {
            xml.WriteStartElement("a", "path", Main);
            xml.WriteAttributeString(null, "path", null, ShapeNames.GradientPath(fill.GradientShape));

            if (fill.Focus is FillRectangle focus)
            {
                WriteRelativeRectangle(xml, "fillToRect", focus);
            }

            xml.WriteEndElement();
        }

        if (fill.TileRectangle is FillRectangle tile)
        {
            WriteRelativeRectangle(xml, "tileRect", tile);
        }
    }

    /// <summary>An a:tile for a tiled picture fill.</summary>
    internal static void WriteTile(XmlWriter xml, PictureTiling tiling)
    {
        xml.WriteStartElement("a", "tile", Main);
        xml.WriteAttributeString(null, "tx", null, DrawingValues.Text(tiling.OffsetX.Emu));
        xml.WriteAttributeString(null, "ty", null, DrawingValues.Text(tiling.OffsetY.Emu));
        xml.WriteAttributeString(null, "sx", null, DrawingValues.Text(DrawingValues.Percent(tiling.ScaleX)));
        xml.WriteAttributeString(null, "sy", null, DrawingValues.Text(DrawingValues.Percent(tiling.ScaleY)));
        xml.WriteAttributeString(null, "flip", null, ShapeNames.Flip(tiling.Flip));
        xml.WriteAttributeString(null, "algn", null, ShapeNames.Alignment(tiling.Alignment));
        xml.WriteEndElement();
    }

    private static void WriteRelativeRectangle(XmlWriter xml, string element, FillRectangle rectangle)
    {
        xml.WriteStartElement("a", element, Main);
        xml.WriteAttributeString(null, "l", null, DrawingValues.Text(DrawingValues.Percent(rectangle.Left)));
        xml.WriteAttributeString(null, "t", null, DrawingValues.Text(DrawingValues.Percent(rectangle.Top)));
        xml.WriteAttributeString(null, "r", null, DrawingValues.Text(DrawingValues.Percent(rectangle.Right)));
        xml.WriteAttributeString(null, "b", null, DrawingValues.Text(DrawingValues.Percent(rectangle.Bottom)));
        xml.WriteEndElement();
    }

    private static void WriteLineEnd(XmlWriter xml, string element, LineEnd? end)
    {
        if (end is null)
        {
            return;
        }

        xml.WriteStartElement("a", element, Main);
        xml.WriteAttributeString(null, "type", null, ShapeNames.EndType(end.Type));

        if (end.Width != LineEndSize.Medium)
        {
            xml.WriteAttributeString(null, "w", null, ShapeNames.EndSize(end.Width));
        }

        if (end.Length != LineEndSize.Medium)
        {
            xml.WriteAttributeString(null, "len", null, ShapeNames.EndSize(end.Length));
        }

        xml.WriteEndElement();
    }

    private static void WriteReference(XmlWriter xml, string element, StyleReference reference)
    {
        xml.WriteStartElement("a", element, Main);
        xml.WriteAttributeString(null, "idx", null, DrawingValues.Text(reference.Index));
        xml.WriteStartElement("a", "schemeClr", Main);
        xml.WriteAttributeString(null, "val", null, reference.Color);

        if (reference.Shade is double shade)
        {
            xml.WriteStartElement("a", "shade", Main);
            xml.WriteAttributeString(null, "val", null, DrawingValues.Text(DrawingValues.Percent(shade)));
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    private static void WriteLock(XmlWriter xml, string name, bool value)
    {
        if (value == true)
        {
            xml.WriteAttributeString(null, name, null, "1");
        }
    }

    private static void WriteBevel(XmlWriter xml, string element, Bevel? bevel)
    {
        if (bevel is null)
        {
            return;
        }

        xml.WriteStartElement("a", element, Main);
        xml.WriteAttributeString(null, "w", null, DrawingValues.Text(DrawingValues.Emu(bevel.Width)));
        xml.WriteAttributeString(null, "h", null, DrawingValues.Text(DrawingValues.Emu(bevel.Height)));

        if (bevel.Type != BevelType.Circle)
        {
            xml.WriteAttributeString(null, "prst", null, ShapeNames.Bevel(bevel.Type));
        }

        xml.WriteEndElement();
    }

    private static void WriteRotation3D(XmlWriter xml, Rotation3D? rotation)
    {
        if (rotation is null)
        {
            return;
        }

        xml.WriteStartElement("a", "rot", Main);
        xml.WriteAttributeString(null, "lat", null, DrawingValues.Text(DrawingValues.Angle(rotation.Latitude)));
        xml.WriteAttributeString(null, "lon", null, DrawingValues.Text(DrawingValues.Angle(rotation.Longitude)));
        xml.WriteAttributeString(null, "rev", null, DrawingValues.Text(DrawingValues.Angle(rotation.Revolution)));
        xml.WriteEndElement();
    }

    private static void WriteVector(XmlWriter xml, string element, string across, string down, string toward, Vector3D vector)
    {
        xml.WriteStartElement("a", element, Main);
        xml.WriteAttributeString(null, across, null, DrawingValues.Text(vector.Across));
        xml.WriteAttributeString(null, down, null, DrawingValues.Text(vector.Down));
        xml.WriteAttributeString(null, toward, null, DrawingValues.Text(vector.Toward));
        xml.WriteEndElement();
    }

    private static void WriteScale(XmlWriter xml, string name, double scale)
    {
        if (scale != 1)
        {
            xml.WriteAttributeString(null, name, null, DrawingValues.Text(DrawingValues.Percent(scale)));
        }
    }

    private static void WriteSkew(XmlWriter xml, string name, double degrees)
    {
        if (degrees != 0)
        {
            xml.WriteAttributeString(null, name, null, DrawingValues.Text((long)Math.Round(degrees * 60_000)));
        }
    }
}
