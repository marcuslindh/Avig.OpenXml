using System.Xml;

namespace Avig.OpenXml.Drawing;

/// <summary>
/// Writes DrawingML: shapes, pictures, groups and chart frames as the elements of a shape tree
/// (in the namespace of <see cref="DrawingContext.Tree"/>), and the fills, lines, effects and text
/// they are made of. What goes around them (a spreadsheet anchor, a slide's spTree) and the
/// relationships in <see cref="DrawingContext.Targets"/> are written by the package that owns the
/// part. Every element is written in the order the schema gives, which Office checks.
/// </summary>
public static partial class DrawingWriter
{
    private const string Main = OpcNamespaces.DrawingMain;

    /// <summary>The extension Office 2016 put on a:blip for the SVG behind a raster picture.</summary>
    private const string SvgExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

    /// <summary>A shape (sp), or for <see cref="ShapeKind.Line"/> a connector (cxnSp).</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="shape">The shape.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size (a:off, a:ext).</param>
    public static void WriteShape(XmlWriter xml, Shape shape, DrawingContext context, DrawingBounds frame)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(shape);
        ArgumentNullException.ThrowIfNull(context);
        ShapeTreeNamespace tree = context.Tree;
        int shapeId = context.NextShapeId();
        string name = shape.Name ?? PresetGeometry.DefaultName(shape) + " " + DrawingValues.Text(shapeId - 1);
        bool line = shape.Kind == ShapeKind.Line;
        bool textBox = shape.Kind == ShapeKind.TextBox;

        if (line == true)
        {
            xml.WriteStartElement(tree.Prefix, "cxnSp", tree.Uri);

            if (tree.MacroAttributes == true)
            {
                xml.WriteAttributeString(null, "macro", null, string.Empty);
            }

            xml.WriteStartElement(tree.Prefix, "nvCxnSpPr", tree.Uri);
            WriteNonVisualProperties(xml, tree, shapeId, name, shape.Description, shape.Link, context);
            xml.WriteElementString(tree.Prefix, "cNvCxnSpPr", tree.Uri, string.Empty);
            WriteApplicationProperties(xml, tree);
            xml.WriteEndElement();
        }
        else
        {
            xml.WriteStartElement(tree.Prefix, "sp", tree.Uri);

            if (tree.MacroAttributes == true)
            {
                xml.WriteAttributeString(null, "macro", null, string.Empty);
                xml.WriteAttributeString(null, "textlink", null, string.Empty);
            }

            xml.WriteStartElement(tree.Prefix, "nvSpPr", tree.Uri);
            WriteNonVisualProperties(xml, tree, shapeId, name, shape.Description, shape.Link, context);
            xml.WriteStartElement(tree.Prefix, "cNvSpPr", tree.Uri);

            if (textBox == true)
            {
                xml.WriteAttributeString(null, "txBox", null, "1");
            }

            xml.WriteEndElement();
            WriteApplicationProperties(xml, tree);
            xml.WriteEndElement();
        }

        xml.WriteStartElement(tree.Prefix, "spPr", tree.Uri);
        xml.WriteStartElement("a", "xfrm", Main);

        if (shape.Rotation != 0)
        {
            xml.WriteAttributeString(null, "rot", null, DrawingValues.Text(DrawingValues.Angle(shape.Rotation)));
        }

        if (shape.FlipHorizontal == true)
        {
            xml.WriteAttributeString(null, "flipH", null, "1");
        }

        if (shape.FlipVertical == true)
        {
            xml.WriteAttributeString(null, "flipV", null, "1");
        }

        WriteOffsetAndExtent(xml, frame);
        xml.WriteEndElement();
        List<ShapeAdjustment> adjustments = [.. shape.Adjustments];

        if (shape.Kind == ShapeKind.RoundedRectangle && shape.CornerRadius is double radius)
        {
            adjustments.RemoveAll(adjustment => adjustment.Name == "adj");
            adjustments.Insert(0, new ShapeAdjustment("adj", DrawingValues.Percent(radius)));
        }

        WritePresetGeometry(xml, shape.Preset, adjustments);
        Fill? fill = shape.Fill;
        Outline? outline = shape.Outline;

        // A text box is text alone unless it is given a fill or a line; a line has no inside.
        if (textBox == true)
        {
            fill ??= Fill.None;
            outline ??= Outline.None;
        }

        if (fill is not null && line == false)
        {
            WriteFill(xml, fill, context);
        }

        if (outline is not null)
        {
            WriteOutline(xml, outline);
        }

        if (shape.Shadow is not null || shape.InnerShadow is not null || shape.SoftEdges is not null)
        {
            WriteEffects(xml, shape.Shadow, shape.InnerShadow, shape.SoftEdges);
        }

        xml.WriteEndElement();
        WriteStyle(xml, tree, shape.Kind);

        if (line == false)
        {
            WriteTextBody(xml, tree, shape.Text);
        }

        xml.WriteEndElement();
    }

    /// <summary>A picture (pic), with the SVG as an extension of the raster blip.</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="picture">The picture.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size.</param>
    public static void WritePicture(XmlWriter xml, Picture picture, DrawingContext context, DrawingBounds frame)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(picture);
        ArgumentNullException.ThrowIfNull(context);
        ShapeTreeNamespace tree = context.Tree;
        int shapeId = context.NextShapeId();
        xml.WriteStartElement(tree.Prefix, "pic", tree.Uri);
        xml.WriteStartElement(tree.Prefix, "nvPicPr", tree.Uri);
        WriteNonVisualProperties(xml, tree, shapeId, "Picture " + DrawingValues.Text(shapeId - 1), picture.Description, picture.Link, context);
        xml.WriteStartElement(tree.Prefix, "cNvPicPr", tree.Uri);
        xml.WriteStartElement("a", "picLocks", Main);
        xml.WriteAttributeString(null, "noChangeAspect", null, "1");
        xml.WriteEndElement();
        xml.WriteEndElement();
        WriteApplicationProperties(xml, tree);
        xml.WriteEndElement();
        xml.WriteStartElement(tree.Prefix, "blipFill", tree.Uri);
        WriteBlip(xml, picture, context, picture.Transparency);
        xml.WriteStartElement("a", "stretch", Main);
        xml.WriteElementString("a", "fillRect", Main, string.Empty);
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteStartElement(tree.Prefix, "spPr", tree.Uri);
        xml.WriteStartElement("a", "xfrm", Main);
        WriteOffsetAndExtent(xml, frame);
        xml.WriteEndElement();
        WritePresetGeometry(xml, "rect", []);

        if (picture.Shadow is not null)
        {
            WriteEffects(xml, picture.Shadow, null, null);
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    /// <summary>A group (grpSp), whose items' own coordinates are mapped onto its frame.</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="group">The group.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size of the whole group.</param>
    public static void WriteGroup(XmlWriter xml, ShapeGroup group, DrawingContext context, DrawingBounds frame)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(group);
        ArgumentNullException.ThrowIfNull(context);
        ShapeTreeNamespace tree = context.Tree;
        int shapeId = context.NextShapeId();
        xml.WriteStartElement(tree.Prefix, "grpSp", tree.Uri);
        xml.WriteStartElement(tree.Prefix, "nvGrpSpPr", tree.Uri);
        WriteNonVisualProperties(xml, tree, shapeId, group.Name ?? "Group " + DrawingValues.Text(shapeId - 1), group.Description, group.Link, context);
        xml.WriteElementString(tree.Prefix, "cNvGrpSpPr", tree.Uri, string.Empty);
        WriteApplicationProperties(xml, tree);
        xml.WriteEndElement();

        // The items' own coordinates (chOff, chExt) are mapped onto the group's place, so the
        // box around them fills it.
        long left = group.Items.Min(item => item.Bounds.Left.Emu);
        long top = group.Items.Min(item => item.Bounds.Top.Emu);
        long right = group.Items.Max(item => item.Bounds.Left.Emu + item.Bounds.Width.Emu);
        long bottom = group.Items.Max(item => item.Bounds.Top.Emu + item.Bounds.Height.Emu);
        xml.WriteStartElement(tree.Prefix, "grpSpPr", tree.Uri);
        xml.WriteStartElement("a", "xfrm", Main);
        WriteOffsetAndExtent(xml, frame);
        xml.WriteStartElement("a", "chOff", Main);
        xml.WriteAttributeString(null, "x", null, DrawingValues.Text(left));
        xml.WriteAttributeString(null, "y", null, DrawingValues.Text(top));
        xml.WriteEndElement();
        xml.WriteStartElement("a", "chExt", Main);
        xml.WriteAttributeString(null, "cx", null, DrawingValues.Text(Math.Max(right - left, 1)));
        xml.WriteAttributeString(null, "cy", null, DrawingValues.Text(Math.Max(bottom - top, 1)));
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();

        foreach (GroupItem item in group.Items)
        {
            if (item.Shape is not null)
            {
                WriteShape(xml, item.Shape, context, item.Bounds);
            }
            else if (item.Picture is not null)
            {
                WritePicture(xml, item.Picture, context, item.Bounds);
            }
            else
            {
                WriteGroup(xml, item.Group!, context, item.Bounds);
            }
        }

        xml.WriteEndElement();
    }

    /// <summary>A graphic frame that shows a chart part; the chart itself is written with <see cref="ChartWriter"/>.</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="chart">What the owner needs to write the chart part; it is handed back in <see cref="DrawingContext.Targets"/>.</param>
    /// <param name="description">The alternative text, or null.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size.</param>
    public static void WriteChartFrame(XmlWriter xml, object chart, string? description, DrawingContext context, DrawingBounds frame)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(chart);
        ArgumentNullException.ThrowIfNull(context);
        ShapeTreeNamespace tree = context.Tree;
        int shapeId = context.NextShapeId();
        string relationshipId = context.Chart(chart);
        xml.WriteStartElement(tree.Prefix, "graphicFrame", tree.Uri);

        if (tree.MacroAttributes == true)
        {
            xml.WriteAttributeString(null, "macro", null, string.Empty);
        }

        xml.WriteStartElement(tree.Prefix, "nvGraphicFramePr", tree.Uri);
        WriteNonVisualProperties(xml, tree, shapeId, "Chart " + DrawingValues.Text(shapeId - 1), description, null, context);
        xml.WriteElementString(tree.Prefix, "cNvGraphicFramePr", tree.Uri, string.Empty);
        WriteApplicationProperties(xml, tree);
        xml.WriteEndElement();
        xml.WriteStartElement(tree.Prefix, "xfrm", tree.Uri);
        WriteOffsetAndExtent(xml, frame);
        xml.WriteEndElement();
        xml.WriteStartElement("a", "graphic", Main);
        xml.WriteStartElement("a", "graphicData", Main);
        xml.WriteAttributeString(null, "uri", null, OpcNamespaces.DrawingChart);
        xml.WriteStartElement("c", "chart", OpcNamespaces.DrawingChart);
        xml.WriteAttributeString("r", "id", OpcNamespaces.Relationships, relationshipId);
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    /// <summary>An a:solidFill, a:gradFill, a:pattFill, a:blipFill or a:noFill.</summary>
    /// <param name="xml">The writer.</param>
    /// <param name="fill">The fill.</param>
    /// <param name="context">The part's relationships, which a picture fill needs; null where there are none, as in a chart.</param>
    public static void WriteFill(XmlWriter xml, Fill fill, DrawingContext? context = null)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(fill);

        switch (fill.Kind)
        {
            case FillKind.None:
                xml.WriteElementString("a", "noFill", Main, string.Empty);

                break;

            case FillKind.Solid:
                xml.WriteStartElement("a", "solidFill", Main);
                WriteColor(xml, fill.Color!, fill.Transparency);
                xml.WriteEndElement();

                break;

            case FillKind.Pattern:
                xml.WriteStartElement("a", "pattFill", Main);
                xml.WriteAttributeString(null, "prst", null, DrawingNames.Pattern(fill.PatternType));
                xml.WriteStartElement("a", "fgClr", Main);
                WriteColor(xml, fill.Color!, 0);
                xml.WriteEndElement();
                xml.WriteStartElement("a", "bgClr", Main);
                WriteColor(xml, fill.BackgroundColor!, 0);
                xml.WriteEndElement();
                xml.WriteEndElement();

                break;

            case FillKind.Picture:
                if (context is null)
                {
                    throw new InvalidOperationException("A picture fill needs a DrawingContext for the image's relationship.");
                }

                xml.WriteStartElement("a", "blipFill", Main);
                xml.WriteAttributeString(null, "rotWithShape", null, "1");
                WriteBlip(xml, fill.Image!, context, 0);
                xml.WriteStartElement("a", "stretch", Main);
                xml.WriteElementString("a", "fillRect", Main, string.Empty);
                xml.WriteEndElement();
                xml.WriteEndElement();

                break;

            default:
                xml.WriteStartElement("a", "gradFill", Main);
                xml.WriteAttributeString(null, "rotWithShape", null, "1");
                xml.WriteStartElement("a", "gsLst", Main);

                foreach (GradientStop stop in fill.Stops.OrderBy(stop => stop.Position))
                {
                    xml.WriteStartElement("a", "gs", Main);
                    xml.WriteAttributeString(null, "pos", null, DrawingValues.Text(DrawingValues.Percent(stop.Position)));
                    WriteColor(xml, stop.Color, stop.Transparency);
                    xml.WriteEndElement();
                }

                xml.WriteEndElement();
                xml.WriteStartElement("a", "lin", Main);
                xml.WriteAttributeString(null, "ang", null, DrawingValues.Text(DrawingValues.Angle(fill.Angle)));
                xml.WriteAttributeString(null, "scaled", null, "0");
                xml.WriteEndElement();
                xml.WriteEndElement();

                break;
        }
    }

    /// <summary>An a:ln.</summary>
    /// <param name="xml">The writer.</param>
    /// <param name="outline">The line.</param>
    public static void WriteOutline(XmlWriter xml, Outline outline)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(outline);
        xml.WriteStartElement("a", "ln", Main);

        if (outline.Color is null)
        {
            xml.WriteElementString("a", "noFill", Main, string.Empty);
            xml.WriteEndElement();

            return;
        }

        xml.WriteAttributeString(null, "w", null, DrawingValues.Text(DrawingValues.Emu(outline.Width)));
        xml.WriteStartElement("a", "solidFill", Main);
        WriteColor(xml, outline.Color, outline.Transparency);
        xml.WriteEndElement();

        if (outline.Dash != LineDash.Solid)
        {
            xml.WriteStartElement("a", "prstDash", Main);
            xml.WriteAttributeString(null, "val", null, DrawingNames.Dash(outline.Dash));
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    /// <summary>An a:effectLst with the shadows and soft edges given, in the schema's order.</summary>
    /// <param name="xml">The writer.</param>
    /// <param name="outer">The shadow behind, or null.</param>
    /// <param name="inner">The shadow inside, or null.</param>
    /// <param name="softEdges">Soft edges in points, or null.</param>
    public static void WriteEffects(XmlWriter xml, Shadow? outer, Shadow? inner, double? softEdges)
    {
        ArgumentNullException.ThrowIfNull(xml);
        xml.WriteStartElement("a", "effectLst", Main);

        if (inner is not null)
        {
            xml.WriteStartElement("a", "innerShdw", Main);
            WriteShadowAttributes(xml, inner);
            WriteColor(xml, inner.Color, inner.Transparency);
            xml.WriteEndElement();
        }

        if (outer is not null)
        {
            xml.WriteStartElement("a", "outerShdw", Main);
            WriteShadowAttributes(xml, outer);
            xml.WriteAttributeString(null, "algn", null, "tl");
            xml.WriteAttributeString(null, "rotWithShape", null, "0");
            WriteColor(xml, outer.Color, outer.Transparency);
            xml.WriteEndElement();
        }

        if (softEdges is double points)
        {
            xml.WriteStartElement("a", "softEdge", Main);
            xml.WriteAttributeString(null, "rad", null, DrawingValues.Text(DrawingValues.Emu(points)));
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    /// <summary>
    /// An a:rPr, a:defRPr or a:endParaRPr. Bold and italic are written either way, so that a
    /// font given without them overrides a bold default such as a title's.
    /// </summary>
    /// <param name="xml">The writer.</param>
    /// <param name="element">rPr, defRPr or endParaRPr.</param>
    /// <param name="font">The font, or null for the element alone.</param>
    public static void WriteRunProperties(XmlWriter xml, string element, TextFont? font)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(element);
        xml.WriteStartElement("a", element, Main);

        if (element != "defRPr")
        {
            xml.WriteAttributeString(null, "lang", null, "en-US");
        }

        if (font is null)
        {
            xml.WriteEndElement();

            return;
        }

        if (font.Size is double size)
        {
            xml.WriteAttributeString(null, "sz", null, DrawingValues.Text((long)Math.Round(size * 100)));
        }

        xml.WriteAttributeString(null, "b", null, DrawingValues.Flag(font.Bold));
        xml.WriteAttributeString(null, "i", null, DrawingValues.Flag(font.Italic));

        if (font.Color is not null)
        {
            xml.WriteStartElement("a", "solidFill", Main);
            WriteColor(xml, font.Color, 0);
            xml.WriteEndElement();
        }

        if (font.Name is not null)
        {
            xml.WriteStartElement("a", "latin", Main);
            xml.WriteAttributeString(null, "typeface", null, font.Name);
            xml.WriteEndElement();
            xml.WriteStartElement("a", "cs", Main);
            xml.WriteAttributeString(null, "typeface", null, font.Name);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    /// <summary>An a:off and an a:ext, in EMU.</summary>
    /// <param name="xml">The writer, inside an xfrm.</param>
    /// <param name="frame">The place and size.</param>
    public static void WriteOffsetAndExtent(XmlWriter xml, DrawingBounds frame)
    {
        ArgumentNullException.ThrowIfNull(xml);
        xml.WriteStartElement("a", "off", Main);
        xml.WriteAttributeString(null, "x", null, DrawingValues.Text(frame.Left.Emu));
        xml.WriteAttributeString(null, "y", null, DrawingValues.Text(frame.Top.Emu));
        xml.WriteEndElement();
        xml.WriteStartElement("a", "ext", Main);
        xml.WriteAttributeString(null, "cx", null, DrawingValues.Text(frame.Width.Emu));
        xml.WriteAttributeString(null, "cy", null, DrawingValues.Text(frame.Height.Emu));
        xml.WriteEndElement();
    }

    /// <summary>The a:blip of a picture: the raster image, with the SVG as an extension older versions skip.</summary>
    private static void WriteBlip(XmlWriter xml, Picture picture, DrawingContext context, double transparency)
    {
        xml.WriteStartElement("a", "blip", Main);

        // An SVG is stored as Office 2016 stores it: the raster fallback is the blip, and the
        // SVG an extension of it.
        Picture raster = picture.Fallback ?? picture;
        xml.WriteAttributeString("r", "embed", OpcNamespaces.Relationships, context.Image(raster));

        if (transparency > 0)
        {
            xml.WriteStartElement("a", "alphaModFix", Main);
            xml.WriteAttributeString(null, "amt", null, DrawingValues.Text(DrawingValues.Percent(1 - transparency)));
            xml.WriteEndElement();
        }

        if (picture.Fallback is not null)
        {
            xml.WriteStartElement("a", "extLst", Main);
            xml.WriteStartElement("a", "ext", Main);
            xml.WriteAttributeString(null, "uri", null, SvgExtensionUri);
            xml.WriteStartElement("asvg", "svgBlip", OpcNamespaces.DrawingSvg);
            xml.WriteAttributeString("r", "embed", OpcNamespaces.Relationships, context.Image(picture));
            xml.WriteEndElement();
            xml.WriteEndElement();
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    /// <summary>The theme references Office gives a new shape, so that what is left unset looks as in Office.</summary>
    private static void WriteStyle(XmlWriter xml, ShapeTreeNamespace tree, ShapeKind? kind)
    {
        (string Line, string Fill, string Effect, string Font, bool Shade) style = kind switch
        {
            ShapeKind.Line => ("1", "0", "0", "tx1", false),
            ShapeKind.TextBox => ("0", "0", "0", "dk1", false),
            _ => ("2", "1", "0", "lt1", true),
        };

        xml.WriteStartElement(tree.Prefix, "style", tree.Uri);
        WriteStyleReference(xml, "lnRef", style.Line, "accent1", style.Shade);
        WriteStyleReference(xml, "fillRef", style.Fill, "accent1", false);
        WriteStyleReference(xml, "effectRef", style.Effect, "accent1", false);
        xml.WriteStartElement("a", "fontRef", Main);
        xml.WriteAttributeString(null, "idx", null, "minor");
        xml.WriteStartElement("a", "schemeClr", Main);
        xml.WriteAttributeString(null, "val", null, style.Font);
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    private static void WriteStyleReference(XmlWriter xml, string element, string index, string color, bool shade)
    {
        xml.WriteStartElement("a", element, Main);
        xml.WriteAttributeString(null, "idx", null, index);
        xml.WriteStartElement("a", "schemeClr", Main);
        xml.WriteAttributeString(null, "val", null, color);

        if (shade == true)
        {
            xml.WriteStartElement("a", "shade", Main);
            xml.WriteAttributeString(null, "val", null, "15000");
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    /// <summary>The shape's txBody: one a:p per line of the runs, each run with its own look.</summary>
    private static void WriteTextBody(XmlWriter xml, ShapeTreeNamespace tree, ShapeText? text)
    {
        xml.WriteStartElement(tree.Prefix, "txBody", tree.Uri);
        xml.WriteStartElement("a", "bodyPr", Main);
        xml.WriteAttributeString(null, "vertOverflow", null, "clip");
        xml.WriteAttributeString(null, "horzOverflow", null, "clip");

        if (text is not null)
        {
            if (text.Wrap == true)
            {
                xml.WriteAttributeString(null, "wrap", null, "square");
            }
            else
            {
                xml.WriteAttributeString(null, "wrap", null, "none");
            }

            WriteInset(xml, "lIns", text.MarginLeft);
            WriteInset(xml, "tIns", text.MarginTop);
            WriteInset(xml, "rIns", text.MarginRight);
            WriteInset(xml, "bIns", text.MarginBottom);
        }

        xml.WriteAttributeString(null, "rtlCol", null, "0");
        xml.WriteAttributeString(null, "anchor", null, DrawingNames.Anchor(text?.Vertical ?? TextAnchor.Top));
        xml.WriteEndElement();
        xml.WriteElementString("a", "lstStyle", Main, string.Empty);
        string alignment = DrawingNames.Alignment(text?.Horizontal ?? TextAlignment.Left);

        if (text is null)
        {
            xml.WriteStartElement("a", "p", Main);
            WriteParagraphProperties(xml, alignment);
            WriteRunProperties(xml, "endParaRPr", new TextFont { Size = 11 });
            xml.WriteEndElement();
            xml.WriteEndElement();

            return;
        }

        foreach (List<TextRun> paragraph in Paragraphs(text.Runs))
        {
            xml.WriteStartElement("a", "p", Main);
            WriteParagraphProperties(xml, alignment);
            TextFont? last = null;

            foreach (TextRun run in paragraph)
            {
                last = run.Font;

                if (run.Text.Length == 0)
                {
                    continue;
                }

                xml.WriteStartElement("a", "r", Main);
                WriteRunProperties(xml, "rPr", run.Font);
                xml.WriteElementString("a", "t", Main, run.Text);
                xml.WriteEndElement();
            }

            // Gives an empty line its height, and the cursor its look at the end of a line.
            WriteRunProperties(xml, "endParaRPr", last);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    /// <summary>The runs split at each line break (\n, \r\n or \r) into paragraphs; a run split so keeps its font on both sides.</summary>
    private static List<List<TextRun>> Paragraphs(IReadOnlyList<TextRun> runs)
    {
        List<List<TextRun>> paragraphs = [[]];

        foreach (TextRun run in runs)
        {
            string[] lines = run.Text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');

            for (int index = 0; index < lines.Length; index++)
            {
                if (index > 0)
                {
                    paragraphs.Add([]);
                }

                paragraphs[^1].Add(new TextRun(lines[index], run.Font));
            }
        }

        return paragraphs;
    }

    private static void WriteParagraphProperties(XmlWriter xml, string alignment)
    {
        xml.WriteStartElement("a", "pPr", Main);
        xml.WriteAttributeString(null, "algn", null, alignment);
        xml.WriteEndElement();
    }

    private static void WriteInset(XmlWriter xml, string name, double? points)
    {
        if (points is double value)
        {
            xml.WriteAttributeString(null, name, null, DrawingValues.Text(DrawingValues.Emu(value)));
        }
    }

    private static void WritePresetGeometry(XmlWriter xml, string preset, IReadOnlyList<ShapeAdjustment> adjustments)
    {
        xml.WriteStartElement("a", "prstGeom", Main);
        xml.WriteAttributeString(null, "prst", null, preset);
        xml.WriteStartElement("a", "avLst", Main);

        foreach (ShapeAdjustment adjustment in adjustments)
        {
            xml.WriteStartElement("a", "gd", Main);
            xml.WriteAttributeString(null, "name", null, adjustment.Name);
            xml.WriteAttributeString(null, "fmla", null, "val " + DrawingValues.Text(adjustment.Value));
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    /// <summary>A cNvPr: id, name, alternative text and the click's link.</summary>
    private static void WriteNonVisualProperties(XmlWriter xml, ShapeTreeNamespace tree, int shapeId, string name, string? description, Hyperlink? link, DrawingContext context)
    {
        xml.WriteStartElement(tree.Prefix, "cNvPr", tree.Uri);
        xml.WriteAttributeString(null, "id", null, DrawingValues.Text(shapeId));
        xml.WriteAttributeString(null, "name", null, name);

        if (description is not null)
        {
            xml.WriteAttributeString(null, "descr", null, description);
        }

        if (link is not null)
        {
            xml.WriteStartElement("a", "hlinkClick", Main);
            xml.WriteAttributeString("r", "id", OpcNamespaces.Relationships, context.Link(link));

            if (link.Tooltip is not null)
            {
                xml.WriteAttributeString(null, "tooltip", null, link.Tooltip);
            }

            if (link.Action is not null)
            {
                xml.WriteAttributeString(null, "action", null, link.Action);
            }

            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private static void WriteApplicationProperties(XmlWriter xml, ShapeTreeNamespace tree)
    {
        if (tree.ApplicationProperties == true)
        {
            xml.WriteElementString(tree.Prefix, "nvPr", tree.Uri, string.Empty);
        }
    }

    private static void WriteShadowAttributes(XmlWriter xml, Shadow shadow)
    {
        xml.WriteAttributeString(null, "blurRad", null, DrawingValues.Text(DrawingValues.Emu(shadow.Blur)));
        xml.WriteAttributeString(null, "dist", null, DrawingValues.Text(DrawingValues.Emu(shadow.Distance)));
        xml.WriteAttributeString(null, "dir", null, DrawingValues.Text(DrawingValues.Angle(shadow.Angle)));
    }

    /// <summary>An a:srgbClr, with an a:alpha when the colour is see-through.</summary>
    private static void WriteColor(XmlWriter xml, string color, double transparency)
    {
        xml.WriteStartElement("a", "srgbClr", Main);
        xml.WriteAttributeString(null, "val", null, color[1..]);

        if (transparency > 0)
        {
            xml.WriteStartElement("a", "alpha", Main);
            xml.WriteAttributeString(null, "val", null, DrawingValues.Text(DrawingValues.Percent(1 - transparency)));
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }
}
