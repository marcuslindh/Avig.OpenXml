using System.Xml;

namespace Avig.OpenXml.Drawing;

/// <summary>
/// Writes DrawingML: shapes, pictures, groups and chart frames as the elements of a shape tree
/// (in the namespace of <see cref="DrawingContext.Tree"/>), and the fills, lines, effects and text
/// they are made of. What goes around them (a spreadsheet anchor, a slide's spTree) and the
/// relationships in <see cref="DrawingContext.Targets"/> are written by the package that owns the
/// part. Every element is written in the order the schema gives, which Office checks.
/// </summary>
public static class DrawingWriter
{
    private const string Main = OpcNamespaces.DrawingMain;

    /// <summary>The extension Office 2016 put on a:blip for the SVG behind a raster picture.</summary>
    private const string SvgExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

    /// <summary>A shape (sp), or for <see cref="ShapeKind.Line"/> a connector (cxnSp).</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="shape">The shape.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size (a:off, a:ext).</param>
    /// <returns>A task.</returns>
    public static async Task WriteShapeAsync(XmlWriter xml, Shape shape, DrawingContext context, DrawingBounds frame)
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
            await xml.WriteStartElementAsync(tree.Prefix, "cxnSp", tree.Uri);

            if (tree.MacroAttributes == true)
            {
                await xml.WriteAttributeStringAsync(null, "macro", null, string.Empty);
            }

            await xml.WriteStartElementAsync(tree.Prefix, "nvCxnSpPr", tree.Uri);
            await WriteNonVisualPropertiesAsync(xml, tree, shapeId, name, shape.Description, shape.Link, context);
            await xml.WriteElementStringAsync(tree.Prefix, "cNvCxnSpPr", tree.Uri, string.Empty);
            await WriteApplicationPropertiesAsync(xml, tree);
            await xml.WriteEndElementAsync();
        }
        else
        {
            await xml.WriteStartElementAsync(tree.Prefix, "sp", tree.Uri);

            if (tree.MacroAttributes == true)
            {
                await xml.WriteAttributeStringAsync(null, "macro", null, string.Empty);
                await xml.WriteAttributeStringAsync(null, "textlink", null, string.Empty);
            }

            await xml.WriteStartElementAsync(tree.Prefix, "nvSpPr", tree.Uri);
            await WriteNonVisualPropertiesAsync(xml, tree, shapeId, name, shape.Description, shape.Link, context);
            await xml.WriteStartElementAsync(tree.Prefix, "cNvSpPr", tree.Uri);

            if (textBox == true)
            {
                await xml.WriteAttributeStringAsync(null, "txBox", null, "1");
            }

            await xml.WriteEndElementAsync();
            await WriteApplicationPropertiesAsync(xml, tree);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteStartElementAsync(tree.Prefix, "spPr", tree.Uri);
        await xml.WriteStartElementAsync("a", "xfrm", Main);

        if (shape.Rotation != 0)
        {
            await xml.WriteAttributeStringAsync(null, "rot", null, DrawingValues.Text(DrawingValues.Angle(shape.Rotation)));
        }

        if (shape.FlipHorizontal == true)
        {
            await xml.WriteAttributeStringAsync(null, "flipH", null, "1");
        }

        if (shape.FlipVertical == true)
        {
            await xml.WriteAttributeStringAsync(null, "flipV", null, "1");
        }

        await WriteOffsetAndExtentAsync(xml, frame);
        await xml.WriteEndElementAsync();
        List<ShapeAdjustment> adjustments = [.. shape.Adjustments];

        if (shape.Kind == ShapeKind.RoundedRectangle && shape.CornerRadius is double radius)
        {
            adjustments.RemoveAll(adjustment => adjustment.Name == "adj");
            adjustments.Insert(0, new ShapeAdjustment("adj", DrawingValues.Percent(radius)));
        }

        await WritePresetGeometryAsync(xml, shape.Preset, adjustments);
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
            await WriteFillAsync(xml, fill, context);
        }

        if (outline is not null)
        {
            await WriteOutlineAsync(xml, outline);
        }

        if (shape.Shadow is not null || shape.InnerShadow is not null || shape.SoftEdges is not null)
        {
            await WriteEffectsAsync(xml, shape.Shadow, shape.InnerShadow, shape.SoftEdges);
        }

        await xml.WriteEndElementAsync();
        await WriteStyleAsync(xml, tree, shape.Kind);

        if (line == false)
        {
            await WriteTextBodyAsync(xml, tree, shape.Text);
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>A picture (pic), with the SVG as an extension of the raster blip.</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="picture">The picture.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size.</param>
    /// <returns>A task.</returns>
    public static async Task WritePictureAsync(XmlWriter xml, Picture picture, DrawingContext context, DrawingBounds frame)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(picture);
        ArgumentNullException.ThrowIfNull(context);
        ShapeTreeNamespace tree = context.Tree;
        int shapeId = context.NextShapeId();
        await xml.WriteStartElementAsync(tree.Prefix, "pic", tree.Uri);
        await xml.WriteStartElementAsync(tree.Prefix, "nvPicPr", tree.Uri);
        await WriteNonVisualPropertiesAsync(xml, tree, shapeId, "Picture " + DrawingValues.Text(shapeId - 1), picture.Description, picture.Link, context);
        await xml.WriteStartElementAsync(tree.Prefix, "cNvPicPr", tree.Uri);
        await xml.WriteStartElementAsync("a", "picLocks", Main);
        await xml.WriteAttributeStringAsync(null, "noChangeAspect", null, "1");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await WriteApplicationPropertiesAsync(xml, tree);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync(tree.Prefix, "blipFill", tree.Uri);
        await WriteBlipAsync(xml, picture, context, picture.Transparency);
        await xml.WriteStartElementAsync("a", "stretch", Main);
        await xml.WriteElementStringAsync("a", "fillRect", Main, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync(tree.Prefix, "spPr", tree.Uri);
        await xml.WriteStartElementAsync("a", "xfrm", Main);
        await WriteOffsetAndExtentAsync(xml, frame);
        await xml.WriteEndElementAsync();
        await WritePresetGeometryAsync(xml, "rect", []);

        if (picture.Shadow is not null)
        {
            await WriteEffectsAsync(xml, picture.Shadow, null, null);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    /// <summary>A group (grpSp), whose items' own coordinates are mapped onto its frame.</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="group">The group.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size of the whole group.</param>
    /// <returns>A task.</returns>
    public static async Task WriteGroupAsync(XmlWriter xml, ShapeGroup group, DrawingContext context, DrawingBounds frame)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(group);
        ArgumentNullException.ThrowIfNull(context);
        ShapeTreeNamespace tree = context.Tree;
        int shapeId = context.NextShapeId();
        await xml.WriteStartElementAsync(tree.Prefix, "grpSp", tree.Uri);
        await xml.WriteStartElementAsync(tree.Prefix, "nvGrpSpPr", tree.Uri);
        await WriteNonVisualPropertiesAsync(xml, tree, shapeId, group.Name ?? "Group " + DrawingValues.Text(shapeId - 1), group.Description, group.Link, context);
        await xml.WriteElementStringAsync(tree.Prefix, "cNvGrpSpPr", tree.Uri, string.Empty);
        await WriteApplicationPropertiesAsync(xml, tree);
        await xml.WriteEndElementAsync();

        // The items' own coordinates (chOff, chExt) are mapped onto the group's place, so the
        // box around them fills it.
        long left = group.Items.Min(item => item.Bounds.Left.Emu);
        long top = group.Items.Min(item => item.Bounds.Top.Emu);
        long right = group.Items.Max(item => item.Bounds.Left.Emu + item.Bounds.Width.Emu);
        long bottom = group.Items.Max(item => item.Bounds.Top.Emu + item.Bounds.Height.Emu);
        await xml.WriteStartElementAsync(tree.Prefix, "grpSpPr", tree.Uri);
        await xml.WriteStartElementAsync("a", "xfrm", Main);
        await WriteOffsetAndExtentAsync(xml, frame);
        await xml.WriteStartElementAsync("a", "chOff", Main);
        await xml.WriteAttributeStringAsync(null, "x", null, DrawingValues.Text(left));
        await xml.WriteAttributeStringAsync(null, "y", null, DrawingValues.Text(top));
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "chExt", Main);
        await xml.WriteAttributeStringAsync(null, "cx", null, DrawingValues.Text(Math.Max(right - left, 1)));
        await xml.WriteAttributeStringAsync(null, "cy", null, DrawingValues.Text(Math.Max(bottom - top, 1)));
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();

        foreach (GroupItem item in group.Items)
        {
            if (item.Shape is not null)
            {
                await WriteShapeAsync(xml, item.Shape, context, item.Bounds);
            }
            else if (item.Picture is not null)
            {
                await WritePictureAsync(xml, item.Picture, context, item.Bounds);
            }
            else
            {
                await WriteGroupAsync(xml, item.Group!, context, item.Bounds);
            }
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>A graphic frame that shows a chart part; the chart itself is written with <see cref="ChartWriter"/>.</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="chart">What the owner needs to write the chart part; it is handed back in <see cref="DrawingContext.Targets"/>.</param>
    /// <param name="description">The alternative text, or null.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size.</param>
    /// <returns>A task.</returns>
    public static async Task WriteChartFrameAsync(XmlWriter xml, object chart, string? description, DrawingContext context, DrawingBounds frame)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(chart);
        ArgumentNullException.ThrowIfNull(context);
        ShapeTreeNamespace tree = context.Tree;
        int shapeId = context.NextShapeId();
        string relationshipId = context.Chart(chart);
        await xml.WriteStartElementAsync(tree.Prefix, "graphicFrame", tree.Uri);

        if (tree.MacroAttributes == true)
        {
            await xml.WriteAttributeStringAsync(null, "macro", null, string.Empty);
        }

        await xml.WriteStartElementAsync(tree.Prefix, "nvGraphicFramePr", tree.Uri);
        await WriteNonVisualPropertiesAsync(xml, tree, shapeId, "Chart " + DrawingValues.Text(shapeId - 1), description, null, context);
        await xml.WriteElementStringAsync(tree.Prefix, "cNvGraphicFramePr", tree.Uri, string.Empty);
        await WriteApplicationPropertiesAsync(xml, tree);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync(tree.Prefix, "xfrm", tree.Uri);
        await WriteOffsetAndExtentAsync(xml, frame);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "graphic", Main);
        await xml.WriteStartElementAsync("a", "graphicData", Main);
        await xml.WriteAttributeStringAsync(null, "uri", null, OpcNamespaces.DrawingChart);
        await xml.WriteStartElementAsync("c", "chart", OpcNamespaces.DrawingChart);
        await xml.WriteAttributeStringAsync("r", "id", OpcNamespaces.Relationships, relationshipId);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    /// <summary>An a:solidFill, a:gradFill, a:pattFill, a:blipFill or a:noFill.</summary>
    /// <param name="xml">The writer.</param>
    /// <param name="fill">The fill.</param>
    /// <param name="context">The part's relationships, which a picture fill needs; null where there are none, as in a chart.</param>
    /// <returns>A task.</returns>
    public static async Task WriteFillAsync(XmlWriter xml, Fill fill, DrawingContext? context = null)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(fill);

        switch (fill.Kind)
        {
            case FillKind.None:
                await xml.WriteElementStringAsync("a", "noFill", Main, string.Empty);

                break;

            case FillKind.Solid:
                await xml.WriteStartElementAsync("a", "solidFill", Main);
                await WriteColorAsync(xml, fill.Color!, fill.Transparency);
                await xml.WriteEndElementAsync();

                break;

            case FillKind.Pattern:
                await xml.WriteStartElementAsync("a", "pattFill", Main);
                await xml.WriteAttributeStringAsync(null, "prst", null, DrawingNames.Pattern(fill.PatternType));
                await xml.WriteStartElementAsync("a", "fgClr", Main);
                await WriteColorAsync(xml, fill.Color!, 0);
                await xml.WriteEndElementAsync();
                await xml.WriteStartElementAsync("a", "bgClr", Main);
                await WriteColorAsync(xml, fill.BackgroundColor!, 0);
                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();

                break;

            case FillKind.Picture:
                if (context is null)
                {
                    throw new InvalidOperationException("A picture fill needs a DrawingContext for the image's relationship.");
                }

                await xml.WriteStartElementAsync("a", "blipFill", Main);
                await xml.WriteAttributeStringAsync(null, "rotWithShape", null, "1");
                await WriteBlipAsync(xml, fill.Image!, context, 0);
                await xml.WriteStartElementAsync("a", "stretch", Main);
                await xml.WriteElementStringAsync("a", "fillRect", Main, string.Empty);
                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();

                break;

            default:
                await xml.WriteStartElementAsync("a", "gradFill", Main);
                await xml.WriteAttributeStringAsync(null, "rotWithShape", null, "1");
                await xml.WriteStartElementAsync("a", "gsLst", Main);

                foreach (GradientStop stop in fill.Stops.OrderBy(stop => stop.Position))
                {
                    await xml.WriteStartElementAsync("a", "gs", Main);
                    await xml.WriteAttributeStringAsync(null, "pos", null, DrawingValues.Text(DrawingValues.Percent(stop.Position)));
                    await WriteColorAsync(xml, stop.Color, stop.Transparency);
                    await xml.WriteEndElementAsync();
                }

                await xml.WriteEndElementAsync();
                await xml.WriteStartElementAsync("a", "lin", Main);
                await xml.WriteAttributeStringAsync(null, "ang", null, DrawingValues.Text(DrawingValues.Angle(fill.Angle)));
                await xml.WriteAttributeStringAsync(null, "scaled", null, "0");
                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();

                break;
        }
    }

    /// <summary>An a:ln.</summary>
    /// <param name="xml">The writer.</param>
    /// <param name="outline">The line.</param>
    /// <returns>A task.</returns>
    public static async Task WriteOutlineAsync(XmlWriter xml, Outline outline)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(outline);
        await xml.WriteStartElementAsync("a", "ln", Main);

        if (outline.Color is null)
        {
            await xml.WriteElementStringAsync("a", "noFill", Main, string.Empty);
            await xml.WriteEndElementAsync();

            return;
        }

        await xml.WriteAttributeStringAsync(null, "w", null, DrawingValues.Text(DrawingValues.Emu(outline.Width)));
        await xml.WriteStartElementAsync("a", "solidFill", Main);
        await WriteColorAsync(xml, outline.Color, outline.Transparency);
        await xml.WriteEndElementAsync();

        if (outline.Dash != LineDash.Solid)
        {
            await xml.WriteStartElementAsync("a", "prstDash", Main);
            await xml.WriteAttributeStringAsync(null, "val", null, DrawingNames.Dash(outline.Dash));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>An a:effectLst with the shadows and soft edges given, in the schema's order.</summary>
    /// <param name="xml">The writer.</param>
    /// <param name="outer">The shadow behind, or null.</param>
    /// <param name="inner">The shadow inside, or null.</param>
    /// <param name="softEdges">Soft edges in points, or null.</param>
    /// <returns>A task.</returns>
    public static async Task WriteEffectsAsync(XmlWriter xml, Shadow? outer, Shadow? inner, double? softEdges)
    {
        ArgumentNullException.ThrowIfNull(xml);
        await xml.WriteStartElementAsync("a", "effectLst", Main);

        if (inner is not null)
        {
            await xml.WriteStartElementAsync("a", "innerShdw", Main);
            await WriteShadowAttributesAsync(xml, inner);
            await WriteColorAsync(xml, inner.Color, inner.Transparency);
            await xml.WriteEndElementAsync();
        }

        if (outer is not null)
        {
            await xml.WriteStartElementAsync("a", "outerShdw", Main);
            await WriteShadowAttributesAsync(xml, outer);
            await xml.WriteAttributeStringAsync(null, "algn", null, "tl");
            await xml.WriteAttributeStringAsync(null, "rotWithShape", null, "0");
            await WriteColorAsync(xml, outer.Color, outer.Transparency);
            await xml.WriteEndElementAsync();
        }

        if (softEdges is double points)
        {
            await xml.WriteStartElementAsync("a", "softEdge", Main);
            await xml.WriteAttributeStringAsync(null, "rad", null, DrawingValues.Text(DrawingValues.Emu(points)));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>
    /// An a:rPr, a:defRPr or a:endParaRPr. Bold and italic are written either way, so that a
    /// font given without them overrides a bold default such as a title's.
    /// </summary>
    /// <param name="xml">The writer.</param>
    /// <param name="element">rPr, defRPr or endParaRPr.</param>
    /// <param name="font">The font, or null for the element alone.</param>
    /// <returns>A task.</returns>
    public static async Task WriteRunPropertiesAsync(XmlWriter xml, string element, TextFont? font)
    {
        ArgumentNullException.ThrowIfNull(xml);
        ArgumentNullException.ThrowIfNull(element);
        await xml.WriteStartElementAsync("a", element, Main);

        if (element != "defRPr")
        {
            await xml.WriteAttributeStringAsync(null, "lang", null, "en-US");
        }

        if (font is null)
        {
            await xml.WriteEndElementAsync();

            return;
        }

        if (font.Size is double size)
        {
            await xml.WriteAttributeStringAsync(null, "sz", null, DrawingValues.Text((long)Math.Round(size * 100)));
        }

        await xml.WriteAttributeStringAsync(null, "b", null, DrawingValues.Flag(font.Bold));
        await xml.WriteAttributeStringAsync(null, "i", null, DrawingValues.Flag(font.Italic));

        if (font.Color is not null)
        {
            await xml.WriteStartElementAsync("a", "solidFill", Main);
            await WriteColorAsync(xml, font.Color, 0);
            await xml.WriteEndElementAsync();
        }

        if (font.Name is not null)
        {
            await xml.WriteStartElementAsync("a", "latin", Main);
            await xml.WriteAttributeStringAsync(null, "typeface", null, font.Name);
            await xml.WriteEndElementAsync();
            await xml.WriteStartElementAsync("a", "cs", Main);
            await xml.WriteAttributeStringAsync(null, "typeface", null, font.Name);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>An a:off and an a:ext, in EMU.</summary>
    /// <param name="xml">The writer, inside an xfrm.</param>
    /// <param name="frame">The place and size.</param>
    /// <returns>A task.</returns>
    public static async Task WriteOffsetAndExtentAsync(XmlWriter xml, DrawingBounds frame)
    {
        ArgumentNullException.ThrowIfNull(xml);
        await xml.WriteStartElementAsync("a", "off", Main);
        await xml.WriteAttributeStringAsync(null, "x", null, DrawingValues.Text(frame.Left.Emu));
        await xml.WriteAttributeStringAsync(null, "y", null, DrawingValues.Text(frame.Top.Emu));
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "ext", Main);
        await xml.WriteAttributeStringAsync(null, "cx", null, DrawingValues.Text(frame.Width.Emu));
        await xml.WriteAttributeStringAsync(null, "cy", null, DrawingValues.Text(frame.Height.Emu));
        await xml.WriteEndElementAsync();
    }

    /// <summary>The a:blip of a picture: the raster image, with the SVG as an extension older versions skip.</summary>
    private static async Task WriteBlipAsync(XmlWriter xml, Picture picture, DrawingContext context, double transparency)
    {
        await xml.WriteStartElementAsync("a", "blip", Main);

        // An SVG is stored as Office 2016 stores it: the raster fallback is the blip, and the
        // SVG an extension of it.
        Picture raster = picture.Fallback ?? picture;
        await xml.WriteAttributeStringAsync("r", "embed", OpcNamespaces.Relationships, context.Image(raster));

        if (transparency > 0)
        {
            await xml.WriteStartElementAsync("a", "alphaModFix", Main);
            await xml.WriteAttributeStringAsync(null, "amt", null, DrawingValues.Text(DrawingValues.Percent(1 - transparency)));
            await xml.WriteEndElementAsync();
        }

        if (picture.Fallback is not null)
        {
            await xml.WriteStartElementAsync("a", "extLst", Main);
            await xml.WriteStartElementAsync("a", "ext", Main);
            await xml.WriteAttributeStringAsync(null, "uri", null, SvgExtensionUri);
            await xml.WriteStartElementAsync("asvg", "svgBlip", OpcNamespaces.DrawingSvg);
            await xml.WriteAttributeStringAsync("r", "embed", OpcNamespaces.Relationships, context.Image(picture));
            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>The theme references Office gives a new shape, so that what is left unset looks as in Office.</summary>
    private static async Task WriteStyleAsync(XmlWriter xml, ShapeTreeNamespace tree, ShapeKind? kind)
    {
        (string Line, string Fill, string Effect, string Font, bool Shade) style = kind switch
        {
            ShapeKind.Line => ("1", "0", "0", "tx1", false),
            ShapeKind.TextBox => ("0", "0", "0", "dk1", false),
            _ => ("2", "1", "0", "lt1", true),
        };

        await xml.WriteStartElementAsync(tree.Prefix, "style", tree.Uri);
        await WriteStyleReferenceAsync(xml, "lnRef", style.Line, "accent1", style.Shade);
        await WriteStyleReferenceAsync(xml, "fillRef", style.Fill, "accent1", false);
        await WriteStyleReferenceAsync(xml, "effectRef", style.Effect, "accent1", false);
        await xml.WriteStartElementAsync("a", "fontRef", Main);
        await xml.WriteAttributeStringAsync(null, "idx", null, "minor");
        await xml.WriteStartElementAsync("a", "schemeClr", Main);
        await xml.WriteAttributeStringAsync(null, "val", null, style.Font);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteStyleReferenceAsync(XmlWriter xml, string element, string index, string color, bool shade)
    {
        await xml.WriteStartElementAsync("a", element, Main);
        await xml.WriteAttributeStringAsync(null, "idx", null, index);
        await xml.WriteStartElementAsync("a", "schemeClr", Main);
        await xml.WriteAttributeStringAsync(null, "val", null, color);

        if (shade == true)
        {
            await xml.WriteStartElementAsync("a", "shade", Main);
            await xml.WriteAttributeStringAsync(null, "val", null, "15000");
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    /// <summary>The shape's txBody: one a:p per line of the runs, each run with its own look.</summary>
    private static async Task WriteTextBodyAsync(XmlWriter xml, ShapeTreeNamespace tree, ShapeText? text)
    {
        await xml.WriteStartElementAsync(tree.Prefix, "txBody", tree.Uri);
        await xml.WriteStartElementAsync("a", "bodyPr", Main);
        await xml.WriteAttributeStringAsync(null, "vertOverflow", null, "clip");
        await xml.WriteAttributeStringAsync(null, "horzOverflow", null, "clip");

        if (text is not null)
        {
            if (text.Wrap == true)
            {
                await xml.WriteAttributeStringAsync(null, "wrap", null, "square");
            }
            else
            {
                await xml.WriteAttributeStringAsync(null, "wrap", null, "none");
            }

            await WriteInsetAsync(xml, "lIns", text.MarginLeft);
            await WriteInsetAsync(xml, "tIns", text.MarginTop);
            await WriteInsetAsync(xml, "rIns", text.MarginRight);
            await WriteInsetAsync(xml, "bIns", text.MarginBottom);
        }

        await xml.WriteAttributeStringAsync(null, "rtlCol", null, "0");
        await xml.WriteAttributeStringAsync(null, "anchor", null, DrawingNames.Anchor(text?.Vertical ?? TextAnchor.Top));
        await xml.WriteEndElementAsync();
        await xml.WriteElementStringAsync("a", "lstStyle", Main, string.Empty);
        string alignment = DrawingNames.Alignment(text?.Horizontal ?? TextAlignment.Left);

        if (text is null)
        {
            await xml.WriteStartElementAsync("a", "p", Main);
            await WriteParagraphPropertiesAsync(xml, alignment);
            await WriteRunPropertiesAsync(xml, "endParaRPr", new TextFont { Size = 11 });
            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();

            return;
        }

        foreach (List<TextRun> paragraph in Paragraphs(text.Runs))
        {
            await xml.WriteStartElementAsync("a", "p", Main);
            await WriteParagraphPropertiesAsync(xml, alignment);
            TextFont? last = null;

            foreach (TextRun run in paragraph)
            {
                last = run.Font;

                if (run.Text.Length == 0)
                {
                    continue;
                }

                await xml.WriteStartElementAsync("a", "r", Main);
                await WriteRunPropertiesAsync(xml, "rPr", run.Font);
                await xml.WriteElementStringAsync("a", "t", Main, run.Text);
                await xml.WriteEndElementAsync();
            }

            // Gives an empty line its height, and the cursor its look at the end of a line.
            await WriteRunPropertiesAsync(xml, "endParaRPr", last);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
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

    private static async Task WriteParagraphPropertiesAsync(XmlWriter xml, string alignment)
    {
        await xml.WriteStartElementAsync("a", "pPr", Main);
        await xml.WriteAttributeStringAsync(null, "algn", null, alignment);
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteInsetAsync(XmlWriter xml, string name, double? points)
    {
        if (points is double value)
        {
            await xml.WriteAttributeStringAsync(null, name, null, DrawingValues.Text(DrawingValues.Emu(value)));
        }
    }

    private static async Task WritePresetGeometryAsync(XmlWriter xml, string preset, IReadOnlyList<ShapeAdjustment> adjustments)
    {
        await xml.WriteStartElementAsync("a", "prstGeom", Main);
        await xml.WriteAttributeStringAsync(null, "prst", null, preset);
        await xml.WriteStartElementAsync("a", "avLst", Main);

        foreach (ShapeAdjustment adjustment in adjustments)
        {
            await xml.WriteStartElementAsync("a", "gd", Main);
            await xml.WriteAttributeStringAsync(null, "name", null, adjustment.Name);
            await xml.WriteAttributeStringAsync(null, "fmla", null, "val " + DrawingValues.Text(adjustment.Value));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    /// <summary>A cNvPr: id, name, alternative text and the click's link.</summary>
    private static async Task WriteNonVisualPropertiesAsync(XmlWriter xml, ShapeTreeNamespace tree, int shapeId, string name, string? description, Hyperlink? link, DrawingContext context)
    {
        await xml.WriteStartElementAsync(tree.Prefix, "cNvPr", tree.Uri);
        await xml.WriteAttributeStringAsync(null, "id", null, DrawingValues.Text(shapeId));
        await xml.WriteAttributeStringAsync(null, "name", null, name);

        if (description is not null)
        {
            await xml.WriteAttributeStringAsync(null, "descr", null, description);
        }

        if (link is not null)
        {
            await xml.WriteStartElementAsync("a", "hlinkClick", Main);
            await xml.WriteAttributeStringAsync("r", "id", OpcNamespaces.Relationships, context.Link(link));

            if (link.Tooltip is not null)
            {
                await xml.WriteAttributeStringAsync(null, "tooltip", null, link.Tooltip);
            }

            if (link.Action is not null)
            {
                await xml.WriteAttributeStringAsync(null, "action", null, link.Action);
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteApplicationPropertiesAsync(XmlWriter xml, ShapeTreeNamespace tree)
    {
        if (tree.ApplicationProperties == true)
        {
            await xml.WriteElementStringAsync(tree.Prefix, "nvPr", tree.Uri, string.Empty);
        }
    }

    private static async Task WriteShadowAttributesAsync(XmlWriter xml, Shadow shadow)
    {
        await xml.WriteAttributeStringAsync(null, "blurRad", null, DrawingValues.Text(DrawingValues.Emu(shadow.Blur)));
        await xml.WriteAttributeStringAsync(null, "dist", null, DrawingValues.Text(DrawingValues.Emu(shadow.Distance)));
        await xml.WriteAttributeStringAsync(null, "dir", null, DrawingValues.Text(DrawingValues.Angle(shadow.Angle)));
    }

    /// <summary>An a:srgbClr, with an a:alpha when the colour is see-through.</summary>
    private static async Task WriteColorAsync(XmlWriter xml, string color, double transparency)
    {
        await xml.WriteStartElementAsync("a", "srgbClr", Main);
        await xml.WriteAttributeStringAsync(null, "val", null, color[1..]);

        if (transparency > 0)
        {
            await xml.WriteStartElementAsync("a", "alpha", Main);
            await xml.WriteAttributeStringAsync(null, "val", null, DrawingValues.Text(DrawingValues.Percent(1 - transparency)));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }
}
