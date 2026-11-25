using System.Xml;

namespace Avig.OpenXml.Drawing;

/// <summary>
/// The task-returning forms of the writer, kept for the packages that write a part with an
/// asynchronous <see cref="XmlWriter"/>. They write synchronously into the writer's buffer, which
/// is what they always did between the writer's flushes; the caller flushes the part.
/// </summary>
public static partial class DrawingWriter
{
    /// <summary>A shape; see <see cref="WriteShape"/>.</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="shape">The shape.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size (a:off, a:ext).</param>
    /// <returns>A completed task.</returns>
    public static Task WriteShapeAsync(XmlWriter xml, Shape shape, DrawingContext context, DrawingBounds frame)
    {
        WriteShape(xml, shape, context, frame);

        return Task.CompletedTask;
    }

    /// <summary>A picture; see <see cref="WritePicture"/>.</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="picture">The picture.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size.</param>
    /// <returns>A completed task.</returns>
    public static Task WritePictureAsync(XmlWriter xml, Picture picture, DrawingContext context, DrawingBounds frame)
    {
        WritePicture(xml, picture, context, frame);

        return Task.CompletedTask;
    }

    /// <summary>A group; see <see cref="WriteGroup"/>.</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="group">The group.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size of the whole group.</param>
    /// <returns>A completed task.</returns>
    public static Task WriteGroupAsync(XmlWriter xml, ShapeGroup group, DrawingContext context, DrawingBounds frame)
    {
        WriteGroup(xml, group, context, frame);

        return Task.CompletedTask;
    }

    /// <summary>A chart frame; see <see cref="WriteChartFrame"/>.</summary>
    /// <param name="xml">The writer, positioned where the element goes.</param>
    /// <param name="chart">What the owner needs to write the chart part.</param>
    /// <param name="description">The alternative text, or null.</param>
    /// <param name="context">The part's ids and relationships.</param>
    /// <param name="frame">The place and size.</param>
    /// <returns>A completed task.</returns>
    public static Task WriteChartFrameAsync(XmlWriter xml, object chart, string? description, DrawingContext context, DrawingBounds frame)
    {
        WriteChartFrame(xml, chart, description, context, frame);

        return Task.CompletedTask;
    }

    /// <summary>A fill; see <see cref="WriteFill"/>.</summary>
    /// <param name="xml">The writer.</param>
    /// <param name="fill">The fill.</param>
    /// <param name="context">The part's relationships, or null.</param>
    /// <returns>A completed task.</returns>
    public static Task WriteFillAsync(XmlWriter xml, Fill fill, DrawingContext? context = null)
    {
        WriteFill(xml, fill, context);

        return Task.CompletedTask;
    }

    /// <summary>A line; see <see cref="WriteOutline"/>.</summary>
    /// <param name="xml">The writer.</param>
    /// <param name="outline">The line.</param>
    /// <returns>A completed task.</returns>
    public static Task WriteOutlineAsync(XmlWriter xml, Outline outline)
    {
        WriteOutline(xml, outline);

        return Task.CompletedTask;
    }

    /// <summary>The effects; see <see cref="WriteEffects"/>.</summary>
    /// <param name="xml">The writer.</param>
    /// <param name="outer">The shadow behind, or null.</param>
    /// <param name="inner">The shadow inside, or null.</param>
    /// <param name="softEdges">Soft edges in points, or null.</param>
    /// <returns>A completed task.</returns>
    public static Task WriteEffectsAsync(XmlWriter xml, Shadow? outer, Shadow? inner, double? softEdges)
    {
        WriteEffects(xml, outer, inner, softEdges);

        return Task.CompletedTask;
    }

    /// <summary>Run properties; see <see cref="WriteRunProperties"/>.</summary>
    /// <param name="xml">The writer.</param>
    /// <param name="element">rPr, defRPr or endParaRPr.</param>
    /// <param name="font">The font, or null for the element alone.</param>
    /// <returns>A completed task.</returns>
    public static Task WriteRunPropertiesAsync(XmlWriter xml, string element, TextFont? font)
    {
        WriteRunProperties(xml, element, font);

        return Task.CompletedTask;
    }

    /// <summary>An a:off and an a:ext; see <see cref="WriteOffsetAndExtent"/>.</summary>
    /// <param name="xml">The writer, inside an xfrm.</param>
    /// <param name="frame">The place and size.</param>
    /// <returns>A completed task.</returns>
    public static Task WriteOffsetAndExtentAsync(XmlWriter xml, DrawingBounds frame)
    {
        WriteOffsetAndExtent(xml, frame);

        return Task.CompletedTask;
    }
}

/// <summary>The task-returning form of the chart writer; see <see cref="DrawingWriter"/>.</summary>
public static partial class ChartWriter
{
    /// <summary>A complete chart part; see <see cref="WriteChartSpace"/>.</summary>
    /// <param name="xml">The writer of the part.</param>
    /// <param name="chart">The chart.</param>
    /// <returns>A completed task.</returns>
    public static Task WriteChartSpaceAsync(XmlWriter xml, ChartSpace chart)
    {
        WriteChartSpace(xml, chart);

        return Task.CompletedTask;
    }
}
