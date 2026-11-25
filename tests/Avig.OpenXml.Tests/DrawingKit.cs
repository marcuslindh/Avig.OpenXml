using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Validation;

namespace Avig.OpenXml.Tests;

/// <summary>
/// Writes DrawingML the way the two hosts do, a spreadsheet drawing (xdr) and a slide (p), reads
/// it back, and runs Microsoft's schema validator over it. The hosts' wrappers are built here, as
/// Avig.Kalkyl and Avig.Bildspel build their own; the layer does not know them.
/// </summary>
internal static class DrawingKit
{
    public const string Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    public static ShapeTreeNamespace Spreadsheet { get; } = new("xdr", Xdr) { MacroAttributes = true };

    public static ShapeTreeNamespace Slide { get; } = new("p", OpcNamespaces.PresentationMain) { ApplicationProperties = true };

    public static IEnumerable<ShapeTreeNamespace> Hosts => [Spreadsheet, Slide];

    public static DrawingBounds Frame { get; } = DrawingBounds.FromPoints(10, 20, 200, 100);

    /// <summary>
    /// The XML of one part: the host's root around what each of <paramref name="items"/> writes
    /// (in a spreadsheet drawing each in an anchor of its own), with the context they used.
    /// </summary>
    public static async Task<(string Xml, DrawingContext Context)> WriteAsync(ShapeTreeNamespace tree, params Func<XmlWriter, DrawingContext, Task>[] items)
    {
        DrawingContext context = new(tree, [], 2);
        StringBuilder text = new();
        XmlWriterSettings settings = new() { Async = true, OmitXmlDeclaration = true };

        await using (XmlWriter xml = XmlWriter.Create(text, settings))
        {
            if (tree == Spreadsheet)
            {
                await xml.WriteStartElementAsync("xdr", "wsDr", Xdr);
                await xml.WriteAttributeStringAsync("xmlns", "a", null, OpcNamespaces.DrawingMain);
                await xml.WriteAttributeStringAsync("xmlns", "r", null, OpcNamespaces.Relationships);

                foreach (Func<XmlWriter, DrawingContext, Task> write in items)
                {
                    await xml.WriteStartElementAsync("xdr", "absoluteAnchor", Xdr);
                    await xml.WriteStartElementAsync("xdr", "pos", Xdr);
                    await xml.WriteAttributeStringAsync(null, "x", null, "0");
                    await xml.WriteAttributeStringAsync(null, "y", null, "0");
                    await xml.WriteEndElementAsync();
                    await xml.WriteStartElementAsync("xdr", "ext", Xdr);
                    await xml.WriteAttributeStringAsync(null, "cx", null, "100");
                    await xml.WriteAttributeStringAsync(null, "cy", null, "100");
                    await xml.WriteEndElementAsync();
                    await write(xml, context);
                    await xml.WriteElementStringAsync("xdr", "clientData", Xdr, string.Empty);
                    await xml.WriteEndElementAsync();
                }

                await xml.WriteEndElementAsync();
            }
            else
            {
                string presentation = OpcNamespaces.PresentationMain;
                await xml.WriteStartElementAsync("p", "sld", presentation);
                await xml.WriteAttributeStringAsync("xmlns", "a", null, OpcNamespaces.DrawingMain);
                await xml.WriteAttributeStringAsync("xmlns", "r", null, OpcNamespaces.Relationships);
                await xml.WriteStartElementAsync("p", "cSld", presentation);
                await xml.WriteStartElementAsync("p", "spTree", presentation);
                await xml.WriteStartElementAsync("p", "nvGrpSpPr", presentation);
                await xml.WriteStartElementAsync("p", "cNvPr", presentation);
                await xml.WriteAttributeStringAsync(null, "id", null, "1");
                await xml.WriteAttributeStringAsync(null, "name", null, string.Empty);
                await xml.WriteEndElementAsync();
                await xml.WriteElementStringAsync("p", "cNvGrpSpPr", presentation, string.Empty);
                await xml.WriteElementStringAsync("p", "nvPr", presentation, string.Empty);
                await xml.WriteEndElementAsync();
                await xml.WriteElementStringAsync("p", "grpSpPr", presentation, string.Empty);

                foreach (Func<XmlWriter, DrawingContext, Task> write in items)
                {
                    await write(xml, context);
                }

                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();
            }
        }

        return (text.ToString(), context);
    }

    /// <summary>The XML of one shape, picture or group on its own, the way the host writes it.</summary>
    public static async Task<(string Xml, DrawingContext Context)> WriteItemAsync(ShapeTreeNamespace tree, object item)
    {
        return await WriteAsync(tree, (xml, context) => item switch
        {
            Shape shape => DrawingWriter.WriteShapeAsync(xml, shape, context, Frame),
            Picture picture => DrawingWriter.WritePictureAsync(xml, picture, context, Frame),
            ShapeGroup group => DrawingWriter.WriteGroupAsync(xml, group, context, Frame),
            _ => DrawingWriter.WriteChartFrameAsync(xml, item, "A chart", context, Frame),
        });
    }

    /// <summary>The shape tree elements of a part written by <see cref="WriteAsync"/>.</summary>
    public static List<XElement> Items(string xml)
    {
        XElement root = XElement.Parse(xml);

        return root.Descendants().Where(element => element.Name.LocalName is "sp" or "cxnSp" or "pic" or "grpSp" or "graphicFrame")
            .Where(element => element.Parent is not null && element.Parent.Name.LocalName is "absoluteAnchor" or "spTree").ToList();
    }

    /// <summary>What a host would read the part with: its relationships, and the images behind them.</summary>
    public static DrawingReadContext ReadContext(DrawingContext written)
    {
        List<OpcRelationship> relationships = [];
        Dictionary<string, byte[]> images = [];

        foreach (DrawingTarget target in written.Targets)
        {
            if (target.Image is not null)
            {
                relationships.Add(new OpcRelationship(target.Id, OpcNamespaces.ImageRelationship, "../media/" + target.Id + "." + target.Image.Extension, false));
                images[target.Id] = target.Image.Data.ToArray();
            }
            else if (target.Link is not null && target.Link.IsExternal == true)
            {
                relationships.Add(new OpcRelationship(target.Id, OpcNamespaces.HyperlinkRelationship, target.Link.Url!, true));
            }
            else if (target.Link is not null)
            {
                relationships.Add(new OpcRelationship(target.Id, OpcNamespaces.HyperlinkRelationship, "#" + target.Link.Location, false));
            }
            else
            {
                relationships.Add(new OpcRelationship(target.Id, OpcNamespaces.ChartRelationship, "../charts/chart1.xml", false));
            }
        }

        return new DrawingReadContext(relationships, (relationship, _) =>
        {
            if (images.TryGetValue(relationship.Id, out byte[]? bytes) == true)
            {
                return Task.FromResult<Result<byte[]>>(bytes);
            }

            return Task.FromResult(Result<byte[]>.NotFound("No such image."));
        });
    }

    /// <summary>Writes an item, reads it back, and writes what was read: the two must be the same XML.</summary>
    public static async Task<ShapeTreeItem> RoundTripAsync(ShapeTreeNamespace tree, object item)
    {
        (string first, DrawingContext context) = await WriteItemAsync(tree, item);
        Assert.Empty(ValidatePart(tree, first));
        XElement element = Assert.Single(Items(first));
        Result<ShapeTreeItem> read = await DrawingReader.ReadAsync(element, ReadContext(context));
        Assert.True(read.IsSuccess(out ShapeTreeItem? readItem), read.Error);

        object again = readItem.Kind switch
        {
            ShapeTreeItemKind.Shape => readItem.Shape!,
            ShapeTreeItemKind.Picture => readItem.Picture!,
            ShapeTreeItemKind.Group => readItem.Group!,
            _ => item,
        };

        (string second, _) = await WriteItemAsync(tree, again);
        Assert.Equal(first, second);

        return readItem;
    }

    /// <summary>Microsoft's schema validator over a part written by <see cref="WriteAsync"/>.</summary>
    public static List<string> ValidatePart(ShapeTreeNamespace tree, string xml)
    {
        DocumentFormat.OpenXml.OpenXmlElement element;

        if (tree == Spreadsheet)
        {
            element = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing(xml);
        }
        else
        {
            element = new DocumentFormat.OpenXml.Presentation.Slide(xml);
        }

        return Validate(element);
    }

    /// <summary>Microsoft's schema validator over a chart part.</summary>
    public static List<string> ValidateChart(string xml)
    {
        return Validate(new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(xml));
    }

    private static List<string> Validate(DocumentFormat.OpenXml.OpenXmlElement element)
    {
        OpenXmlValidator validator = new(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        List<string> problems = [];

        foreach (ValidationErrorInfo error in validator.Validate(element))
        {
            problems.Add($"{error.Path?.XPath}: {error.Description}");
        }

        return problems;
    }

    /// <summary>A chart part as text, from the declaration on.</summary>
    public static async Task<string> WriteChartAsync(ChartSpace chart)
    {
        StringBuilder text = new();
        XmlWriterSettings settings = new() { Async = true, OmitXmlDeclaration = true };

        await using (XmlWriter xml = XmlWriter.Create(text, settings))
        {
            await ChartWriter.WriteChartSpaceAsync(xml, chart);
        }

        return text.ToString();
    }

    /// <summary>A transparent 1×1 PNG.</summary>
    public static byte[] Png { get; } = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNgYGBgAAAABQABeqhXUAAAAABJRU5ErkJggg==");

    public static Picture PngPicture()
    {
        Assert.True(Picture.FromBytes(Png).IsSuccess(out Picture? picture));

        return picture;
    }

    public static Picture SvgPicture()
    {
        Result<Picture> read = Picture.FromSvg("""<svg xmlns="http://www.w3.org/2000/svg" width="48" height="24" viewBox="0 0 48 24"><rect width="48" height="24" fill="#2384F9"/></svg>"""u8, PngPicture());
        Assert.True(read.IsSuccess(out Picture? picture), read.Error);

        return picture;
    }
}
