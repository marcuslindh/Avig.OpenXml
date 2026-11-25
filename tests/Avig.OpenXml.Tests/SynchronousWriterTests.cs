using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Avig.OpenXml.Tests;

/// <summary>
/// The synchronous writer: a writer without Async, and a writer into an XElement, give the same
/// XML as the task-returning forms, which a package editing a part in memory needs.
/// </summary>
public class SynchronousWriterTests
{
    [Fact]
    public async Task A_synchronous_writer_writes_what_the_asynchronous_one_does()
    {
        Shape shape = new(ShapeKind.RoundedRectangle)
        {
            Fill = Fill.Solid("#2384F9", 0.25),
            Outline = new Outline("#1F2940", 1.5),
            Shadow = new Shadow { Angle = 90, Blur = 6 },
            Text = new ShapeText("Text"),
            Link = Hyperlink.ToUrl("https://example.com/"),
        };

        ShapeGroup group = new([
            GroupItem.Of(new Shape(ShapeKind.Ellipse), DrawingBounds.FromPoints(0, 0, 10, 10)),
            GroupItem.Of(DrawingKit.PngPicture(), DrawingBounds.FromPoints(10, 10, 10, 10)),
        ]);

        foreach (ShapeTreeNamespace tree in DrawingKit.Hosts)
        {
            (string expected, _) = await DrawingKit.WriteAsync(
                tree,
                (xml, context) => DrawingWriter.WriteShapeAsync(xml, shape, context, DrawingKit.Frame),
                (xml, context) => DrawingWriter.WritePictureAsync(xml, DrawingKit.SvgPicture(), context, DrawingKit.Frame),
                (xml, context) => DrawingWriter.WriteGroupAsync(xml, group, context, DrawingKit.Frame));

            List<XElement> written = WriteSynchronously(tree, (xml, context) =>
            {
                DrawingWriter.WriteShape(xml, shape, context, DrawingKit.Frame);
                DrawingWriter.WritePicture(xml, DrawingKit.SvgPicture(), context, DrawingKit.Frame);
                DrawingWriter.WriteGroup(xml, group, context, DrawingKit.Frame);
            });

            List<string> expectedItems = DrawingKit.Items(expected).Select(element => element.ToString(SaveOptions.DisableFormatting)).ToList();
            Assert.Equal(expectedItems, written.Select(element => element.ToString(SaveOptions.DisableFormatting)).ToList());
        }
    }

    [Fact]
    public void A_writer_into_an_element_needs_no_task_and_no_stream()
    {
        XElement root = new(XName.Get("root", OpcNamespaces.PresentationMain), new XAttribute(XNamespace.Xmlns + "a", OpcNamespaces.DrawingMain));

        using (XmlWriter xml = root.CreateWriter())
        {
            DrawingWriter.WriteFill(xml, Fill.Solid("#FF0000"));
            DrawingWriter.WriteOutline(xml, new Outline("#00FF00", 2));
            DrawingWriter.WriteEffects(xml, new Shadow(), null, 2);
            DrawingWriter.WriteRunProperties(xml, "rPr", new TextFont { Bold = true, Size = 20 });
        }

        Assert.Equal(["solidFill", "ln", "effectLst", "rPr"], root.Elements().Select(element => element.Name.LocalName).ToList());
        Assert.All(root.Elements(), element => Assert.Equal(OpcNamespaces.DrawingMain, element.Name.NamespaceName));
    }

    [Fact]
    public async Task A_chart_part_is_the_same_written_synchronously()
    {
        ChartSpace chart = new(ChartType.Column)
        {
            Title = "Sales",
            Series = [new ChartSpaceSeries(ChartData.FromNumbers([1, 2, 3])) { Name = "2026", Categories = ChartData.FromTexts(["a", "b", "c"]) }],
        };

        string expected = await DrawingKit.WriteChartAsync(chart);
        StringBuilder text = new();

        using (XmlWriter xml = XmlWriter.Create(text, new XmlWriterSettings { OmitXmlDeclaration = true }))
        {
            ChartWriter.WriteChartSpace(xml, chart);
        }

        Assert.Equal(expected, text.ToString());
    }

    private static List<XElement> WriteSynchronously(ShapeTreeNamespace tree, Action<XmlWriter, DrawingContext> write)
    {
        DrawingContext context = new(tree, [], 2);
        StringBuilder text = new();

        using (XmlWriter xml = XmlWriter.Create(text, new XmlWriterSettings { OmitXmlDeclaration = true }))
        {
            xml.WriteStartElement(tree.Prefix, "root", tree.Uri);
            xml.WriteAttributeString("xmlns", "a", null, OpcNamespaces.DrawingMain);
            xml.WriteAttributeString("xmlns", "r", null, OpcNamespaces.Relationships);
            write(xml, context);
            xml.WriteEndElement();
        }

        return XElement.Parse(text.ToString()).Elements().ToList();
    }
}
