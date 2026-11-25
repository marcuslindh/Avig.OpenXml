using System.Xml.Linq;

namespace Avig.OpenXml.Tests;

/// <summary>Theme colours and colour modifiers, and the synchronous reader for parts already in memory (for Avig.Bildspel).</summary>
public class DrawingColorsTests
{
    private const string Main = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly Dictionary<string, string> Scheme = new(StringComparer.Ordinal)
    {
        ["accent1"] = "#156082",
        ["tx1"] = "#000000",
        ["bg1"] = "#FFFFFF",
        ["dk2"] = "#0E2841",
    };

    [Theory]
    [InlineData("<a:srgbClr val=\"ff0000\"/>", "#FF0000")]
    [InlineData("<a:sysClr val=\"windowText\" lastClr=\"123456\"/>", "#123456")]
    [InlineData("<a:sysClr val=\"window\"/>", "#FFFFFF")]
    [InlineData("<a:schemeClr val=\"accent1\"/>", "#156082")]
    [InlineData("<a:prstClr val=\"black\"/>", "#000000")]
    [InlineData("<a:scrgbClr r=\"100000\" g=\"0\" b=\"0\"/>", "#FF0000")]
    [InlineData("<a:hslClr hue=\"0\" sat=\"100000\" lum=\"50000\"/>", "#FF0000")]
    // What PowerPoint shows for "Text 1, lighter 25%" and "Background 1, darker 15%".
    [InlineData("<a:schemeClr val=\"tx1\"><a:lumMod val=\"75000\"/><a:lumOff val=\"25000\"/></a:schemeClr>", "#404040")]
    [InlineData("<a:schemeClr val=\"bg1\"><a:lumMod val=\"85000\"/></a:schemeClr>", "#D9D9D9")]
    public void A_colour_is_read_with_its_modifiers(string xml, string expected)
    {
        XElement fill = Parse($"<a:solidFill xmlns:a=\"{Main}\">{xml}</a:solidFill>");

        Assert.True(DrawingColors.TryRead(fill, Scheme, out string? color, out _));
        Assert.Equal(expected, color);
    }

    [Fact]
    public void Alpha_is_a_transparency_and_an_unknown_theme_colour_is_none()
    {
        XElement fill = Parse($"<a:solidFill xmlns:a=\"{Main}\"><a:schemeClr val=\"accent1\"><a:alpha val=\"40000\"/></a:schemeClr></a:solidFill>");

        Assert.True(DrawingColors.TryRead(fill, Scheme, out _, out double transparency));
        Assert.Equal(0.6, transparency, 3);
        Assert.False(DrawingColors.TryRead(Parse($"<a:solidFill xmlns:a=\"{Main}\"><a:schemeClr val=\"accent6\"/></a:solidFill>"), Scheme, out _, out _));
        Assert.False(DrawingColors.TryRead(fill, null, out _, out _));
    }

    [Fact]
    public void A_fill_in_a_theme_colour_is_read_when_the_theme_is_given_and_left_out_when_not()
    {
        XElement properties = Parse($"<p:spPr xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"{Main}\"><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill><a:ln w=\"12700\"><a:solidFill><a:schemeClr val=\"dk2\"/></a:solidFill></a:ln></p:spPr>");

        Assert.False(DrawingReader.TryReadFill(properties, out _));
        Assert.True(DrawingReader.TryReadFill(properties, Scheme, out Fill? fill));
        Assert.Equal("#156082", fill.Color);
        Assert.True(DrawingReader.TryReadOutline(properties, Scheme, out Outline? outline));
        Assert.Equal("#0E2841", outline.Color);
    }

    [Fact]
    public async Task The_synchronous_reader_reads_what_the_asynchronous_one_does()
    {
        XElement shape = Parse($"<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"{Main}\"><p:nvSpPr><p:cNvPr id=\"4\" name=\"Box\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"12700\" y=\"25400\"/><a:ext cx=\"127000\" cy=\"254000\"/></a:xfrm><a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></p:spPr></p:sp>");
        DrawingReadContext sync = new([], _ => Result<byte[]>.NotFound("none")) { SchemeColors = Scheme };

        DrawingReadContext async = new([], (_, _) => Task.FromResult(Result<byte[]>.NotFound("none"))) { SchemeColors = Scheme };

        Assert.True(sync.IsSynchronous);
        Assert.False(async.IsSynchronous);
        Assert.Throws<InvalidOperationException>(() => DrawingReader.Read(shape, async));

        Result<ShapeTreeItem> now = DrawingReader.Read(shape, sync);
        Result<ShapeTreeItem> later = await DrawingReader.ReadAsync(shape, sync);

        Assert.True(now.IsSuccess(out ShapeTreeItem? first));
        Assert.True(later.IsSuccess(out ShapeTreeItem? second));
        Assert.Equal(ShapeKind.Ellipse, first.Shape!.Kind);
        Assert.Equal("#156082", first.Shape.Fill!.Color);
        Assert.Equal(second.Shape!.Fill, first.Shape.Fill);
        Assert.Equal(first.Bounds, second.Bounds);
    }

    [Fact]
    public void A_link_read_from_a_file_keeps_any_target()
    {
        Hyperlink link = Hyperlink.Read("file:///C:/Windows/notepad.exe", null) with { Action = "ppaction://program" };


        Assert.Equal("file:///C:/Windows/notepad.exe", link.Url);
        Assert.Equal("ppaction://program", link.Action);
    }

    [Fact]
    public void Bounds_read_from_a_file_may_start_off_the_edge()
    {
        DrawingBounds bounds = DrawingBounds.Read(-12_700, 25_400, 127_000, -5);

        Assert.Equal(-1, bounds.Left.ToPt(), 6);
        Assert.Equal(0, bounds.Height.Emu);
        Assert.Throws<ArgumentOutOfRangeException>(() => DrawingBounds.FromEmu(-12_700, 0, 1, 1));
    }

    private static XElement Parse(string xml)
    {
        return XElement.Parse(xml);
    }
}
