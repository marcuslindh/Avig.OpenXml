namespace Avig.OpenXml.Tests;

public class PathTests
{
    [Theory]
    [InlineData("word/document.xml", "media/image1.png", "word/media/image1.png")]
    [InlineData("word/document.xml", "../customXml/item1.xml", "customXml/item1.xml")]
    [InlineData("word/document.xml", "/word/styles.xml", "word/styles.xml")]
    [InlineData("", "word/document.xml", "word/document.xml")]
    [InlineData("word/document.xml", "../../../x.xml", "x.xml")]
    [InlineData("word/document.xml", ".\\media\\a.png", "word/media/a.png")]
    public void A_target_resolves_against_its_source(string source, string target, string expected)
    {
        Assert.Equal(expected, OpcPath.ResolveTarget(source, target));
    }

    [Theory]
    [InlineData("word/document.xml", "word/media/image1.png", "media/image1.png")]
    [InlineData("word/document.xml", "customXml/item1.xml", "../customXml/item1.xml")]
    [InlineData("word/document.xml", "word/styles.xml", "styles.xml")]
    public void A_relative_target_resolves_back_to_its_part(string source, string target, string expected)
    {
        string relative = OpcPath.RelativeTarget(source, target);

        Assert.Equal(expected, relative);
        Assert.Equal(target, OpcPath.ResolveTarget(source, relative));
    }

    [Theory]
    [InlineData("", "_rels/.rels")]
    [InlineData("word/document.xml", "word/_rels/document.xml.rels")]
    [InlineData("a.xml", "_rels/a.xml.rels")]
    public void The_rels_part_sits_beside_its_part(string part, string expected)
    {
        Assert.Equal(expected, OpcPath.RelationshipsPath(part));
    }

    [Fact]
    public void Relationship_types_compare_on_the_last_segment_so_strict_ooxml_matches()
    {
        Assert.True(OpcPath.IsRelationship("http://purl.oclc.org/ooxml/officeDocument/relationships/styles", "styles"));
        Assert.False(OpcPath.IsRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/stylesWithEffects", "styles"));
    }
}
