using System.Text;

namespace Avig.OpenXml.Tests;

/// <summary>Relationships that point out of the package: listed with their part, and hyperlinks told apart from what Office follows on its own.</summary>
public class ExternalRelationshipsTests
{
    private const string Template = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate";
    private const string OleObject = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    private const string Body = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>";

    [Theory]
    [InlineData(OpcPackageKind.Zip)]
    [InlineData(OpcPackageKind.FlatOpc)]
    public async Task Every_external_relationship_is_listed_with_its_part(OpcPackageKind kind)
    {
        using MemoryStream file = new();

        await using (OpcPackageWriter writer = await OpcPackageWriter.CreateAsync(file, kind, leaveOpen: true))
        {
            await writer.WritePartAsync("word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", Encoding.UTF8.GetBytes(Body));
            await writer.WritePartAsync("word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml", "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"u8.ToArray());
            writer.AddRelationship(string.Empty, OpcNamespaces.OfficeDocumentRelationship, "word/document.xml");
            writer.AddRelationship("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settings.xml");
            writer.AddExternalRelationship("word/document.xml", OpcNamespaces.HyperlinkRelationship, "https://example.com/");
            writer.AddExternalRelationship("word/document.xml", OleObject, "file://server/share/sheet.xlsx");
            writer.AddExternalRelationship("word/settings.xml", Template, "https://evil.example/template.dotm");
            writer.AddExternalRelationship(string.Empty, OpcNamespaces.HyperlinkRelationship, "https://example.com/package");
            await writer.CompleteAsync();
        }

        file.Position = 0;
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(file);
        Assert.True(opened.IsSuccess(out OpcPackage? package), opened.Error);

        await using (package)
        {
            IReadOnlyList<OpcExternalRelationship> external = await package.GetExternalRelationshipsAsync();

            Assert.Equal(4, external.Count);
            OpcExternalRelationship template = Assert.Single(external, item => item.Kind == "attachedTemplate");
            Assert.Equal("word/settings.xml", template.SourcePart);
            Assert.Equal("https://evil.example/template.dotm", template.Target);
            Assert.False(template.IsHyperlink);

            OpcExternalRelationship ole = Assert.Single(external, item => item.Kind == "oleObject");
            Assert.Equal("word/document.xml", ole.SourcePart);
            Assert.False(ole.IsHyperlink);

            Assert.Equal(2, external.Count(item => item.IsHyperlink == true));
            Assert.Contains(external, item => item.SourcePart.Length == 0 && item.IsHyperlink == true);
        }
    }

    [Fact]
    public async Task A_package_without_external_relationships_lists_none_and_internal_ones_are_left_out()
    {
        byte[] file = await PackageTests.WriteSampleAsync(OpcPackageKind.Zip);
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file));
        Assert.True(opened.IsSuccess(out OpcPackage? package), opened.Error);

        await using (package)
        {
            OpcExternalRelationship link = Assert.Single(await package.GetExternalRelationshipsAsync());
            Assert.True(link.IsHyperlink);
            Assert.Equal("hyperlink", link.Kind);
        }
    }

    [Theory]
    [InlineData("_rels/.rels", "")]
    [InlineData("word/_rels/document.xml.rels", "word/document.xml")]
    [InlineData("/word/_rels/settings.xml.rels", "word/settings.xml")]
    [InlineData("xl/worksheets/_rels/sheet1.xml.rels", "xl/worksheets/sheet1.xml")]
    [InlineData("_rels/a.xml.rels", "a.xml")]
    public void The_part_a_rels_part_belongs_to_is_found(string relationshipsPath, string expected)
    {
        Assert.True(OpcPath.TryGetSourcePart(relationshipsPath, out string? source));
        Assert.Equal(expected, source);
        Assert.Equal(relationshipsPath.TrimStart('/'), OpcPath.RelationshipsPath(source));
    }

    [Theory]
    [InlineData("word/document.xml")]
    [InlineData("word/document.xml.rels")]
    [InlineData("rels/a.xml.rels")]
    public void A_part_that_is_not_a_rels_part_has_no_source(string name)
    {
        Assert.False(OpcPath.TryGetSourcePart(name, out _));
    }
}
