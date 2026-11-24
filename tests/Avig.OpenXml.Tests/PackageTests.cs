using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Avig.OpenXml.Tests;

public class PackageTests
{
    private const string DocumentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string Body = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Hej</w:t></w:r></w:p></w:body></w:document>";

    [Theory]
    [InlineData(OpcPackageKind.Zip)]
    [InlineData(OpcPackageKind.FlatOpc)]
    public async Task A_written_package_reads_back_with_its_parts_types_and_relationships(OpcPackageKind kind)
    {
        byte[] file = await WriteSampleAsync(kind);

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file));

        Assert.True(opened.IsSuccess(out OpcPackage? package));

        await using (package)
        {
            Assert.Equal(kind, package.Kind);
            Assert.True((await package.FindMainPartAsync()).IsSuccess(out string? main));
            Assert.Equal("word/document.xml", main);
            Assert.True(package.TryGetContentType("word/document.xml", out string? type));
            Assert.Equal(DocumentType, type);
            Assert.True(package.TryGetContentType("/word/media/image1.png", out string? image));
            Assert.Equal("image/png", image);

            Assert.True((await package.ReadPartAsync("word/media/image1.png")).IsSuccess(out byte[]? png));
            Assert.Equal([1, 2, 3, 4], png);

            Assert.True((await package.FindRelatedPartAsync("word/document.xml", "image")).IsSuccess(out string? imagePart));
            Assert.Equal("word/media/image1.png", imagePart);

            IReadOnlyList<OpcRelationship> relationships = await package.GetRelationshipsAsync("word/document.xml");
            OpcRelationship link = Assert.Single(relationships, relationship => relationship.IsExternal == true);
            Assert.Equal("https://example.com/", link.Target);

            Assert.True((await package.ReadPartAsync("word/document.xml")).IsSuccess(out byte[]? xml));
            Assert.Contains("Hej", Encoding.UTF8.GetString(xml), StringComparison.Ordinal);
            Assert.False(package.HasMacros);
        }
    }

    [Fact]
    public async Task A_zip_package_gets_defaults_for_images_and_overrides_for_xml_parts()
    {
        byte[] file = await WriteSampleAsync(OpcPackageKind.Zip);
        using ZipArchive archive = new(new MemoryStream(file));
        using StreamReader reader = new(archive.GetEntry("[Content_Types].xml")!.Open());
        string types = await reader.ReadToEndAsync();

        Assert.Contains("<Default Extension=\"png\" ContentType=\"image/png\"", types, StringComparison.Ordinal);
        Assert.Contains("<Override PartName=\"/word/document.xml\"", types, StringComparison.Ordinal);
        Assert.DoesNotContain("PartName=\"/word/media/image1.png\"", types, StringComparison.Ordinal);
        Assert.NotNull(archive.GetEntry("word/_rels/document.xml.rels"));
        Assert.NotNull(archive.GetEntry("_rels/.rels"));
    }

    [Fact]
    public async Task A_flat_opc_part_that_relies_on_a_namespace_declared_on_the_package_still_parses()
    {
        string flat = """
            <?xml version="1.0" encoding="UTF-8"?>
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/xml">
                <pkg:xmlData><w:document><w:body><w:p/></w:body></w:document></pkg:xmlData>
              </pkg:part>
              <pkg:part pkg:name="/word/media/a.bin" pkg:contentType="application/octet-stream">
                <pkg:binaryData>AQID
            BA==</pkg:binaryData>
              </pkg:part>
            </pkg:package>
            """;

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(Encoding.UTF8.GetBytes(flat)));

        Assert.True(opened.IsSuccess(out OpcPackage? package));

        await using (package)
        {
            Assert.True((await package.OpenPartAsync("word/document.xml")).IsSuccess(out Stream? part));
            using XmlReader reader = package.CreateXmlReader(part);
            await reader.MoveToContentAsync();
            Assert.Equal("http://schemas.openxmlformats.org/wordprocessingml/2006/main", reader.NamespaceURI);
            Assert.True((await package.ReadPartAsync("word/media/a.bin")).IsSuccess(out byte[]? data));
            Assert.Equal([1, 2, 3, 4], data);
        }
    }

    [Fact]
    public async Task A_missing_part_is_not_found()
    {
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(await WriteSampleAsync(OpcPackageKind.Zip)));
        Assert.True(opened.IsSuccess(out OpcPackage? package));

        await using (package)
        {
            Result<Stream> part = await package.OpenPartAsync("word/nothing.xml");
            Assert.True(part.IsFailure(out _, out ResultStatus status));
            Assert.Equal(ResultStatus.NotFound, status);
            Assert.Empty(await package.GetRelationshipsAsync("word/nothing.xml"));
        }
    }

    [Fact]
    public async Task A_package_with_a_vba_project_has_macros()
    {
        using MemoryStream file = new();

        await using (OpcPackageWriter writer = await OpcPackageWriter.CreateAsync(file, leaveOpen: true))
        {
            await writer.WritePartAsync("word/vbaProject.bin", OpcNamespaces.VbaProjectContentType, new byte[] { 0 });
            await writer.CompleteAsync();
        }

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file.ToArray()));
        Assert.True(opened.IsSuccess(out OpcPackage? package));

        await using (package)
        {
            Assert.True(package.HasMacros);
        }
    }

    [Fact]
    public async Task The_writer_refuses_a_part_twice_and_content_types_by_hand()
    {
        await using OpcPackageWriter writer = await OpcPackageWriter.CreateAsync(new MemoryStream());
        await writer.WritePartAsync("a.xml", "application/xml", "<a/>"u8.ToArray());

        await Assert.ThrowsAsync<InvalidOperationException>(() => writer.WritePartAsync("A.xml", "application/xml", "<a/>"u8.ToArray()));
        await Assert.ThrowsAsync<InvalidOperationException>(() => writer.WritePartAsync("[Content_Types].xml", "application/xml", "<a/>"u8.ToArray()));
        await Assert.ThrowsAsync<ArgumentException>(() => writer.WritePartAsync("../evil.xml", "application/xml", "<a/>"u8.ToArray()));
    }

    [Fact]
    public async Task A_copied_part_keeps_its_content_type_and_relationships()
    {
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(await WriteSampleAsync(OpcPackageKind.Zip)));
        Assert.True(opened.IsSuccess(out OpcPackage? source));
        using MemoryStream copy = new();

        await using (source)
        {
            await using OpcPackageWriter writer = await OpcPackageWriter.CreateAsync(copy, leaveOpen: true);

            foreach (string name in source.PartNames.Where(name => name.EndsWith(".rels", StringComparison.Ordinal) == false && name != "[Content_Types].xml"))
            {
                await writer.CopyPartAsync(source, name);
            }

            writer.AddRelationship(string.Empty, OpcNamespaces.OfficeDocumentRelationship, "word/document.xml");
            await writer.CompleteAsync();
        }

        Result<OpcPackage> again = await OpcPackage.OpenAsync(new MemoryStream(copy.ToArray()));
        Assert.True(again.IsSuccess(out OpcPackage? package));

        await using (package)
        {
            Assert.True(package.TryGetContentType("word/document.xml", out string? type));
            Assert.Equal(DocumentType, type);
            Assert.Equal(2, (await package.GetRelationshipsAsync("word/document.xml")).Count);
        }
    }

    [Theory]
    [InlineData(OpcPackageKind.Zip)]
    [InlineData(OpcPackageKind.FlatOpc)]
    public async Task A_part_written_bit_by_bit_is_done_when_its_stream_is_disposed(OpcPackageKind kind)
    {
        using MemoryStream file = new();

        await using (OpcPackageWriter writer = await OpcPackageWriter.CreateAsync(file, kind, leaveOpen: true))
        {
            await using (Stream part = await writer.BeginPartAsync("a.xml", "application/xml"))
            {
                await part.WriteAsync("<a>"u8.ToArray());
                await Assert.ThrowsAsync<InvalidOperationException>(() => writer.BeginPartAsync("b.xml", "application/xml"));
                await part.WriteAsync("</a>"u8.ToArray());
            }

            await writer.WritePartAsync("b.xml", "application/xml", "<b/>"u8.ToArray());
            await writer.CompleteAsync();
        }

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file.ToArray()));
        Assert.True(opened.IsSuccess(out OpcPackage? package));

        await using (package)
        {
            Assert.True((await package.ReadPartAsync("a.xml")).IsSuccess(out byte[]? data));
            Assert.Contains("<a", Encoding.UTF8.GetString(data), StringComparison.Ordinal);
            Assert.True(package.Contains("b.xml"));
        }
    }

    internal static async Task<byte[]> WriteSampleAsync(OpcPackageKind kind)
    {
        using MemoryStream file = new();

        await using (OpcPackageWriter writer = await OpcPackageWriter.CreateAsync(file, kind, leaveOpen: true))
        {
            await writer.WritePartAsync("word/document.xml", DocumentType, Encoding.UTF8.GetBytes(Body));
            await writer.WritePartAsync("word/media/image1.png", "image/png", new byte[] { 1, 2, 3, 4 });
            writer.AddRelationship(string.Empty, OpcNamespaces.OfficeDocumentRelationship, "word/document.xml");
            Assert.Equal("rId1", writer.AddRelationship("word/document.xml", OpcNamespaces.ImageRelationship, "word/media/image1.png"));
            Assert.Equal("rId2", writer.AddExternalRelationship("word/document.xml", OpcNamespaces.HyperlinkRelationship, "https://example.com/"));
            await writer.CompleteAsync();
        }

        return file.ToArray();
    }
}
