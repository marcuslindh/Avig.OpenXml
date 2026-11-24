using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Avig.OpenXml.Tests;

/// <summary>Files made to hurt a reader. Every one must be refused with a message that names the kind of failure.</summary>
public class HostileFilesTests
{
    [Fact]
    public async Task A_file_that_is_neither_zip_nor_xml_is_not_a_package()
    {
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream("hello world"u8.ToArray()), new OpcOpenOptions { FormatName = "a .docx document" });

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith("The file is not a .docx document", error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_compound_file_is_not_a_package_and_says_why()
    {
        byte[] ole = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0];

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(ole));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith("The file is not an Open XML package", error, StringComparison.Ordinal);
        Assert.Contains("old binary format", error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task An_xml_file_that_is_not_flat_opc_is_not_a_package()
    {
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream("<root/>"u8.ToArray()));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith("The file is not an Open XML package", error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Flat_opc_is_refused_when_it_is_not_allowed()
    {
        byte[] flat = await PackageTests.WriteSampleAsync(OpcPackageKind.FlatOpc);

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(flat), new OpcOpenOptions { AllowFlatOpc = false });

        Assert.True(opened.IsFailure(out _, out _));
    }

    [Fact]
    public async Task A_zip_that_unpacks_past_the_limit_is_too_large()
    {
        byte[] zip = Zip(("a.xml", new byte[10_000]));

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(zip), new OpcOpenOptions { MaxUncompressedBytes = 1_000, LimitsName = "OrdLimits" });

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.Contains("the limit set in OrdLimits.MaxUncompressedBytes", error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_flat_file_past_the_limit_is_too_large()
    {
        byte[] flat = await PackageTests.WriteSampleAsync(OpcPackageKind.FlatOpc);

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(flat), new OpcOpenOptions { MaxUncompressedBytes = 100 });

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.Contains(OpcErrors.LimitMarker, error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Too_many_parts_is_too_large()
    {
        byte[] zip = Zip(("a.xml", [1]), ("b.xml", [1]), ("c.xml", [1]));

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(zip), new OpcOpenOptions { MaxParts = 2 });

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.Contains("the limit set in OpcOpenOptions.MaxParts", error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("../evil.xml")]
    [InlineData("/etc/passwd")]
    [InlineData("C:/evil.xml")]
    [InlineData("a\u0001.xml")]
    public async Task A_part_name_that_points_outside_the_package_is_damaged(string name)
    {
        byte[] zip = Zip((name, [1]));

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(zip));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_part_that_is_in_the_zip_twice_is_damaged()
    {
        byte[] zip = Zip(("a.xml", [1]), ("A.xml", [2]));

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(zip));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.Contains("is in it twice", error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_flat_part_that_is_in_it_twice_is_damaged()
    {
        string flat = """
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
              <pkg:part pkg:name="/a.xml" pkg:contentType="application/xml"><pkg:xmlData><a/></pkg:xmlData></pkg:part>
              <pkg:part pkg:name="/A.xml" pkg:contentType="application/xml"><pkg:xmlData><a/></pkg:xmlData></pkg:part>
            </pkg:package>
            """;

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(Encoding.UTF8.GetBytes(flat)));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_flat_part_with_a_name_outside_the_package_is_damaged()
    {
        string flat = """
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
              <pkg:part pkg:name="/../a.xml" pkg:contentType="application/xml"><pkg:xmlData><a/></pkg:xmlData></pkg:part>
            </pkg:package>
            """;

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(Encoding.UTF8.GetBytes(flat)));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_dtd_is_refused()
    {
        string flat = """
            <?xml version="1.0"?>
            <!DOCTYPE lolz [<!ENTITY lol "lol">]>
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"/>
            """;

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(Encoding.UTF8.GetBytes(flat)));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Nesting_past_the_depth_limit_is_broken_xml()
    {
        string deep = string.Concat(Enumerable.Repeat("<a>", 50)) + string.Concat(Enumerable.Repeat("</a>", 50));
        byte[] zip = Zip(("deep.xml", Encoding.UTF8.GetBytes(deep)));
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(zip), new OpcOpenOptions { MaxXmlDepth = 10 });
        Assert.True(opened.IsSuccess(out OpcPackage? package));

        await using (package)
        {
            Assert.True((await package.OpenPartAsync("deep.xml")).IsSuccess(out Stream? part));
            using XmlReader reader = package.CreateXmlReader(part);

            XmlException exception = await Assert.ThrowsAsync<XmlException>(async () =>
            {
                while (await reader.ReadAsync() == true)
                {
                }
            });

            Assert.Contains("more than 10 levels", exception.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task A_part_that_unpacks_past_its_header_is_damaged()
    {
        await using Stream part = OpcZip.LimitStream(new MemoryStream(new byte[100]), 50, "word/document.xml");

        InvalidDataException exception = await Assert.ThrowsAsync<InvalidDataException>(async () => await part.CopyToAsync(Stream.Null));

        Assert.StartsWith(OpcErrors.DamagedPrefix, exception.Message, StringComparison.Ordinal);
        Assert.Contains("word/document.xml", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_broken_zip_is_damaged()
    {
        byte[] zip = Zip(("a.xml", new byte[1000]));
        byte[] broken = zip[..(zip.Length - 30)];

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(broken));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_stream_that_cannot_seek_and_never_ends_is_stopped_at_the_limit()
    {
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new EndlessStream(), new OpcOpenOptions { MaxUncompressedBytes = 1_000_000 });

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.Contains(OpcErrors.LimitMarker, error, StringComparison.Ordinal);
    }

    private static byte[] Zip(params (string Name, byte[] Data)[] entries)
    {
        using MemoryStream file = new();

        using (ZipArchive archive = new(file, ZipArchiveMode.Create, true))
        {
            foreach ((string name, byte[] data) in entries)
            {
                using Stream stream = archive.CreateEntry(name).Open();
                stream.Write(data);
            }
        }

        return file.ToArray();
    }

    private sealed class EndlessStream : Stream
    {
        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            Array.Fill(buffer, (byte)'P', offset, count);

            return count;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }
    }
}
