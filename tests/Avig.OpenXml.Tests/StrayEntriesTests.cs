using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace Avig.OpenXml.Tests;

/// <summary>
/// Entries that do not belong in a package, put into a good one. A name that a reader, an
/// unpacker or a later save would take somewhere else is refused as damaged; folders, which zip
/// programs add when a document is packed again, and ordinary unusual names are not.
/// </summary>
public class StrayEntriesTests
{
    [Theory]
    [InlineData("word/../../evil.xml")]
    [InlineData("..\\evil.xml")]
    [InlineData("word\\..\\..\\evil.xml")]
    [InlineData("//server/share/evil.xml")]
    [InlineData("\\\\server\\share\\evil.xml")]
    [InlineData("word/./document.xml")]
    [InlineData("./evil.xml")]
    [InlineData("word//evil.xml")]
    [InlineData("evil.xml.")]
    [InlineData("evil.xml ")]
    [InlineData("word./evil.xml")]
    [InlineData("word /evil.xml")]
    [InlineData("evil.xml:hidden")]
    [InlineData("word/C:evil.xml")]
    [InlineData("%2e%2e/evil.xml")]
    [InlineData("word%2f..%2f..%2fevil.xml")]
    [InlineData("word%5c..%5cevil.xml")]
    [InlineData("evil%00.xml")]
    [InlineData("CON")]
    [InlineData("word/nul.xml")]
    [InlineData("word/Aux")]
    [InlineData("com1.txt")]
    [InlineData("lpt9")]
    [InlineData("\uFF0E\uFF0E/evil.xml")]
    [InlineData("word\uFF0F\uFF0E\uFF0E\uFF0F\uFF0E\uFF0E\uFF0Fevil.xml")]
    [InlineData("word\uFF3C..\uFF3Cevil.xml")]
    [InlineData("evil\uFF1Ahidden")]
    [InlineData("tab\t.xml")]
    [InlineData("line\n.xml")]
    [InlineData("delete\u007F.xml")]
    public async Task A_stray_entry_with_a_name_no_package_may_have_is_damaged(string name)
    {
        byte[] file = await WithEntriesAsync((name, [1]));

        AssertDamaged(await OpcPackage.OpenAsync(new MemoryStream(file)));
    }

    [Theory]
    [InlineData("word/document.xml", "WORD/DOCUMENT.XML")]
    [InlineData("word/document.xml", "word\\document.xml")]
    [InlineData("word/document.xml", "word/\uFF44ocument.xml")]
    [InlineData("[Content_Types].xml", "[content_types].XML")]
    [InlineData("word/media/image1.png", "word/media/image\uFF11.png")]
    public async Task A_stray_entry_that_is_a_part_already_there_under_another_spelling_is_damaged(string part, string spelling)
    {
        byte[] file = await PackageTests.WriteSampleAsync(OpcPackageKind.Zip);
        Assert.Contains(part, Names(file));

        file = await WithEntriesAsync((spelling, [1]));

        AssertDamaged(await OpcPackage.OpenAsync(new MemoryStream(file)));
    }

    [Fact]
    public async Task Folders_that_a_zip_program_added_are_not_parts()
    {
        byte[] file = await WithEntriesAsync(("word/", []), ("word/media/", []), ("_rels\\", []), ("customXml/", []));

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file));

        Assert.True(opened.IsSuccess(out OpcPackage? package), opened.Error);

        await using (package)
        {
            Assert.DoesNotContain(package.PartNames, name => name.EndsWith('/') == true);
            Assert.True((await package.FindMainPartAsync()).IsSuccess(out _));
        }
    }

    [Theory]
    [InlineData("../")]
    [InlineData("word/../../")]
    [InlineData("C:/")]
    [InlineData("CON/")]
    [InlineData("\uFF0E\uFF0E/")]
    public async Task A_folder_with_a_name_no_package_may_have_is_damaged(string name)
    {
        byte[] file = await WithEntriesAsync((name, []));

        AssertDamaged(await OpcPackage.OpenAsync(new MemoryStream(file)));
    }

    [Fact]
    public async Task A_folder_that_holds_data_is_damaged()
    {
        byte[] file = await WithEntriesAsync(("word/", [1, 2, 3]));

        AssertDamaged(await OpcPackage.OpenAsync(new MemoryStream(file)));
    }

    [Theory]
    [InlineData("word/media/image 2.png")]
    [InlineData("word/media/bild-två.png")]
    [InlineData("word/\uFB01le.xml")]
    [InlineData("customXml/item1.xml")]
    [InlineData("word/a.b.c.xml")]
    [InlineData("word/consoles.xml")]
    [InlineData("word/connect.xml")]
    [InlineData("word/media/image%201.png")]
    [InlineData(".hidden")]
    [InlineData("unknown/no-content-type.bin")]
    public async Task An_unusual_name_a_package_may_have_still_opens(string name)
    {
        byte[] file = await WithEntriesAsync((name, [1, 2, 3]));

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file));

        Assert.True(opened.IsSuccess(out OpcPackage? package), opened.Error);

        await using (package)
        {
            Assert.Contains(name, package.PartNames);
            Assert.True((await package.ReadPartAsync(name)).IsSuccess(out byte[]? data));
            Assert.Equal([1, 2, 3], data);
        }
    }

    [Fact]
    public async Task An_encrypted_entry_is_damaged()
    {
        byte[] file = Patch(StoredZip("word/document.xml", Encoding.UTF8.GetBytes("<w:document/>")), local: 6, central: 8, value: 1);

        AssertDamaged(await OpcPackage.OpenAsync(new MemoryStream(file)));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(12)]
    [InlineData(14)]
    [InlineData(99)]
    public async Task An_entry_with_a_compression_method_dotnet_does_not_have_is_damaged(ushort method)
    {
        byte[] file = Patch(StoredZip("word/document.xml", Encoding.UTF8.GetBytes("<w:document/>")), local: 8, central: 10, value: method);

        AssertDamaged(await OpcPackage.OpenAsync(new MemoryStream(file)));
    }

    [Theory]
    [InlineData("/word/../../evil.xml")]
    [InlineData("/evil.xml:hidden")]
    [InlineData("/%2e%2e/evil.xml")]
    [InlineData("/CON")]
    [InlineData("/\uFF0E\uFF0E/evil.xml")]
    [InlineData("/word//evil.xml")]
    public async Task A_flat_part_with_a_name_no_package_may_have_is_damaged(string name)
    {
        string flat = $"""
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
              <pkg:part pkg:name="{name}" pkg:contentType="application/xml"><pkg:xmlData><a/></pkg:xmlData></pkg:part>
            </pkg:package>
            """;

        AssertDamaged(await OpcPackage.OpenAsync(new MemoryStream(Encoding.UTF8.GetBytes(flat))));
    }

    [Fact]
    public async Task A_flat_part_that_is_there_under_another_spelling_is_damaged()
    {
        string flat = """
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
              <pkg:part pkg:name="/a.xml" pkg:contentType="application/xml"><pkg:xmlData><a/></pkg:xmlData></pkg:part>
              <pkg:part pkg:name="/&#xFF41;.xml" pkg:contentType="application/xml"><pkg:xmlData><a/></pkg:xmlData></pkg:part>
            </pkg:package>
            """;

        AssertDamaged(await OpcPackage.OpenAsync(new MemoryStream(Encoding.UTF8.GetBytes(flat))));
    }

    [Theory]
    [InlineData("word/../../evil.xml")]
    [InlineData("evil.xml:hidden")]
    [InlineData("CON")]
    [InlineData("\uFF0E\uFF0E/evil.xml")]
    public void No_unsafe_name_normalizes(string name)
    {
        Assert.False(OpcPath.TryNormalize(name, out _));
        Assert.False(OpcPath.TryNormalize("/" + name, out _));
    }

    private static void AssertDamaged(Result<OpcPackage> opened)
    {
        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
    }

    /// <summary>The sample package, with entries added after its own.</summary>
    private static async Task<byte[]> WithEntriesAsync(params (string Name, byte[] Data)[] entries)
    {
        using MemoryStream file = new();
        byte[] sample = await PackageTests.WriteSampleAsync(OpcPackageKind.Zip);
        await file.WriteAsync(sample);

        using (ZipArchive archive = new(file, ZipArchiveMode.Update, true))
        {
            foreach ((string name, byte[] data) in entries)
            {
                using Stream stream = archive.CreateEntry(name).Open();
                stream.Write(data);
            }
        }

        return file.ToArray();
    }

    private static List<string> Names(byte[] file)
    {
        using ZipArchive archive = new(new MemoryStream(file));

        return [.. archive.Entries.Select(entry => entry.FullName)];
    }

    private static byte[] StoredZip(string name, byte[] data)
    {
        using MemoryStream file = new();

        using (ZipArchive archive = new(file, ZipArchiveMode.Create, true))
        {
            using Stream stream = archive.CreateEntry(name, CompressionLevel.NoCompression).Open();
            stream.Write(data);
        }

        return file.ToArray();
    }

    /// <summary>
    /// Sets a 16-bit field in every local file header and central directory header: the general
    /// purpose flags (6 and 8; bit 0 is encryption) or the compression method (8 and 10).
    /// </summary>
    private static byte[] Patch(byte[] zip, int local, int central, ushort value)
    {
        for (int position = 0; position + 4 <= zip.Length; position++)
        {
            uint signature = BinaryPrimitives.ReadUInt32LittleEndian(zip.AsSpan(position));

            if (signature == 0x04034b50)
            {
                BinaryPrimitives.WriteUInt16LittleEndian(zip.AsSpan(position + local), value);
            }
            else if (signature == 0x02014b50)
            {
                BinaryPrimitives.WriteUInt16LittleEndian(zip.AsSpan(position + central), value);
            }
        }

        return zip;
    }
}
