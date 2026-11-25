using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace Avig.OpenXml.Tests;

/// <summary>
/// Zip files that say two things: the local headers and the central directory disagree, or
/// there are bytes no entry owns, where a scanner and Office would see different documents.
/// The files are written byte by byte with <see cref="RawZip"/>.
/// </summary>
public class ZipStructureTests
{
    private const string Types = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>""";

    [Fact]
    public async Task A_zip_written_byte_by_byte_as_office_writes_it_opens()
    {
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>").Build();

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file));

        Assert.True(opened.IsSuccess(out OpcPackage? package), opened.Error);
        await package.DisposeAsync();
    }

    [Fact]
    public async Task A_zip_with_data_descriptors_opens()
    {
        // .NET writes a data descriptor after each entry when the stream cannot seek, as many other writers do.
        using MemoryStream file = new();

        using (ZipArchive archive = new(new ForwardOnlyStream(file), ZipArchiveMode.Create, true))
        {
            foreach ((string name, string text) in new[] { ("[Content_Types].xml", Types), ("a.xml", "<a/>"), ("b.xml", "<b/>") })
            {
                using Stream stream = archive.CreateEntry(name).Open();
                stream.Write(Encoding.UTF8.GetBytes(text));
            }
        }

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file.ToArray()));

        Assert.True(opened.IsSuccess(out OpcPackage? package), opened.Error);
        await package.DisposeAsync();
    }

    [Theory]
    [InlineData(OpcPackageKind.Zip)]
    public async Task The_sample_package_and_a_zip_with_a_comment_open(OpcPackageKind kind)
    {
        byte[] file = await PackageTests.WriteSampleAsync(kind);
        byte[] commented = new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>").Build(comment: "made by hand");

        Assert.True((await OpcZip.CheckStructureAsync(new MemoryStream(file), 1000, CancellationToken.None)).IsSuccess(out _, out _));
        Assert.True((await OpcZip.CheckStructureAsync(new MemoryStream(commented), 1000, CancellationToken.None)).IsSuccess(out _, out _));
    }

    [Fact]
    public async Task A_local_header_with_another_name_than_the_directory_is_damaged()
    {
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>", localName: "b.xml").Build();

        AssertTwoThings(await OpcPackage.OpenAsync(new MemoryStream(file)), "another name");
    }

    [Fact]
    public async Task A_local_header_with_another_compression_than_the_directory_is_damaged()
    {
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>", localMethod: 8).Build();

        AssertTwoThings(await OpcPackage.OpenAsync(new MemoryStream(file)), "another compression");
    }

    [Fact]
    public async Task An_entry_only_the_local_headers_have_is_damaged()
    {
        // A scanner that walks the local headers sees evil.xml; .NET and Office never do.
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Add("evil.xml", "<evil/>", inDirectory: false).Add("a.xml", "<a/>").Build();

        AssertTwoThings(await OpcPackage.OpenAsync(new MemoryStream(file)), "belong to no entry");
    }

    [Fact]
    public async Task A_hidden_entry_before_the_first_one_is_damaged()
    {
        byte[] file = new RawZip().Add("hidden.xml", "<hidden/>", inDirectory: false).Add("[Content_Types].xml", Types).Add("a.xml", "<a/>").Build();

        AssertTwoThings(await OpcPackage.OpenAsync(new MemoryStream(file)), "belong to no entry");
    }

    [Fact]
    public async Task Bytes_between_two_entries_are_damaged()
    {
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Gap(100).Add("a.xml", "<a/>").Build();

        AssertTwoThings(await OpcPackage.OpenAsync(new MemoryStream(file)), "belong to no entry");
    }

    [Fact]
    public async Task Bytes_between_the_last_entry_and_the_directory_are_damaged()
    {
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>").Gap(64).Build();

        AssertTwoThings(await OpcPackage.OpenAsync(new MemoryStream(file)), "between the last entry");
    }

    [Fact]
    public async Task Two_directory_entries_that_share_their_data_are_damaged()
    {
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>").AliasOfLast("a.xml").Build();

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Bytes_after_the_end_record_are_damaged()
    {
        byte[] file = [.. new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>").Build(), .. "<script>evil</script>"u8.ToArray()];

        AssertTwoThings(await OpcPackage.OpenAsync(new MemoryStream(file)), "after it");
    }

    [Fact]
    public async Task A_second_zip_after_the_first_is_damaged()
    {
        byte[] first = new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>").Build();
        byte[] second = new RawZip().Add("[Content_Types].xml", Types).Add("evil.xml", "<evil/>").Build();

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream([.. first, .. second]));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task An_end_record_that_counts_more_entries_than_the_directory_has_is_damaged()
    {
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>").Build(extraCount: 1);

        AssertTwoThings(await OpcPackage.OpenAsync(new MemoryStream(file)), "fewer entries");
    }

    [Fact]
    public async Task A_truncated_zip_is_damaged()
    {
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>").Build();

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file[..^30]));

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task The_check_can_be_turned_off_for_a_zip_that_net_reads()
    {
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Gap(100).Add("a.xml", "<a/>").Build();

        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(file), new OpcOpenOptions { StrictZip = false });

        Assert.True(opened.IsSuccess(out OpcPackage? package), opened.Error);
        await package.DisposeAsync();
    }

    [Fact]
    public async Task The_check_leaves_the_stream_where_it_was()
    {
        byte[] file = new RawZip().Add("[Content_Types].xml", Types).Add("a.xml", "<a/>").Build();
        using MemoryStream stream = new([.. new byte[10], .. file]);
        stream.Position = 10;

        Assert.True((await OpcZip.CheckStructureAsync(stream, 1000, CancellationToken.None)).IsSuccess(out _, out _));
        Assert.Equal(10, stream.Position);
    }

    private static void AssertTwoThings(Result<OpcPackage> opened, string reason)
    {
        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.StartsWith(OpcErrors.DamagedPrefix, error, StringComparison.Ordinal);
        Assert.Contains("does not say one thing only", error, StringComparison.Ordinal);
        Assert.Contains(reason, error, StringComparison.Ordinal);
    }

    /// <summary>Writes a stored (uncompressed) zip byte by byte, with room to make it lie.</summary>
    private sealed class RawZip
    {
        private readonly List<byte> body = [];
        private readonly List<(string Name, int Method, long Offset, byte[] Data)> directory = [];

        public RawZip Add(string name, string text, string? localName = null, int localMethod = 0, bool inDirectory = true)
        {
            byte[] data = Encoding.UTF8.GetBytes(text);
            long offset = body.Count;
            byte[] local = Encoding.UTF8.GetBytes(localName ?? name);
            Span<byte> header = stackalloc byte[30];
            BinaryPrimitives.WriteUInt32LittleEndian(header, 0x04034b50);
            BinaryPrimitives.WriteUInt16LittleEndian(header[4..], 20);
            BinaryPrimitives.WriteUInt16LittleEndian(header[8..], (ushort)localMethod);
            BinaryPrimitives.WriteUInt32LittleEndian(header[14..], Crc(data));
            BinaryPrimitives.WriteUInt32LittleEndian(header[18..], (uint)data.Length);
            BinaryPrimitives.WriteUInt32LittleEndian(header[22..], (uint)data.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(header[26..], (ushort)local.Length);
            body.AddRange(header.ToArray());
            body.AddRange(local);
            body.AddRange(data);

            if (inDirectory == true)
            {
                directory.Add((name, 0, offset, data));
            }

            return this;
        }

        public RawZip Gap(int length)
        {
            body.AddRange(new byte[length]);

            return this;
        }

        /// <summary>A second directory entry for the last entry's data, under another name.</summary>
        public RawZip AliasOfLast(string name)
        {
            (string _, int method, long offset, byte[] data) = directory[^1];
            directory.Add((name + ".copy", method, offset, data));

            return this;
        }

        public byte[] Build(string comment = "", int extraCount = 0)
        {
            using MemoryStream file = new();
            file.Write(body.ToArray());
            long directoryOffset = file.Position;

            foreach ((string name, int method, long offset, byte[] data) in directory)
            {
                byte[] encoded = Encoding.UTF8.GetBytes(name);
                Span<byte> header = new byte[46];
                BinaryPrimitives.WriteUInt32LittleEndian(header, 0x02014b50);
                BinaryPrimitives.WriteUInt16LittleEndian(header[4..], 20);
                BinaryPrimitives.WriteUInt16LittleEndian(header[6..], 20);
                BinaryPrimitives.WriteUInt16LittleEndian(header[10..], (ushort)method);
                BinaryPrimitives.WriteUInt32LittleEndian(header[16..], Crc(data));
                BinaryPrimitives.WriteUInt32LittleEndian(header[20..], (uint)data.Length);
                BinaryPrimitives.WriteUInt32LittleEndian(header[24..], (uint)data.Length);
                BinaryPrimitives.WriteUInt16LittleEndian(header[28..], (ushort)encoded.Length);
                BinaryPrimitives.WriteUInt32LittleEndian(header[42..], (uint)offset);
                file.Write(header);
                file.Write(encoded);
            }

            long directorySize = file.Position - directoryOffset;
            byte[] note = Encoding.UTF8.GetBytes(comment);
            Span<byte> end = stackalloc byte[22];
            BinaryPrimitives.WriteUInt32LittleEndian(end, 0x06054b50);
            BinaryPrimitives.WriteUInt16LittleEndian(end[8..], (ushort)(directory.Count + extraCount));
            BinaryPrimitives.WriteUInt16LittleEndian(end[10..], (ushort)(directory.Count + extraCount));
            BinaryPrimitives.WriteUInt32LittleEndian(end[12..], (uint)directorySize);
            BinaryPrimitives.WriteUInt32LittleEndian(end[16..], (uint)directoryOffset);
            BinaryPrimitives.WriteUInt16LittleEndian(end[20..], (ushort)note.Length);
            file.Write(end);
            file.Write(note);

            return file.ToArray();
        }

        private static uint Crc(byte[] data)
        {
            uint crc = 0xFFFFFFFF;

            foreach (byte value in data)
            {
                crc ^= value;

                for (int bit = 0; bit < 8; bit++)
                {
                    uint mask = 0u - (crc & 1);
                    crc = (crc >> 1) ^ (0xEDB88320 & mask);
                }
            }

            return ~crc;
        }
    }

    /// <summary>A stream that can only be written forward, so that ZipArchive writes data descriptors.</summary>
    private sealed class ForwardOnlyStream(Stream inner) : Stream
    {
        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => true;

        public override long Length => throw new NotSupportedException();

        public override long Position { get => inner.Position; set => throw new NotSupportedException(); }

        public override void Flush()
        {
            inner.Flush();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
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
            inner.Write(buffer, offset, count);
        }
    }
}
