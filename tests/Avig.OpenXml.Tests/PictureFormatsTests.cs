using System.Buffers.Binary;
using Avig.OpenXml.Drawing;

namespace Avig.OpenXml.Tests;

/// <summary>BMP and TIFF pictures: the size is read from the header, and a damaged header is refused.</summary>
public class PictureFormatsTests
{
    [Fact]
    public void A_bmp_gives_its_size_and_is_stored_as_bmp()
    {
        Assert.True(Picture.FromBytes(Bmp(40, 320, -200)).IsSuccess(out Picture? topDown));
        Assert.Equal((PictureFormat.Bmp, 320, 200, "bmp", "image/bmp"), (topDown.Format, topDown.Width, topDown.Height, topDown.Extension, topDown.ContentType));
        Assert.True(Picture.FromBytes(Bmp(124, 16, 9)).IsSuccess(out Picture? v5));
        Assert.Equal((16, 9), (v5.Width, v5.Height));

        byte[] core = new byte[26];
        core[0] = (byte)'B';
        core[1] = (byte)'M';
        BinaryPrimitives.WriteUInt32LittleEndian(core.AsSpan(14), 12);
        BinaryPrimitives.WriteUInt16LittleEndian(core.AsSpan(18), 7);
        BinaryPrimitives.WriteUInt16LittleEndian(core.AsSpan(20), 5);
        Assert.True(Picture.FromBytes(core).IsSuccess(out Picture? old));
        Assert.Equal((7, 5), (old.Width, old.Height));
    }

    [Fact]
    public void A_tiff_gives_its_size_in_either_byte_order()
    {
        Assert.True(Picture.FromBytes(Tiff(true, 640, 480, false)).IsSuccess(out Picture? little));
        Assert.Equal((PictureFormat.Tiff, 640, 480, "tiff", "image/tiff"), (little.Format, little.Width, little.Height, little.Extension, little.ContentType));
        Assert.True(Picture.FromBytes(Tiff(false, 100_000, 3, true)).IsSuccess(out Picture? big));
        Assert.Equal((100_000, 3), (big.Width, big.Height));
    }

    [Fact]
    public void Damaged_bmp_and_tiff_headers_are_refused()
    {
        Assert.False(Picture.FromBytes(Bmp(40, 0, 10)).IsSuccess(out _));
        Assert.False(Picture.FromBytes(Bmp(20, 10, 10)).IsSuccess(out _));
        Assert.False(Picture.FromBytes("BM"u8).IsSuccess(out _));

        byte[] tiff = Tiff(true, 10, 10, false);
        BinaryPrimitives.WriteUInt32LittleEndian(tiff.AsSpan(4), 0xFFFF_FFF0);
        Assert.False(Picture.FromBytes(tiff).IsSuccess(out _));

        byte[] cut = Tiff(true, 10, 10, false)[..14];
        Assert.False(Picture.FromBytes(cut).IsSuccess(out _));

        byte[] many = Tiff(true, 10, 10, false);
        BinaryPrimitives.WriteUInt16LittleEndian(many.AsSpan(8), 65_535);
        Assert.False(Picture.FromBytes(many).IsSuccess(out _));
    }

    private static byte[] Bmp(uint header, int width, int height)
    {
        byte[] data = new byte[14 + Math.Max(header, 12)];
        data[0] = (byte)'B';
        data[1] = (byte)'M';
        BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(14), header);
        BinaryPrimitives.WriteInt32LittleEndian(data.AsSpan(18), width);
        BinaryPrimitives.WriteInt32LittleEndian(data.AsSpan(22), height);

        return data;
    }

    private static byte[] Tiff(bool little, uint width, uint height, bool longValues)
    {
        byte[] data = new byte[8 + 2 + (2 * 12) + 4];

        if (little == true)
        {
            data[0] = 0x49;
            data[1] = 0x49;
            data[2] = 0x2A;
        }
        else
        {
            data[0] = 0x4D;
            data[1] = 0x4D;
            data[3] = 0x2A;
        }

        Write32(data, 4, 8, little);
        Write16(data, 8, 2, little);
        Entry(data, 10, 256, width, little, longValues);
        Entry(data, 22, 257, height, little, longValues);

        return data;
    }

    private static void Entry(byte[] data, int position, ushort tag, uint value, bool little, bool longValue)
    {
        Write16(data, position, tag, little);
        Write16(data, position + 2, TypeOf(longValue), little);
        Write32(data, position + 4, 1, little);

        if (longValue == true)
        {
            Write32(data, position + 8, value, little);
        }
        else
        {
            Write16(data, position + 8, (ushort)value, little);
        }
    }

    private static ushort TypeOf(bool longValue)
    {
        if (longValue == true)
        {
            return 4;
        }

        return 3;
    }

    private static void Write16(byte[] data, int position, ushort value, bool little)
    {
        if (little == true)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(position), value);
        }
        else
        {
            BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(position), value);
        }
    }

    private static void Write32(byte[] data, int position, uint value, bool little)
    {
        if (little == true)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(position), value);
        }
        else
        {
            BinaryPrimitives.WriteUInt32BigEndian(data.AsSpan(position), value);
        }
    }
}
