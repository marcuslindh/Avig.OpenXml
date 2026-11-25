using System.Buffers.Binary;
using System.Globalization;
using System.Xml;

namespace Avig.OpenXml.Drawing;

/// <summary>The image formats a <see cref="Picture"/> can hold.</summary>
public enum PictureFormat
{
    /// <summary>PNG.</summary>
    Png,

    /// <summary>JPEG.</summary>
    Jpeg,

    /// <summary>GIF.</summary>
    Gif,

    /// <summary>SVG, stored with a PNG, JPEG or GIF that versions of Office before 2016 show instead.</summary>
    Svg,
}

/// <summary>
/// An image to put on a sheet or a slide (Avig.Kalkyl's <c>AddPicture</c>), in a group, or as a
/// <see cref="Fill.Picture(Picture)"/>. PNG, JPEG and GIF are accepted; the format and the size in
/// pixels are read from the file's header, so the image is stored exactly as given. An SVG comes in
/// through <see cref="FromSvg(ReadOnlySpan{byte}, Picture)"/>, with a raster image that older
/// versions of Office show instead.
/// </summary>
public sealed record Picture
{
    private readonly byte[] data;
    private readonly double transparency;

    private Picture(byte[] data, PictureFormat format, int width, int height, Picture? fallback = null)
    {
        this.data = data;
        Format = format;
        Width = width;
        Height = height;
        Fallback = fallback;
    }

    /// <summary>The format.</summary>
    public PictureFormat Format { get; }

    /// <summary>The width in pixels.</summary>
    public int Width { get; }

    /// <summary>The height in pixels.</summary>
    public int Height { get; }

    /// <summary>The alternative text that screen readers read, or null. Set it with <c>picture with { Description = "..." }</c>.</summary>
    public string? Description { get; init; }

    /// <summary>Where a click on the picture goes, or null. Set it with <c>picture with { Link = ... }</c>.</summary>
    public Hyperlink? Link { get; init; }

    /// <summary>How see-through the picture is, from 0 (opaque) to 1 (invisible).</summary>
    public double Transparency
    {
        get => transparency;
        init => transparency = DrawingValues.Transparency(value, nameof(Transparency));
    }

    /// <summary>A shadow behind the picture, or null for none.</summary>
    public Shadow? Shadow { get; init; }

    /// <summary>The image file's bytes, as they are stored in the package.</summary>
    public ReadOnlyMemory<byte> Data => data;

    /// <summary>The raster image an SVG is stored with; null for the other formats.</summary>
    public Picture? Fallback { get; }

    /// <summary>The array behind the data, which identifies the image when it is put on a sheet more than once.</summary>
    internal byte[] Identity => data;

    /// <summary>The extension of the image part: png, jpeg, gif or svg.</summary>
    public string Extension => Format switch
    {
        PictureFormat.Jpeg => "jpeg",
        PictureFormat.Gif => "gif",
        PictureFormat.Svg => "svg",
        _ => "png",
    };

    /// <summary>The content type of the image part, such as image/png.</summary>
    public string ContentType => Format switch
    {
        PictureFormat.Svg => "image/svg+xml",
        _ => "image/" + Extension,
    };

    /// <summary>A fully transparent 1×1 PNG: the fallback of an SVG given without one.</summary>
    private static readonly byte[] TransparentPixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNgYGBgAAAABQABeqhXUAAAAABJRU5ErkJggg==");

    /// <summary>Creates a picture from the bytes of an image file. The bytes are copied.</summary>
    /// <param name="data">The file's bytes.</param>
    /// <returns>The picture, or Invalid when the bytes are not a PNG, JPEG or GIF whose size can be read.</returns>
    public static Result<Picture> FromBytes(ReadOnlySpan<byte> data)
    {
        if (TryReadHeader(data, out PictureFormat format, out int width, out int height) == false)
        {
            return Result<Picture>.Invalid("The image is not a PNG, JPEG or GIF file, or its header is damaged.");
        }

        return new Picture(data.ToArray(), format, width, height);
    }

    /// <summary>Reads a picture from an image file.</summary>
    /// <param name="path">The file.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    /// <returns>The picture, NotFound when there is no such file, or Invalid when it is not a PNG, JPEG or GIF.</returns>
    public static async Task<Result<Picture>> FromFileAsync(string path, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        byte[] bytes;

        try
        {
            bytes = await File.ReadAllBytesAsync(path, cancellationToken);
        }
        catch (Exception exception) when (exception is FileNotFoundException or DirectoryNotFoundException)
        {
            return Result<Picture>.NotFound($"'{path}' does not exist.");
        }
        catch (UnauthorizedAccessException)
        {
            return Result<Picture>.Forbidden($"'{path}' may not be read.");
        }
        catch (IOException exception)
        {
            return Result<Picture>.Conflict($"'{path}' could not be read: {exception.Message}");
        }

        if (TryReadHeader(bytes, out PictureFormat format, out int width, out int height) == false)
        {
            return Result<Picture>.Invalid($"'{path}' is not a PNG, JPEG or GIF file, or its header is damaged.");
        }

        return new Picture(bytes, format, width, height);
    }

    /// <summary>
    /// Creates a picture from an SVG file's bytes, with a transparent 1×1 PNG as the image older
    /// versions of Excel show instead. Excel 2016 and later, and Microsoft 365, draw the SVG.
    /// </summary>
    /// <param name="svg">The SVG file's bytes. They are copied.</param>
    /// <returns>The picture, or Invalid when the bytes are not an SVG document.</returns>
    public static Result<Picture> FromSvg(ReadOnlySpan<byte> svg)
    {
        Picture pixel = new(TransparentPixel, PictureFormat.Png, 1, 1);

        return FromSvg(svg, pixel);
    }

    /// <summary>
    /// Creates a picture from an SVG file's bytes and the PNG, JPEG or GIF that versions of Excel
    /// before 2016 show instead. The size is the SVG's own (its width and height, or its viewBox),
    /// or the fallback's when the SVG has none.
    /// </summary>
    /// <param name="svg">The SVG file's bytes. They are copied.</param>
    /// <param name="fallback">The raster image, not itself an SVG.</param>
    /// <returns>The picture, or Invalid when the bytes are not an SVG document.</returns>
    public static Result<Picture> FromSvg(ReadOnlySpan<byte> svg, Picture fallback)
    {
        ArgumentNullException.ThrowIfNull(fallback);

        if (fallback.Format == PictureFormat.Svg)
        {
            throw new ArgumentException("The fallback of an SVG must be a PNG, JPEG or GIF.", nameof(fallback));
        }

        byte[] bytes = svg.ToArray();

        if (TryReadSvgSize(bytes, out double width, out double height) == false)
        {
            return Result<Picture>.Invalid("The bytes are not an SVG document: the first element must be <svg>.");
        }

        int pixelsWide = fallback.Width;
        int pixelsHigh = fallback.Height;

        if (width > 0 && height > 0)
        {
            pixelsWide = (int)Math.Clamp(Math.Round(width), 1, 100_000);
            pixelsHigh = (int)Math.Clamp(Math.Round(height), 1, 100_000);
        }
        else if (fallback.Width == 1 && fallback.Height == 1)
        {
            // Neither says how large it is: the size Excel gives an inserted icon.
            pixelsWide = 96;
            pixelsHigh = 96;
        }

        return new Picture(bytes, PictureFormat.Svg, pixelsWide, pixelsHigh, fallback with { Link = null, Description = null, Shadow = null, Transparency = 0 });
    }

    /// <summary>
    /// True when the first element is &lt;svg&gt;, with its size in pixels from width and height
    /// (px, pt, in, cm and mm are converted; percentages are no size) or else the viewBox. The
    /// size is 0 when neither gives one. A DTD is skipped, never processed.
    /// </summary>
    private static bool TryReadSvgSize(byte[] bytes, out double width, out double height)
    {
        width = 0;
        height = 0;
        XmlReaderSettings settings = new()
        {
            DtdProcessing = DtdProcessing.Ignore,
            XmlResolver = null,
            MaxCharactersFromEntities = 1024,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
        };

        try
        {
            using MemoryStream stream = new(bytes, false);
            using XmlReader reader = XmlReader.Create(stream, settings);

            if (reader.MoveToContent() != XmlNodeType.Element || reader.LocalName != "svg")
            {
                return false;
            }

            width = SvgLength(reader.GetAttribute("width"));
            height = SvgLength(reader.GetAttribute("height"));

            if ((width <= 0 || height <= 0) && reader.GetAttribute("viewBox") is string viewBox)
            {
                string[] parts = viewBox.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length == 4
                    && double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double boxWidth) == true
                    && double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out double boxHeight) == true
                    && double.IsFinite(boxWidth) == true && double.IsFinite(boxHeight) == true && boxWidth > 0 && boxHeight > 0)
                {
                    width = boxWidth;
                    height = boxHeight;
                }
            }

            return true;
        }
        catch (XmlException)
        {
            return false;
        }
    }

    /// <summary>An SVG length in pixels at 96 to the inch, or 0 when it is missing or relative.</summary>
    private static double SvgLength(string? text)
    {
        if (text is null)
        {
            return 0;
        }

        string trimmed = text.Trim();
        (string Unit, double Pixels)[] units = [("px", 1), ("pt", 96.0 / 72), ("in", 96), ("cm", 96 / 2.54), ("mm", 96 / 25.4)];
        double factor = 1;

        foreach ((string unit, double pixels) in units)
        {
            if (trimmed.EndsWith(unit, StringComparison.OrdinalIgnoreCase) == true)
            {
                trimmed = trimmed[..^unit.Length];
                factor = pixels;

                break;
            }
        }

        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) == false || double.IsFinite(value) == false || value <= 0)
        {
            return 0;
        }

        return value * factor;
    }

    /// <summary>
    /// Finds the format and size. Every read is checked against the length, since the bytes may
    /// come from anyone: a damaged header gives false, never an exception or a long loop.
    /// </summary>
    internal static bool TryReadHeader(ReadOnlySpan<byte> data, out PictureFormat format, out int width, out int height)
    {
        format = PictureFormat.Png;
        width = 0;
        height = 0;
        bool found = false;

        if (data.Length >= 24 && data[..8].SequenceEqual((ReadOnlySpan<byte>)[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]) == true
            && data.Slice(12, 4).SequenceEqual("IHDR"u8) == true)
        {
            format = PictureFormat.Png;
            width = (int)Math.Min(BinaryPrimitives.ReadUInt32BigEndian(data.Slice(16, 4)), int.MaxValue);
            height = (int)Math.Min(BinaryPrimitives.ReadUInt32BigEndian(data.Slice(20, 4)), int.MaxValue);
            found = true;
        }
        else if (data.Length >= 10 && (data[..6].SequenceEqual("GIF87a"u8) == true || data[..6].SequenceEqual("GIF89a"u8) == true))
        {
            format = PictureFormat.Gif;
            width = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(6, 2));
            height = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(8, 2));
            found = true;
        }
        else if (data.Length >= 4 && data[0] == 0xFF && data[1] == 0xD8)
        {
            format = PictureFormat.Jpeg;
            found = TryReadJpegSize(data, out width, out height);
        }

        return found == true && width > 0 && height > 0;
    }

    /// <summary>Walks the JPEG segments to the first frame header (SOF), which holds the size.</summary>
    private static bool TryReadJpegSize(ReadOnlySpan<byte> data, out int width, out int height)
    {
        width = 0;
        height = 0;
        int position = 2;

        // Each step moves forward by at least two bytes, so the loop ends.
        while (position + 4 <= data.Length)
        {
            if (data[position] != 0xFF)
            {
                return false;
            }

            byte marker = data[position + 1];

            // Fill bytes between segments.
            if (marker == 0xFF)
            {
                position++;

                continue;
            }

            // Markers without a length: TEM and RST0 to RST7.
            if (marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7))
            {
                position += 2;

                continue;
            }

            // The end of the image, or the start of the compressed data, before any frame header.
            if (marker == 0xD9 || marker == 0xDA)
            {
                return false;
            }

            int length = BinaryPrimitives.ReadUInt16BigEndian(data.Slice(position + 2, 2));

            if (length < 2)
            {
                return false;
            }

            // SOF0 to SOF15, except DHT (C4), JPG (C8) and DAC (CC), which share the range.
            bool frame = marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;

            if (frame == true)
            {
                if (position + 9 > data.Length)
                {
                    return false;
                }

                height = BinaryPrimitives.ReadUInt16BigEndian(data.Slice(position + 5, 2));
                width = BinaryPrimitives.ReadUInt16BigEndian(data.Slice(position + 7, 2));

                return true;
            }

            position += 2 + length;
        }

        return false;
    }
}
