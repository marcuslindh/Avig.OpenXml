namespace Avig.OpenXml.Drawing;

/// <summary>The shapes a gradient spreads in (a:lin, a:path path), as on Office's Type menu for a gradient.</summary>
public enum GradientShape
{
    /// <summary>Along a line at <see cref="Fill.Angle"/>.</summary>
    Linear,

    /// <summary>In circles from the focus (Office's Radial).</summary>
    Circle,

    /// <summary>In rectangles from the focus (Office's Rectangular).</summary>
    Rectangle,

    /// <summary>Following the shape's outline (Office's Path).</summary>
    Shape,
}

/// <summary>
/// A rectangle inside the shape, as fractions of its size inset from each side (a:fillToRect,
/// a:tileRect): (0.5, 0.5, 0.5, 0.5) is the centre, (0, 0, 1, 1) the top-left corner. Values may be
/// negative or above 1, reaching outside the shape.
/// </summary>
/// <param name="Left">Inset from the left.</param>
/// <param name="Top">Inset from the top.</param>
/// <param name="Right">Inset from the right.</param>
/// <param name="Bottom">Inset from the bottom.</param>
public sealed record FillRectangle(double Left, double Top, double Right, double Bottom)
{
    /// <summary>The centre: a radial gradient from the middle.</summary>
    public static FillRectangle Center { get; } = new(0.5, 0.5, 0.5, 0.5);

    /// <summary>Inset from the left.</summary>
    public double Left { get; init; } = DrawingValues.Range(Left, -100, 100, nameof(Left), "An inset");

    /// <summary>Inset from the top.</summary>
    public double Top { get; init; } = DrawingValues.Range(Top, -100, 100, nameof(Top), "An inset");

    /// <summary>Inset from the right.</summary>
    public double Right { get; init; } = DrawingValues.Range(Right, -100, 100, nameof(Right), "An inset");

    /// <summary>Inset from the bottom.</summary>
    public double Bottom { get; init; } = DrawingValues.Range(Bottom, -100, 100, nameof(Bottom), "An inset");
}

/// <summary>How tiles are mirrored (a:tile flip).</summary>
public enum TileFlip
{
    /// <summary>Not at all.</summary>
    None,

    /// <summary>Every other tile left to right.</summary>
    Horizontal,

    /// <summary>Every other tile top to bottom.</summary>
    Vertical,

    /// <summary>Both.</summary>
    Both,
}

/// <summary>How a picture fill is tiled (a:tile) instead of stretched, as on Office's "Tile picture as texture".</summary>
public sealed record PictureTiling
{
    private readonly double scaleX = 1;
    private readonly double scaleY = 1;

    /// <summary>The horizontal offset of the tiles.</summary>
    public Length OffsetX { get; init; }

    /// <summary>The vertical offset of the tiles.</summary>
    public Length OffsetY { get; init; }

    /// <summary>The horizontal scale of a tile, 0 to 100; 1 is the picture's own size.</summary>
    public double ScaleX
    {
        get => scaleX;
        init => scaleX = DrawingValues.Range(value, 0, 100, nameof(ScaleX), "A tile's scale");
    }

    /// <summary>The vertical scale of a tile.</summary>
    public double ScaleY
    {
        get => scaleY;
        init => scaleY = DrawingValues.Range(value, 0, 100, nameof(ScaleY), "A tile's scale");
    }

    /// <summary>How the tiles are mirrored.</summary>
    public TileFlip Flip { get; init; }

    /// <summary>Where the first tile sits; the top-left corner unless set.</summary>
    public RectangleAlignment Alignment { get; init; } = RectangleAlignment.TopLeft;
}

public sealed partial record Fill
{
    /// <summary>The group's fill: a shape in a group painted with what the group is filled with.</summary>
    public static Fill Group { get; } = new(FillKind.Group, null, 0, 0, []);

    /// <summary>
    /// The slide's background: the shape shows the background where it lies, hiding what is
    /// under it. Only a slide's shapes take it (p:sp useBgFill); written elsewhere, it is no fill.
    /// </summary>
    public static Fill Background { get; } = new(FillKind.Background, null, 0, 0, []);

    /// <summary>How a gradient spreads; <see cref="GradientShape.Linear"/> for a linear one and the other kinds.</summary>
    public GradientShape GradientShape { get; private init; }

    /// <summary>Where a radial, rectangular or path gradient starts (a:fillToRect), or null for the top-left corner.</summary>
    public FillRectangle? Focus { get; private init; }

    /// <summary>Whether a gradient or a picture turns with the shape; true unless made with <see cref="WithoutRotation"/>.</summary>
    public bool RotateWithShape { get; private init; } = true;

    /// <summary>Whether a linear gradient's angle is scaled with the shape (a:lin scaled).</summary>
    public bool Scaled { get; private init; }

    /// <summary>How a gradient is placed before it repeats (a:tileRect), or null for the shape itself.</summary>
    public FillRectangle? TileRectangle { get; private init; }

    /// <summary>How a picture fill is tiled, or null when it is stretched.</summary>
    public PictureTiling? Tiling { get; private init; }

    /// <summary>A radial, rectangular or path gradient through two or more colours.</summary>
    /// <param name="shape">How it spreads: <see cref="GradientShape.Circle"/>, <see cref="GradientShape.Rectangle"/> or <see cref="GradientShape.Shape"/>.</param>
    /// <param name="focus">Where it starts, such as <see cref="FillRectangle.Center"/>; null for the top-left corner.</param>
    /// <param name="stops">The colours from the focus out, at least two.</param>
    /// <returns>The fill.</returns>
    public static Fill PathGradient(GradientShape shape, FillRectangle? focus, params GradientStop[] stops)
    {
        if (shape == GradientShape.Linear || Enum.IsDefined(shape) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(shape), shape, "A path gradient spreads as a circle, a rectangle or the shape; a linear one is made with Gradient.");
        }

        return Gradient(0, stops) with { GradientShape = shape, Focus = focus };
    }

    /// <summary>A picture tiled over the shape, as a texture.</summary>
    /// <param name="image">The image.</param>
    /// <param name="tiling">How the tiles are placed; null for Office's default (the picture's own size from the top-left corner).</param>
    /// <returns>The fill.</returns>
    public static Fill TiledPicture(Picture image, PictureTiling? tiling = null)
    {
        return Picture(image) with { Tiling = tiling ?? new PictureTiling() };
    }

    /// <summary>The same gradient or picture fill, but kept still when the shape turns.</summary>
    /// <returns>The fill.</returns>
    public Fill WithoutRotation()
    {
        return this with { RotateWithShape = false };
    }

    /// <summary>The same linear gradient, with its angle scaled with the shape (a:lin scaled="1").</summary>
    /// <returns>The fill.</returns>
    public Fill WithScaledAngle()
    {
        return this with { Scaled = true };
    }

    /// <summary>The same gradient, placed in a rectangle before it repeats (a:tileRect).</summary>
    /// <param name="rectangle">The rectangle.</param>
    /// <returns>The fill.</returns>
    public Fill WithTileRectangle(FillRectangle rectangle)
    {
        ArgumentNullException.ThrowIfNull(rectangle);

        return this with { TileRectangle = rectangle };
    }
}
