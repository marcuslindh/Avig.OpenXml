namespace Avig.OpenXml.Drawing;

/// <summary>
/// A place and a size, as DrawingML stores them in a:off and a:ext: on a sheet or a slide from
/// its top-left corner, or inside a group from the group's own corner. The package that owns the
/// drawing decides what the place is measured from.
/// </summary>
public readonly record struct DrawingBounds
{
    /// <summary>Creates bounds. The place and the size must be 0 or more.</summary>
    /// <param name="left">From the left edge.</param>
    /// <param name="top">From the top edge.</param>
    /// <param name="width">The width.</param>
    /// <param name="height">The height.</param>
    public DrawingBounds(Length left, Length top, Length width, Length height)
    {
        Check(left.Emu, top.Emu, width.Emu, height.Emu);
        Left = left;
        Top = top;
        Width = width;
        Height = height;
    }

    private static void Check(long left, long top, long width, long height)
    {
        if (left < 0 || top < 0 || width < 0 || height < 0 || left > DrawingValues.MaxCoordinate || top > DrawingValues.MaxCoordinate
            || width > DrawingValues.MaxCoordinate || height > DrawingValues.MaxCoordinate)
        {
            throw new ArgumentOutOfRangeException(nameof(left), "A drawing's position and size must be 0 or more and inside the sheet.");
        }
    }

    private DrawingBounds(long left, long top, long width, long height)
    {
        Left = Length.FromEmu(left);
        Top = Length.FromEmu(top);
        Width = Length.FromEmu(width);
        Height = Length.FromEmu(height);
    }

    /// <summary>The distance from the left edge.</summary>
    public Length Left { get; }

    /// <summary>The distance from the top edge.</summary>
    public Length Top { get; }

    /// <summary>The width.</summary>
    public Length Width { get; }

    /// <summary>The height.</summary>
    public Length Height { get; }

    /// <summary>The right edge: <see cref="Left"/> plus <see cref="Width"/>.</summary>
    public Length Right => Left + Width;

    /// <summary>The bottom edge: <see cref="Top"/> plus <see cref="Height"/>.</summary>
    public Length Bottom => Top + Height;

    /// <summary>A place in points (72 to the inch). Excel's default column is 48 points wide and its default row 15 high.</summary>
    /// <param name="left">From the left edge.</param>
    /// <param name="top">From the top edge.</param>
    /// <param name="width">The width.</param>
    /// <param name="height">The height.</param>
    /// <returns>The bounds.</returns>
    public static DrawingBounds FromPoints(double left, double top, double width, double height)
    {
        foreach (double value in (ReadOnlySpan<double>)[left, top, width, height])
        {
            if (double.IsFinite(value) == false || value < 0 || value > 2_000_000)
            {
                throw new ArgumentOutOfRangeException(nameof(left), value, "A drawing's position and size go from 0 to 2 000 000 points.");
            }
        }

        return FromEmu(DrawingValues.Emu(left), DrawingValues.Emu(top), DrawingValues.Emu(width), DrawingValues.Emu(height));
    }

    /// <summary>A place in EMU, as DrawingML stores it: 914 400 to the inch, 12 700 to the point.</summary>
    /// <param name="left">From the left edge.</param>
    /// <param name="top">From the top edge.</param>
    /// <param name="width">The width.</param>
    /// <param name="height">The height.</param>
    /// <returns>The bounds.</returns>
    public static DrawingBounds FromEmu(long left, long top, long width, long height)
    {
        Check(left, top, width, height);

        return new DrawingBounds(left, top, width, height);
    }

    /// <summary>
    /// Bounds read from a file, kept as they are: a slide may place a shape partly off its left
    /// or top edge, which <see cref="FromEmu"/> refuses for what a program writes. A negative
    /// size is read as 0. For readers; bounds to write are made with <see cref="FromEmu"/>.
    /// </summary>
    /// <param name="left">From the left edge, in EMU; may be negative.</param>
    /// <param name="top">From the top edge, in EMU; may be negative.</param>
    /// <param name="width">The width in EMU.</param>
    /// <param name="height">The height in EMU.</param>
    /// <returns>The bounds.</returns>
    public static DrawingBounds Read(long left, long top, long width, long height)
    {
        return new DrawingBounds(left, top, Math.Max(width, 0), Math.Max(height, 0));
    }
}
