using System.Globalization;

namespace Avig.OpenXml;

/// <summary>
/// A length on the page, the sheet or the slide: a margin, an indent, a picture's size, a font
/// size, a shape's place. Office stores lengths in several units (twentieths of a point for
/// Word's margins, half-points for Word's font sizes, English Metric Units for everything in
/// DrawingML); this type holds them all exactly, as EMUs, so the unit is never a guess. Shared by
/// Avig.Kalkyl, Avig.Ord and Avig.Bildspel.
/// </summary>
/// <example>
/// <code>
/// Length margin = Length.Cm(2.5);
/// Length size = Length.Pt(11);
/// double inMillimetres = margin.ToMm(); // 25
/// </code>
/// </example>
public readonly record struct Length : IComparable<Length>
{
    /// <summary>EMUs in one point: 12 700.</summary>
    public const long EmuPerPoint = 12_700;

    /// <summary>EMUs in one twip (a twentieth of a point): 635.</summary>
    public const long EmuPerTwip = 635;

    /// <summary>EMUs in one inch: 914 400.</summary>
    public const long EmuPerInch = 914_400;

    /// <summary>EMUs in one centimetre: 360 000.</summary>
    public const long EmuPerCm = 360_000;

    /// <summary>EMUs in one millimetre: 36 000.</summary>
    public const long EmuPerMm = 36_000;

    private Length(long emu)
    {
        Emu = emu;
    }

    /// <summary>No length.</summary>
    public static Length Zero { get; } = new(0);

    /// <summary>The length in English Metric Units, 914 400 to the inch, as Word stores picture sizes.</summary>
    public long Emu { get; }

    /// <summary>A length in points, 72 to the inch, as Word shows font sizes.</summary>
    public static Length Pt(double points)
    {
        return new Length((long)Math.Round(points * EmuPerPoint));
    }

    /// <summary>A length in centimetres, as Swedish Word shows margins.</summary>
    public static Length Cm(double centimetres)
    {
        return new Length((long)Math.Round(centimetres * EmuPerCm));
    }

    /// <summary>A length in millimetres.</summary>
    public static Length Mm(double millimetres)
    {
        return new Length((long)Math.Round(millimetres * EmuPerMm));
    }

    /// <summary>A length in inches.</summary>
    public static Length Inch(double inches)
    {
        return new Length((long)Math.Round(inches * EmuPerInch));
    }

    /// <summary>A length in English Metric Units.</summary>
    public static Length FromEmu(long emu)
    {
        return new Length(emu);
    }

    /// <summary>A length in pixels at 96 per inch, as a picture's size in pixels is shown on a screen at 100 %.</summary>
    public static Length Px(double pixels)
    {
        return new Length((long)Math.Round(pixels * EmuPerInch / 96));
    }

    /// <summary>Adds two lengths.</summary>
    public static Length operator +(Length left, Length right)
    {
        return new Length(left.Emu + right.Emu);
    }

    /// <summary>Subtracts a length.</summary>
    public static Length operator -(Length left, Length right)
    {
        return new Length(left.Emu - right.Emu);
    }

    /// <summary>Scales a length.</summary>
    public static Length operator *(Length length, double factor)
    {
        return new Length((long)Math.Round(length.Emu * factor));
    }

    /// <summary>True when the left length is shorter.</summary>
    public static bool operator <(Length left, Length right)
    {
        return left.Emu < right.Emu;
    }

    /// <summary>True when the left length is longer.</summary>
    public static bool operator >(Length left, Length right)
    {
        return left.Emu > right.Emu;
    }

    /// <summary>True when the left length is shorter or the same.</summary>
    public static bool operator <=(Length left, Length right)
    {
        return left.Emu <= right.Emu;
    }

    /// <summary>True when the left length is longer or the same.</summary>
    public static bool operator >=(Length left, Length right)
    {
        return left.Emu >= right.Emu;
    }

    /// <summary>The length in points.</summary>
    public double ToPt()
    {
        return (double)Emu / EmuPerPoint;
    }

    /// <summary>The length in centimetres.</summary>
    public double ToCm()
    {
        return (double)Emu / EmuPerCm;
    }

    /// <summary>The length in millimetres.</summary>
    public double ToMm()
    {
        return (double)Emu / EmuPerMm;
    }

    /// <summary>The length in inches.</summary>
    public double ToInch()
    {
        return (double)Emu / EmuPerInch;
    }

    /// <summary>The length in pixels at 96 per inch.</summary>
    public double ToPx()
    {
        return (double)Emu * 96 / EmuPerInch;
    }

    /// <inheritdoc/>
    public int CompareTo(Length other)
    {
        return Emu.CompareTo(other.Emu);
    }

    /// <summary>The length in points, for logs and debugging: "12pt".</summary>
    public override string ToString()
    {
        return ToPt().ToString("0.##", CultureInfo.InvariantCulture) + "pt";
    }

    /// <summary>
    /// A length from twentieths of a point, Word's unit for margins, indents and spacing. Past what
    /// a length holds (about 1.45e16 twips) it saturates, as <see cref="Pt"/> and the other
    /// factories do: the longest length of the same sign, never one of the other sign.
    /// </summary>
    public static Length FromTwips(long twips)
    {
        return new Length(Saturating(twips, EmuPerTwip));
    }

    /// <summary>
    /// A length from half-points, Word's unit for font sizes. Past what a length holds (about
    /// 1.45e15 half-points) it saturates, as <see cref="Pt"/> and the other factories do.
    /// </summary>
    public static Length FromHalfPoints(long halfPoints)
    {
        return new Length(Saturating(halfPoints, EmuPerPoint / 2));
    }

    /// <summary>value × factor, clamped to the range of a long instead of wrapping around.</summary>
    private static long Saturating(long value, long factor)
    {
        return long.CreateSaturating((Int128)value * factor);
    }

    /// <summary>The length in whole twips, rounded.</summary>
    public long ToTwips()
    {
        return (long)Math.Round((double)Emu / EmuPerTwip);
    }

    /// <summary>The length in whole half-points, rounded.</summary>
    public long ToHalfPoints()
    {
        return (long)Math.Round((double)Emu * 2 / EmuPerPoint);
    }
}
