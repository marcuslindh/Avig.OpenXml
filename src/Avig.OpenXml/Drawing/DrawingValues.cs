using System.Globalization;

namespace Avig.OpenXml.Drawing;

/// <summary>Checks and conversions shared by the drawing types: colours, units and ranges.</summary>
internal static class DrawingValues
{
    /// <summary>EMU (English Metric Units) per point.</summary>
    public const long EmuPerPoint = Length.EmuPerPoint;

    /// <summary>The largest coordinate DrawingML accepts (ST_Coordinate).</summary>
    public const long MaxCoordinate = 27_273_042_316_900;

    /// <summary>The most characters a text run holds; the same as an Excel cell.</summary>
    public const int MaxTextLength = 32_767;

    /// <summary>A colour as #RRGGBB in upper case, or an ArgumentException naming the parameter.</summary>
    public static string Color(string value, string parameter)
    {
        ArgumentNullException.ThrowIfNull(value, parameter);
        string hex = value;

        if (hex.StartsWith('#') == true)
        {
            hex = hex[1..];
        }

        if (hex.Length != 6 || int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _) == false)
        {
            throw new ArgumentException($"'{value}' is not a colour. Expected #RRGGBB.", parameter);
        }

        return "#" + hex.ToUpperInvariant();
    }

    /// <summary>
    /// Text for a drawing or a chart, which DrawingML stores as plain XML text: a control
    /// character other than tab and line breaks, or half a surrogate pair, cannot be written, so
    /// it is an ArgumentException here rather than a failed save later.
    /// </summary>
    public static string XmlText(string value, string parameter)
    {
        for (int index = 0; index < value.Length; index++)
        {
            char current = value[index];

            if (char.IsHighSurrogate(current) == true && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1]) == true)
            {
                index++;

                continue;
            }

            if (System.Xml.XmlConvert.IsXmlChar(current) == false)
            {
                throw new ArgumentException($"The text holds the character U+{(int)current:X4} at position {index}, which an XML file cannot store.", parameter);
            }
        }

        return value;
    }

    /// <summary>A transparency from 0 (opaque) to 1 (invisible).</summary>
    public static double Transparency(double value, string parameter)
    {
        if (double.IsFinite(value) == false || value < 0 || value > 1)
        {
            throw new ArgumentOutOfRangeException(parameter, value, "A transparency goes from 0 (opaque) to 1 (invisible).");
        }

        return value;
    }

    /// <summary>A finite number in a range, inclusive.</summary>
    public static double Range(double value, double minimum, double maximum, string parameter, string what)
    {
        if (double.IsFinite(value) == false || value < minimum || value > maximum)
        {
            throw new ArgumentOutOfRangeException(parameter, value, $"{what} goes from {Text(minimum)} to {Text(maximum)}.");
        }

        return value;
    }

    /// <summary>An integer in a range, inclusive.</summary>
    public static int Range(int value, int minimum, int maximum, string parameter, string what)
    {
        if (value < minimum || value > maximum)
        {
            throw new ArgumentOutOfRangeException(parameter, value, $"{what} goes from {Text(minimum)} to {Text(maximum)}.");
        }

        return value;
    }

    /// <summary>Points to EMU, rounded.</summary>
    public static long Emu(double points)
    {
        return (long)Math.Round(points * EmuPerPoint);
    }

    /// <summary>EMU to points.</summary>
    public static double Points(long emu)
    {
        return (double)emu / EmuPerPoint;
    }

    /// <summary>Degrees to DrawingML's 60 000ths of a degree, turned into 0 to 360.</summary>
    public static long Angle(double degrees)
    {
        double turned = degrees % 360;

        if (turned < 0)
        {
            turned += 360;
        }

        return (long)Math.Round(turned * 60_000) % 21_600_000;
    }

    /// <summary>A fraction (0 to 1) to DrawingML's 100 000ths.</summary>
    public static long Percent(double fraction)
    {
        return (long)Math.Round(fraction * 100_000);
    }

    public static string Text(double number)
    {
        return number.ToString("R", CultureInfo.InvariantCulture);
    }

    public static string Text(long number)
    {
        return number.ToString(CultureInfo.InvariantCulture);
    }

    public static string Flag(bool value)
    {
        if (value == true)
        {
            return "1";
        }

        return "0";
    }

    /// <summary>A whole number from an attribute, or false when it is missing or not one.</summary>
    public static bool TryLong(string? text, out long value)
    {
        return long.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value);
    }

    /// <summary>A number from an attribute or an element, or false when it is missing or not a finite one.</summary>
    public static bool TryDouble(string? text, out double value)
    {
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value) == true && double.IsFinite(value) == true;
    }

    /// <summary>An xsd:boolean: 1 and true are true, 0 and false false; anything else is the default.</summary>
    public static bool Boolean(string? text, bool fallback)
    {
        return text switch
        {
            "1" or "true" => true,
            "0" or "false" => false,
            _ => fallback,
        };
    }
}
