using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Xml.Linq;

namespace Avig.OpenXml.Drawing;

/// <summary>
/// Reads a DrawingML colour as the application shows it: an a:srgbClr, a:sysClr, a:scrgbClr,
/// a:hslClr, a:prstClr or a theme colour (a:schemeClr), with its modifiers applied in the order
/// the file gives them (a:lumMod, a:lumOff, a:tint, a:shade, a:satMod, a:alpha …). Theme colours
/// need the colours of the theme, which the package that owns the part knows: in PowerPoint the
/// colour map of the slide master decides that <c>tx1</c> is the theme's <c>dk1</c>, so the
/// caller gives every name it can meet (<c>tx1</c>, <c>bg1</c>, <c>accent1</c>, <c>dk1</c> …).
/// Elements are read by local name, so Strict OOXML works too.
/// </summary>
public static class DrawingColors
{
    /// <summary>The colour elements of DrawingML (EG_ColorChoice), by local name.</summary>
    private static readonly string[] ColorElements = ["srgbClr", "schemeClr", "sysClr", "scrgbClr", "hslClr", "prstClr"];

    /// <summary>The most modifiers that are applied to one colour; the rest are ignored, so that a hostile file cannot make one colour cost much.</summary>
    private const int MaxModifiers = 64;

    /// <summary>Reads the first colour element inside <paramref name="parent"/>, such as the child of an a:solidFill.</summary>
    /// <param name="parent">The element that holds the colour, or null.</param>
    /// <param name="schemeColors">The theme's colours by name as <c>#RRGGBB</c> (<c>accent1</c>, <c>tx1</c>, <c>bg1</c> …), or null when theme colours cannot be resolved.</param>
    /// <param name="color">The colour as <c>#RRGGBB</c> in upper case.</param>
    /// <param name="transparency">How transparent it is, 0 (opaque) to 1, from its a:alpha.</param>
    /// <returns>True when the parent holds a colour that can be resolved.</returns>
    public static bool TryRead(XElement? parent, IReadOnlyDictionary<string, string>? schemeColors, [NotNullWhen(true)] out string? color, out double transparency)
    {
        color = null;
        transparency = 0;
        XElement? element = parent?.Elements().FirstOrDefault(child => ColorElements.Contains(child.Name.LocalName) == true);

        if (element is null)
        {
            return false;
        }

        return TryReadElement(element, schemeColors, out color, out transparency);
    }

    /// <summary>Reads a colour element itself: an a:srgbClr, a:schemeClr, a:sysClr, a:scrgbClr, a:hslClr or a:prstClr.</summary>
    /// <param name="element">The colour element.</param>
    /// <param name="schemeColors">The theme's colours by name, or null.</param>
    /// <param name="color">The colour as <c>#RRGGBB</c> in upper case.</param>
    /// <param name="transparency">How transparent it is, 0 (opaque) to 1.</param>
    /// <returns>True when the element is a colour that can be resolved.</returns>
    public static bool TryReadElement(XElement element, IReadOnlyDictionary<string, string>? schemeColors, [NotNullWhen(true)] out string? color, out double transparency)
    {
        ArgumentNullException.ThrowIfNull(element);
        color = null;
        transparency = 0;

        if (TryBase(element, schemeColors, out double red, out double green, out double blue) == false)
        {
            return false;
        }

        double alpha = 1;
        int applied = 0;

        foreach (XElement modifier in element.Elements())
        {
            if (applied >= MaxModifiers)
            {
                break;
            }

            applied++;

            if (DrawingValues.TryLong(modifier.Attribute("val")?.Value, out long raw) == false)
            {
                if (modifier.Name.LocalName is "comp" or "inv" or "gray")
                {
                    raw = 0;
                }
                else
                {
                    continue;
                }
            }

            double value = raw / 100_000.0;
            Apply(modifier.Name.LocalName, value, raw, ref red, ref green, ref blue, ref alpha);
        }

        color = Hex(red, green, blue);
        transparency = Math.Clamp(1 - alpha, 0, 1);

        return true;
    }

    /// <summary>The colour before its modifiers, as sRGB from 0 to 1.</summary>
    private static bool TryBase(XElement element, IReadOnlyDictionary<string, string>? schemeColors, out double red, out double green, out double blue)
    {
        red = 0;
        green = 0;
        blue = 0;
        string? hex = null;

        switch (element.Name.LocalName)
        {
            case "srgbClr":
                hex = element.Attribute("val")?.Value;

                break;

            case "sysClr":
                hex = element.Attribute("lastClr")?.Value;

                if (hex is null && element.Attribute("val")?.Value == "window")
                {
                    hex = "FFFFFF";
                }
                else if (hex is null && element.Attribute("val")?.Value == "windowText")
                {
                    hex = "000000";
                }

                break;

            case "schemeClr":
                string? name = element.Attribute("val")?.Value;

                if (name is not null && schemeColors is not null && schemeColors.TryGetValue(name, out string? scheme) == true)
                {
                    hex = scheme.TrimStart('#');
                }

                break;

            case "prstClr":
                if (PresetColors.TryGetValue(element.Attribute("val")?.Value ?? string.Empty, out string? preset) == true)
                {
                    hex = preset;
                }

                break;

            case "scrgbClr":
                if (DrawingValues.TryLong(element.Attribute("r")?.Value, out long linearRed) == true
                    && DrawingValues.TryLong(element.Attribute("g")?.Value, out long linearGreen) == true
                    && DrawingValues.TryLong(element.Attribute("b")?.Value, out long linearBlue) == true)
                {
                    red = FromLinear(Math.Clamp(linearRed / 100_000.0, 0, 1));
                    green = FromLinear(Math.Clamp(linearGreen / 100_000.0, 0, 1));
                    blue = FromLinear(Math.Clamp(linearBlue / 100_000.0, 0, 1));

                    return true;
                }

                return false;

            case "hslClr":
                if (DrawingValues.TryLong(element.Attribute("hue")?.Value, out long hue) == true
                    && DrawingValues.TryLong(element.Attribute("sat")?.Value, out long saturation) == true
                    && DrawingValues.TryLong(element.Attribute("lum")?.Value, out long luminance) == true)
                {
                    FromHsl(Math.Clamp(hue / 60_000.0, 0, 360) / 360, Math.Clamp(saturation / 100_000.0, 0, 1), Math.Clamp(luminance / 100_000.0, 0, 1), out red, out green, out blue);

                    return true;
                }

                return false;
        }

        if (hex is null || hex.Length != 6 || hex.All(char.IsAsciiHexDigit) == false)
        {
            return false;
        }

        int rgb = int.Parse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        red = ((rgb >> 16) & 0xFF) / 255.0;
        green = ((rgb >> 8) & 0xFF) / 255.0;
        blue = (rgb & 0xFF) / 255.0;

        return true;
    }

    private static void Apply(string modifier, double value, long raw, ref double red, ref double green, ref double blue, ref double alpha)
    {
        switch (modifier)
        {
            case "alpha":
                alpha = Math.Clamp(value, 0, 1);

                return;

            case "alphaMod":
                alpha = Math.Clamp(alpha * value, 0, 1);

                return;

            case "alphaOff":
                alpha = Math.Clamp(alpha + value, 0, 1);

                return;

            case "tint":
                // Towards white, in linear RGB, as Office does.
                red = FromLinear(1 - (1 - ToLinear(red)) * Math.Clamp(value, 0, 1));
                green = FromLinear(1 - (1 - ToLinear(green)) * Math.Clamp(value, 0, 1));
                blue = FromLinear(1 - (1 - ToLinear(blue)) * Math.Clamp(value, 0, 1));

                return;

            case "shade":
                red = FromLinear(ToLinear(red) * Math.Clamp(value, 0, 1));
                green = FromLinear(ToLinear(green) * Math.Clamp(value, 0, 1));
                blue = FromLinear(ToLinear(blue) * Math.Clamp(value, 0, 1));

                return;

            case "comp":
                ToHsl(red, green, blue, out double compHue, out double compSaturation, out double compLuminance);
                FromHsl((compHue + 0.5) % 1, compSaturation, compLuminance, out red, out green, out blue);

                return;

            case "inv":
                red = 1 - red;
                green = 1 - green;
                blue = 1 - blue;

                return;

            case "gray":
                double gray = 0.3 * red + 0.59 * green + 0.11 * blue;
                red = gray;
                green = gray;
                blue = gray;

                return;
        }

        ToHsl(red, green, blue, out double hue, out double saturation, out double luminance);

        switch (modifier)
        {
            case "lumMod":
                luminance *= value;

                break;

            case "lumOff":
                luminance += value;

                break;

            case "lum":
                luminance = value;

                break;

            case "satMod":
                saturation *= value;

                break;

            case "satOff":
                saturation += value;

                break;

            case "sat":
                saturation = value;

                break;

            case "hueMod":
                hue = hue * value % 1;

                break;

            case "hueOff":
                hue = (hue + raw / 60_000.0 / 360) % 1;

                break;

            case "hue":
                hue = raw / 60_000.0 / 360 % 1;

                break;

            default:
                return;
        }

        if (hue < 0)
        {
            hue += 1;
        }

        FromHsl(hue, Math.Clamp(saturation, 0, 1), Math.Clamp(luminance, 0, 1), out red, out green, out blue);
    }

    private static double ToLinear(double value)
    {
        if (value <= 0.04045)
        {
            return value / 12.92;
        }

        return Math.Pow((value + 0.055) / 1.055, 2.4);
    }

    private static double FromLinear(double value)
    {
        double clamped = Math.Clamp(value, 0, 1);

        if (clamped <= 0.0031308)
        {
            return clamped * 12.92;
        }

        return 1.055 * Math.Pow(clamped, 1 / 2.4) - 0.055;
    }

    private static void ToHsl(double red, double green, double blue, out double hue, out double saturation, out double luminance)
    {
        double max = Math.Max(red, Math.Max(green, blue));
        double min = Math.Min(red, Math.Min(green, blue));
        luminance = (max + min) / 2;
        hue = 0;
        saturation = 0;

        if (max - min < 1e-12)
        {
            return;
        }

        double delta = max - min;

        if (luminance > 0.5)
        {
            saturation = delta / (2 - max - min);
        }
        else
        {
            saturation = delta / (max + min);
        }

        if (max == red)
        {
            hue = (green - blue) / delta;

            if (green < blue)
            {
                hue += 6;
            }
        }
        else if (max == green)
        {
            hue = (blue - red) / delta + 2;
        }
        else
        {
            hue = (red - green) / delta + 4;
        }

        hue /= 6;
    }

    private static void FromHsl(double hue, double saturation, double luminance, out double red, out double green, out double blue)
    {
        if (saturation < 1e-12)
        {
            red = luminance;
            green = luminance;
            blue = luminance;

            return;
        }

        double upper;

        if (luminance < 0.5)
        {
            upper = luminance * (1 + saturation);
        }
        else
        {
            upper = luminance + saturation - luminance * saturation;
        }

        double lower = 2 * luminance - upper;
        red = Channel(lower, upper, hue + 1.0 / 3);
        green = Channel(lower, upper, hue);
        blue = Channel(lower, upper, hue - 1.0 / 3);
    }

    private static double Channel(double lower, double upper, double hue)
    {
        double turn = hue;

        if (turn < 0)
        {
            turn += 1;
        }

        if (turn > 1)
        {
            turn -= 1;
        }

        if (turn < 1.0 / 6)
        {
            return lower + (upper - lower) * 6 * turn;
        }

        if (turn < 0.5)
        {
            return upper;
        }

        if (turn < 2.0 / 3)
        {
            return lower + (upper - lower) * (2.0 / 3 - turn) * 6;
        }

        return lower;
    }

    private static string Hex(double red, double green, double blue)
    {
        int redByte = (int)Math.Round(Math.Clamp(red, 0, 1) * 255, MidpointRounding.AwayFromZero);
        int greenByte = (int)Math.Round(Math.Clamp(green, 0, 1) * 255, MidpointRounding.AwayFromZero);
        int blueByte = (int)Math.Round(Math.Clamp(blue, 0, 1) * 255, MidpointRounding.AwayFromZero);

        return string.Create(CultureInfo.InvariantCulture, $"#{redByte:X2}{greenByte:X2}{blueByte:X2}");
    }

    /// <summary>The preset colours (ST_PresetColorVal) that Office writes, by name.</summary>
    private static readonly Dictionary<string, string> PresetColors = new(StringComparer.Ordinal)
    {
        ["black"] = "000000",
        ["white"] = "FFFFFF",
        ["red"] = "FF0000",
        ["green"] = "008000",
        ["lime"] = "00FF00",
        ["blue"] = "0000FF",
        ["yellow"] = "FFFF00",
        ["cyan"] = "00FFFF",
        ["aqua"] = "00FFFF",
        ["magenta"] = "FF00FF",
        ["fuchsia"] = "FF00FF",
        ["gray"] = "808080",
        ["grey"] = "808080",
        ["silver"] = "C0C0C0",
        ["maroon"] = "800000",
        ["navy"] = "000080",
        ["olive"] = "808000",
        ["purple"] = "800080",
        ["teal"] = "008080",
        ["orange"] = "FFA500",
        ["darkGray"] = "A9A9A9",
        ["lightGray"] = "D3D3D3",
        ["darkRed"] = "8B0000",
        ["darkBlue"] = "00008B",
        ["darkGreen"] = "006400",
    };
}
