namespace Avig.OpenXml.Drawing;

/// <summary>The names DrawingML stores the shape, line and effect enums as, both ways.</summary>
internal static class ShapeNames
{
    private static readonly string[] PathFills = ["norm", "none", "lighten", "lightenLess", "darken", "darkenLess"];
    private static readonly string[] Blends = ["over", "mult", "screen", "darken", "lighten"];
    private static readonly string[] Alignments = ["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"];
    private static readonly string[] Directions = ["tl", "t", "tr", "l", "r", "bl", "b", "br"];
    private static readonly string[] Fonts = ["minor", "major", "none"];
    private static readonly string[] Caps = ["flat", "rnd", "sq"];
    private static readonly string[] Compounds = ["sng", "dbl", "thickThin", "thinThick", "tri"];
    private static readonly string[] EndTypes = ["none", "triangle", "stealth", "diamond", "oval", "arrow"];
    private static readonly string[] EndSizes = ["med", "sm", "lg"];
    private static readonly string[] GradientPaths = ["lin", "circle", "rect", "shape"];
    private static readonly string[] Flips = ["none", "x", "y", "xy"];

    private static readonly string[] Bevels =
    [
        "circle", "relaxedInset", "slope", "cross", "angle", "softRound", "convex", "coolSlant", "divot", "riblet", "hardEdge", "artDeco",
    ];

    private static readonly string[] Materials =
    [
        "warmMatte", "legacyMatte", "legacyPlastic", "legacyMetal", "legacyWireframe", "matte", "plastic", "metal", "translucentPowder",
        "powder", "dkEdge", "softEdge", "clear", "flat", "softmetal",
    ];

    private static readonly string[] Lights =
    [
        "threePt", "legacyFlat1", "legacyFlat2", "legacyFlat3", "legacyFlat4", "legacyNormal1", "legacyNormal2", "legacyNormal3",
        "legacyNormal4", "legacyHarsh1", "legacyHarsh2", "legacyHarsh3", "legacyHarsh4", "balanced", "soft", "harsh", "flood",
        "contrasting", "morning", "sunrise", "sunset", "chilly", "freezing", "flat", "twoPt", "glow", "brightRoom",
    ];

    private static readonly string[] Cameras = Enum.GetValues<CameraPreset>().Select(camera => CamelCase(camera.ToString())).ToArray();

    public static string PathFill(PathFill fill) => PathFills[(int)fill];

    public static PathFill PathFillOf(string? name) => Of(PathFills, name, Drawing.PathFill.Normal);

    public static string Blend(BlendMode blend) => Blends[(int)blend];

    public static BlendMode BlendOf(string? name) => Of(Blends, name, BlendMode.Over);

    public static string Alignment(RectangleAlignment alignment) => Alignments[(int)alignment];

    public static RectangleAlignment AlignmentOf(string? name, RectangleAlignment fallback) => Of(Alignments, name, fallback);

    public static string Direction(LightDirection direction) => Directions[(int)direction];

    public static LightDirection DirectionOf(string? name) => Of(Directions, name, LightDirection.Top);

    public static string Font(ThemeFont font) => Fonts[(int)font];

    public static ThemeFont FontOf(string? name) => Of(Fonts, name, ThemeFont.Minor);

    public static string Cap(LineCap cap) => Caps[(int)cap];

    public static bool TryCapOf(string? name, out LineCap cap) => TryOf(Caps, name, out cap);

    public static string Compound(LineCompound compound) => Compounds[(int)compound];

    public static LineCompound CompoundOf(string? name) => Of(Compounds, name, LineCompound.SingleLine);

    public static string EndType(LineEndType type) => EndTypes[(int)type];

    public static LineEndType EndTypeOf(string? name) => Of(EndTypes, name, LineEndType.None);

    public static string EndSize(LineEndSize size) => EndSizes[(int)size];

    public static LineEndSize EndSizeOf(string? name) => Of(EndSizes, name, LineEndSize.Medium);

    public static string GradientPath(GradientShape shape) => GradientPaths[(int)shape];

    public static GradientShape GradientPathOf(string? name) => Of(GradientPaths, name, GradientShape.Circle);

    public static string Flip(TileFlip flip) => Flips[(int)flip];

    public static TileFlip FlipOf(string? name) => Of(Flips, name, TileFlip.None);

    public static string Bevel(BevelType bevel) => Bevels[(int)bevel];

    public static BevelType BevelOf(string? name) => Of(Bevels, name, BevelType.Circle);

    public static string Material(MaterialType material) => Materials[(int)material];

    public static MaterialType MaterialOf(string? name) => Of(Materials, name, MaterialType.WarmMatte);

    public static string Light(LightRigType light) => Lights[(int)light];

    public static LightRigType LightOf(string? name) => Of(Lights, name, LightRigType.ThreePoint);

    public static string Camera(CameraPreset camera) => Cameras[(int)camera];

    public static CameraPreset CameraOf(string? name) => Of(Cameras, name, CameraPreset.OrthographicFront);

    private static T Of<T>(string[] names, string? name, T fallback)
        where T : struct, Enum
    {
        if (TryOf(names, name, out T found) == true)
        {
            return found;
        }

        return fallback;
    }

    private static bool TryOf<T>(string[] names, string? name, out T found)
        where T : struct, Enum
    {
        int index = Array.IndexOf(names, name);
        found = default;

        if (index < 0)
        {
            return false;
        }

        found = (T)(object)index;

        return true;
    }

    private static string CamelCase(string name)
    {
        return char.ToLowerInvariant(name[0]) + name[1..];
    }
}
