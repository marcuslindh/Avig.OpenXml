namespace Avig.OpenXml.Drawing;

/// <summary>A glow around a shape (a:glow), as on Office's Glow menu.</summary>
public sealed record GlowEffect
{
    /// <summary>Creates a glow.</summary>
    /// <param name="color">The colour as <c>#RRGGBB</c>.</param>
    /// <param name="radius">How far it reaches, in points, 0 to 150; Office's presets are 5, 8, 11 and 18.</param>
    /// <param name="transparency">0 for opaque to 1 for invisible; Office's presets are 0.6.</param>
    public GlowEffect(string color, double radius = 8, double transparency = 0.6)
    {
        Color = DrawingValues.Color(color, nameof(color));
        Radius = DrawingValues.Range(radius, 0, 150, nameof(radius), "A glow's radius");
        Transparency = DrawingValues.Transparency(transparency, nameof(transparency));
    }

    /// <summary>The colour as <c>#RRGGBB</c>.</summary>
    public string Color { get; }

    /// <summary>How far it reaches, in points.</summary>
    public double Radius { get; }

    /// <summary>0 for opaque to 1 for invisible.</summary>
    public double Transparency { get; }
}

/// <summary>
/// A reflection below a shape (a:reflection). The defaults are Office's "Tight Reflection: Touching".
/// </summary>
public sealed record ReflectionEffect
{
    private readonly double blur = 0.5;
    private readonly double startTransparency = 0.5;
    private readonly double startPosition;
    private readonly double endTransparency = 0.997;
    private readonly double size = 0.35;
    private readonly double distance;
    private readonly double direction = 90;
    private readonly double fadeDirection = 90;
    private readonly double scaleX = 1;
    private readonly double scaleY = -1;

    /// <summary>How soft it is, in points, 0 to 100.</summary>
    public double Blur
    {
        get => blur;
        init => blur = DrawingValues.Range(value, 0, 100, nameof(Blur), "A reflection's blur");
    }

    /// <summary>The transparency where it starts, 0 to 1.</summary>
    public double StartTransparency
    {
        get => startTransparency;
        init => startTransparency = DrawingValues.Transparency(value, nameof(StartTransparency));
    }

    /// <summary>Where the fade starts, 0 to 1 of the shape.</summary>
    public double StartPosition
    {
        get => startPosition;
        init => startPosition = DrawingValues.Range(value, 0, 1, nameof(StartPosition), "A reflection's start");
    }

    /// <summary>The transparency where it ends, 0 to 1.</summary>
    public double EndTransparency
    {
        get => endTransparency;
        init => endTransparency = DrawingValues.Transparency(value, nameof(EndTransparency));
    }

    /// <summary>How much of the shape is reflected, 0 to 1 (Office's Size).</summary>
    public double Size
    {
        get => size;
        init => size = DrawingValues.Range(value, 0, 1, nameof(Size), "A reflection's size");
    }

    /// <summary>The gap between the shape and its reflection, in points, 0 to 200.</summary>
    public double Distance
    {
        get => distance;
        init => distance = DrawingValues.Range(value, 0, 200, nameof(Distance), "A reflection's distance");
    }

    /// <summary>The direction it is moved in, in degrees clockwise from the right; 90 is down.</summary>
    public double Direction
    {
        get => direction;
        init => direction = DrawingValues.Range(value, -360, 360, nameof(Direction), "A reflection's direction");
    }

    /// <summary>The direction it fades in, in degrees.</summary>
    public double FadeDirection
    {
        get => fadeDirection;
        init => fadeDirection = DrawingValues.Range(value, -360, 360, nameof(FadeDirection), "A reflection's fade direction");
    }

    /// <summary>The horizontal scale, -2 to 2; 1 unless set.</summary>
    public double ScaleX
    {
        get => scaleX;
        init => scaleX = DrawingValues.Range(value, -2, 2, nameof(ScaleX), "A scale");
    }

    /// <summary>The vertical scale, -2 to 2; -1 (mirrored) unless set.</summary>
    public double ScaleY
    {
        get => scaleY;
        init => scaleY = DrawingValues.Range(value, -2, 2, nameof(ScaleY), "A scale");
    }

    /// <summary>What the reflection is lined up to; the bottom-left corner unless set.</summary>
    public RectangleAlignment Alignment { get; init; } = RectangleAlignment.BottomLeft;

    /// <summary>Whether it turns with the shape.</summary>
    public bool RotateWithShape { get; init; }
}

/// <summary>A point of a rectangle that an effect is lined up to (ST_RectAlignment).</summary>
public enum RectangleAlignment
{
    /// <summary>The top-left corner.</summary>
    TopLeft,

    /// <summary>The middle of the top.</summary>
    Top,

    /// <summary>The top-right corner.</summary>
    TopRight,

    /// <summary>The middle of the left side.</summary>
    Left,

    /// <summary>The centre.</summary>
    Center,

    /// <summary>The middle of the right side.</summary>
    Right,

    /// <summary>The bottom-left corner.</summary>
    BottomLeft,

    /// <summary>The middle of the bottom.</summary>
    Bottom,

    /// <summary>The bottom-right corner.</summary>
    BottomRight,
}

/// <summary>A blur of the whole shape (a:blur).</summary>
/// <param name="Radius">How soft, in points, 0 to 100.</param>
/// <param name="Grow">Whether the shape grows by the blur; true as in Office.</param>
public sealed record BlurEffect(double Radius, bool Grow = true)
{
    /// <summary>How soft, in points.</summary>
    public double Radius { get; init; } = DrawingValues.Range(Radius, 0, 100, nameof(Radius), "A blur's radius");
}

/// <summary>One of DrawingML's twenty preset shadows (a:prstShdw), kept for files from older Office.</summary>
public sealed record PresetShadow
{
    private readonly int preset = 1;
    private readonly string color = "#000000";
    private readonly double transparency;
    private readonly double distance = 3;
    private readonly double angle = 45;

    /// <summary>The preset, 1 to 20 (shdw1 to shdw20).</summary>
    public int Preset
    {
        get => preset;
        init => preset = DrawingValues.Range(value, 1, 20, nameof(Preset), "A preset shadow");
    }

    /// <summary>The colour as <c>#RRGGBB</c>; black unless set.</summary>
    public string Color
    {
        get => color;
        init => color = DrawingValues.Color(value, nameof(Color));
    }

    /// <summary>0 for opaque to 1 for invisible.</summary>
    public double Transparency
    {
        get => transparency;
        init => transparency = DrawingValues.Transparency(value, nameof(Transparency));
    }

    /// <summary>How far it falls, in points, 0 to 200.</summary>
    public double Distance
    {
        get => distance;
        init => distance = DrawingValues.Range(value, 0, 200, nameof(Distance), "A shadow's distance");
    }

    /// <summary>The direction it falls in, in degrees clockwise from the right.</summary>
    public double Angle
    {
        get => angle;
        init => angle = DrawingValues.Range(value, -360, 360, nameof(Angle), "A shadow's angle");
    }
}

/// <summary>How a <see cref="FillOverlayEffect"/> is blended with the shape (ST_BlendMode).</summary>
public enum BlendMode
{
    /// <summary>Over the shape.</summary>
    Over,

    /// <summary>Multiplied.</summary>
    Multiply,

    /// <summary>Screened.</summary>
    Screen,

    /// <summary>The darker of the two.</summary>
    Darken,

    /// <summary>The lighter of the two.</summary>
    Lighten,
}

/// <summary>A fill laid over the shape's own (a:fillOverlay).</summary>
public sealed record FillOverlayEffect
{
    /// <summary>Creates an overlay.</summary>
    /// <param name="fill">The fill; a solid, gradient or pattern fill.</param>
    /// <param name="blend">How it is blended.</param>
    public FillOverlayEffect(Fill fill, BlendMode blend = BlendMode.Over)
    {
        ArgumentNullException.ThrowIfNull(fill);

        if (fill.Kind is FillKind.Picture or FillKind.Group or FillKind.Background)
        {
            throw new ArgumentException("An overlay is a solid, gradient or pattern fill, or none.", nameof(fill));
        }

        if (Enum.IsDefined(blend) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(blend), blend, "Not a blend mode.");
        }

        Fill = fill;
        Blend = blend;
    }

    /// <summary>The fill.</summary>
    public Fill Fill { get; }

    /// <summary>How it is blended.</summary>
    public BlendMode Blend { get; }
}

/// <summary>The bevels of a <see cref="Shape3D"/> (ST_BevelPresetType), as on Office's Bevel menu.</summary>
public enum BevelType
{
    /// <summary>Circle, the default.</summary>
    Circle,

    /// <summary>Relaxed inset.</summary>
    RelaxedInset,

    /// <summary>Slope.</summary>
    Slope,

    /// <summary>Cross.</summary>
    Cross,

    /// <summary>Angle.</summary>
    Angle,

    /// <summary>Soft round.</summary>
    SoftRound,

    /// <summary>Convex.</summary>
    Convex,

    /// <summary>Cool slant.</summary>
    CoolSlant,

    /// <summary>Divot.</summary>
    Divot,

    /// <summary>Riblet.</summary>
    Riblet,

    /// <summary>Hard edge.</summary>
    HardEdge,

    /// <summary>Art deco.</summary>
    ArtDeco,
}

/// <summary>A bevel on the top or the bottom of a 3-D shape (a:bevelT, a:bevelB).</summary>
public sealed record Bevel
{
    private readonly BevelType type;
    private readonly double width = 6;
    private readonly double height = 6;

    /// <summary>The bevel; <see cref="BevelType.Circle"/> unless set.</summary>
    public BevelType Type
    {
        get => type;
        init
        {
            if (Enum.IsDefined(value) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(Type), value, "Not a bevel.");
            }

            type = value;
        }
    }

    /// <summary>The width in points, 0 to 1584; 6 unless set.</summary>
    public double Width
    {
        get => width;
        init => width = DrawingValues.Range(value, 0, 1584, nameof(Width), "A bevel's width");
    }

    /// <summary>The height in points, 0 to 1584; 6 unless set.</summary>
    public double Height
    {
        get => height;
        init => height = DrawingValues.Range(value, 0, 1584, nameof(Height), "A bevel's height");
    }
}

/// <summary>The surfaces of a <see cref="Shape3D"/> (ST_PresetMaterialType), as on Office's Material menu.</summary>
public enum MaterialType
{
    /// <summary>Warm matte, the default.</summary>
    WarmMatte,

    /// <summary>Matte (the older one, <c>legacyMatte</c>).</summary>
    LegacyMatte,

    /// <summary>Plastic (the older one).</summary>
    LegacyPlastic,

    /// <summary>Metal (the older one).</summary>
    LegacyMetal,

    /// <summary>Wire frame (the older one).</summary>
    LegacyWireframe,

    /// <summary>Matte.</summary>
    Matte,

    /// <summary>Plastic.</summary>
    Plastic,

    /// <summary>Metal.</summary>
    Metal,

    /// <summary>Translucent powder.</summary>
    TranslucentPowder,

    /// <summary>Powder.</summary>
    Powder,

    /// <summary>Dark edge.</summary>
    DarkEdge,

    /// <summary>Soft edge.</summary>
    SoftEdge,

    /// <summary>Clear.</summary>
    Clear,

    /// <summary>Flat.</summary>
    Flat,

    /// <summary>Soft metal.</summary>
    SoftMetal,
}

/// <summary>
/// The 3-D format of a shape (a:sp3d), as on Office's 3-D Format pane: bevels, depth, contour,
/// material and the distance from the ground. It shows when the shape also has a
/// <see cref="Scene3D"/> that turns it, or a bevel.
/// </summary>
public sealed record Shape3D
{
    private readonly double depth;
    private readonly double contour;
    private readonly double distance;
    private readonly string? depthColor;
    private readonly string? contourColor;
    private readonly MaterialType material;

    /// <summary>The bevel on the top, or null for none.</summary>
    public Bevel? TopBevel { get; init; }

    /// <summary>The bevel on the bottom, or null for none.</summary>
    public Bevel? BottomBevel { get; init; }

    /// <summary>The depth in points, 0 to 1584.</summary>
    public double Depth
    {
        get => depth;
        init => depth = DrawingValues.Range(value, 0, 1584, nameof(Depth), "A depth");
    }

    /// <summary>The colour of the depth as <c>#RRGGBB</c>, or null for the shape's fill.</summary>
    public string? DepthColor
    {
        get => depthColor;
        init
        {
            if (value is not null)
            {
                value = DrawingValues.Color(value, nameof(DepthColor));
            }

            depthColor = value;
        }
    }

    /// <summary>The width of the contour in points, 0 to 1584.</summary>
    public double Contour
    {
        get => contour;
        init => contour = DrawingValues.Range(value, 0, 1584, nameof(Contour), "A contour");
    }

    /// <summary>The colour of the contour as <c>#RRGGBB</c>, or null for the shape's line.</summary>
    public string? ContourColor
    {
        get => contourColor;
        init
        {
            if (value is not null)
            {
                value = DrawingValues.Color(value, nameof(ContourColor));
            }

            contourColor = value;
        }
    }

    /// <summary>The surface; <see cref="MaterialType.WarmMatte"/> unless set.</summary>
    public MaterialType Material
    {
        get => material;
        init
        {
            if (Enum.IsDefined(value) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(Material), value, "Not a material.");
            }

            material = value;
        }
    }

    /// <summary>The distance from the ground in points, -1584 to 1584.</summary>
    public double Distance
    {
        get => distance;
        init => distance = DrawingValues.Range(value, -1584, 1584, nameof(Distance), "A distance from the ground");
    }
}

/// <summary>A turn in three dimensions (a:rot): latitude, longitude and revolution, in degrees from 0 to 360.</summary>
public sealed record Rotation3D
{
    /// <summary>Creates a turn.</summary>
    /// <param name="latitude">Around the horizontal axis (Office's Y rotation).</param>
    /// <param name="longitude">Around the vertical axis (Office's X rotation).</param>
    /// <param name="revolution">Around the axis through the shape (Office's Z rotation).</param>
    public Rotation3D(double latitude, double longitude, double revolution)
    {
        Latitude = DrawingValues.Range(latitude, 0, 360, nameof(latitude), "A latitude");
        Longitude = DrawingValues.Range(longitude, 0, 360, nameof(longitude), "A longitude");
        Revolution = DrawingValues.Range(revolution, 0, 360, nameof(revolution), "A revolution");
    }

    /// <summary>Around the horizontal axis.</summary>
    public double Latitude { get; }

    /// <summary>Around the vertical axis.</summary>
    public double Longitude { get; }

    /// <summary>Around the axis through the shape.</summary>
    public double Revolution { get; }
}

/// <summary>A point or a direction in 3-D (a:anchor, a:norm, a:up), in EMU.</summary>
/// <param name="Across">Along the width (x).</param>
/// <param name="Down">Along the height (y).</param>
/// <param name="Toward">Towards the viewer (z).</param>
public readonly record struct Vector3D(long Across, long Down, long Toward);

/// <summary>A plane behind a 3-D scene (a:backdrop): its anchor, the direction it faces and its up.</summary>
/// <param name="Anchor">A point on the plane.</param>
/// <param name="Normal">The direction the plane faces.</param>
/// <param name="Up">The plane's up direction.</param>
public sealed record Backdrop(Vector3D Anchor, Vector3D Normal, Vector3D Up);

/// <summary>
/// The camera and the light of a 3-D shape (a:scene3d), as on Office's 3-D Rotation pane. A
/// shape with no scene is flat; <see cref="CameraPreset.OrthographicFront"/> with a
/// <see cref="Shape3D"/> shows bevels from the front.
/// </summary>
public sealed record Scene3D
{
    private readonly CameraPreset camera;
    private readonly double? fieldOfView;
    private readonly double zoom = 1;
    private readonly LightRigType light = LightRigType.ThreePoint;
    private readonly LightDirection lightDirection = LightDirection.Top;

    /// <summary>The camera; <see cref="CameraPreset.OrthographicFront"/> unless set.</summary>
    public CameraPreset Camera
    {
        get => camera;
        init
        {
            if (Enum.IsDefined(value) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(Camera), value, "Not a camera.");
            }

            camera = value;
        }
    }

    /// <summary>The perspective in degrees, 0 to 180, or null for the camera's own.</summary>
    public double? FieldOfView
    {
        get => fieldOfView;
        init
        {
            if (value is double degrees)
            {
                DrawingValues.Range(degrees, 0, 180, nameof(FieldOfView), "A field of view");
            }

            fieldOfView = value;
        }
    }

    /// <summary>The camera's zoom, 0 to 10; 1 unless set.</summary>
    public double Zoom
    {
        get => zoom;
        init => zoom = DrawingValues.Range(value, 0, 10, nameof(Zoom), "A zoom");
    }

    /// <summary>The camera's own turn, or null for the preset's.</summary>
    public Rotation3D? CameraRotation { get; init; }

    /// <summary>The lighting; <see cref="LightRigType.ThreePoint"/> unless set.</summary>
    public LightRigType Light
    {
        get => light;
        init
        {
            if (Enum.IsDefined(value) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(Light), value, "Not a lighting.");
            }

            light = value;
        }
    }

    /// <summary>Where the light comes from; <see cref="LightDirection.Top"/> unless set.</summary>
    public LightDirection LightDirection
    {
        get => lightDirection;
        init
        {
            if (Enum.IsDefined(value) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(LightDirection), value, "Not a light direction.");
            }

            lightDirection = value;
        }
    }

    /// <summary>The light's own turn (Office's Angle under Lighting), or null.</summary>
    public Rotation3D? LightRotation { get; init; }

    /// <summary>A plane behind the scene, or null.</summary>
    public Backdrop? Backdrop { get; init; }

    /// <summary>Keeps the shape's text flat when the shape turns (a:bodyPr/a:flatTx).</summary>
    public bool KeepTextFlat { get; init; }
}

/// <summary>Where the light of a <see cref="Scene3D"/> comes from (ST_LightRigDirection).</summary>
public enum LightDirection
{
    /// <summary>The top-left.</summary>
    TopLeft,

    /// <summary>The top.</summary>
    Top,

    /// <summary>The top-right.</summary>
    TopRight,

    /// <summary>The left.</summary>
    Left,

    /// <summary>The right.</summary>
    Right,

    /// <summary>The bottom-left.</summary>
    BottomLeft,

    /// <summary>The bottom.</summary>
    Bottom,

    /// <summary>The bottom-right.</summary>
    BottomRight,
}
