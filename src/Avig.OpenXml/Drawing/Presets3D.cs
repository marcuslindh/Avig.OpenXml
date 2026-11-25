namespace Avig.OpenXml.Drawing;

/// <summary>The cameras of a <see cref="Scene3D"/> (ST_PresetCameraType), as on Office's 3-D Rotation presets.</summary>
public enum CameraPreset
{
    /// <summary>The <c>orthographicFront</c> camera.</summary>
    OrthographicFront,

    /// <summary>The <c>legacyObliqueTopLeft</c> camera.</summary>
    LegacyObliqueTopLeft,

    /// <summary>The <c>legacyObliqueTop</c> camera.</summary>
    LegacyObliqueTop,

    /// <summary>The <c>legacyObliqueTopRight</c> camera.</summary>
    LegacyObliqueTopRight,

    /// <summary>The <c>legacyObliqueLeft</c> camera.</summary>
    LegacyObliqueLeft,

    /// <summary>The <c>legacyObliqueFront</c> camera.</summary>
    LegacyObliqueFront,

    /// <summary>The <c>legacyObliqueRight</c> camera.</summary>
    LegacyObliqueRight,

    /// <summary>The <c>legacyObliqueBottomLeft</c> camera.</summary>
    LegacyObliqueBottomLeft,

    /// <summary>The <c>legacyObliqueBottom</c> camera.</summary>
    LegacyObliqueBottom,

    /// <summary>The <c>legacyObliqueBottomRight</c> camera.</summary>
    LegacyObliqueBottomRight,

    /// <summary>The <c>legacyPerspectiveTopLeft</c> camera.</summary>
    LegacyPerspectiveTopLeft,

    /// <summary>The <c>legacyPerspectiveTop</c> camera.</summary>
    LegacyPerspectiveTop,

    /// <summary>The <c>legacyPerspectiveTopRight</c> camera.</summary>
    LegacyPerspectiveTopRight,

    /// <summary>The <c>legacyPerspectiveLeft</c> camera.</summary>
    LegacyPerspectiveLeft,

    /// <summary>The <c>legacyPerspectiveFront</c> camera.</summary>
    LegacyPerspectiveFront,

    /// <summary>The <c>legacyPerspectiveRight</c> camera.</summary>
    LegacyPerspectiveRight,

    /// <summary>The <c>legacyPerspectiveBottomLeft</c> camera.</summary>
    LegacyPerspectiveBottomLeft,

    /// <summary>The <c>legacyPerspectiveBottom</c> camera.</summary>
    LegacyPerspectiveBottom,

    /// <summary>The <c>legacyPerspectiveBottomRight</c> camera.</summary>
    LegacyPerspectiveBottomRight,

    /// <summary>The <c>isometricTopUp</c> camera.</summary>
    IsometricTopUp,

    /// <summary>The <c>isometricTopDown</c> camera.</summary>
    IsometricTopDown,

    /// <summary>The <c>isometricBottomUp</c> camera.</summary>
    IsometricBottomUp,

    /// <summary>The <c>isometricBottomDown</c> camera.</summary>
    IsometricBottomDown,

    /// <summary>The <c>isometricLeftUp</c> camera.</summary>
    IsometricLeftUp,

    /// <summary>The <c>isometricLeftDown</c> camera.</summary>
    IsometricLeftDown,

    /// <summary>The <c>isometricRightUp</c> camera.</summary>
    IsometricRightUp,

    /// <summary>The <c>isometricRightDown</c> camera.</summary>
    IsometricRightDown,

    /// <summary>The <c>isometricOffAxis1Left</c> camera.</summary>
    IsometricOffAxis1Left,

    /// <summary>The <c>isometricOffAxis1Right</c> camera.</summary>
    IsometricOffAxis1Right,

    /// <summary>The <c>isometricOffAxis1Top</c> camera.</summary>
    IsometricOffAxis1Top,

    /// <summary>The <c>isometricOffAxis2Left</c> camera.</summary>
    IsometricOffAxis2Left,

    /// <summary>The <c>isometricOffAxis2Right</c> camera.</summary>
    IsometricOffAxis2Right,

    /// <summary>The <c>isometricOffAxis2Top</c> camera.</summary>
    IsometricOffAxis2Top,

    /// <summary>The <c>isometricOffAxis3Left</c> camera.</summary>
    IsometricOffAxis3Left,

    /// <summary>The <c>isometricOffAxis3Right</c> camera.</summary>
    IsometricOffAxis3Right,

    /// <summary>The <c>isometricOffAxis3Bottom</c> camera.</summary>
    IsometricOffAxis3Bottom,

    /// <summary>The <c>isometricOffAxis4Left</c> camera.</summary>
    IsometricOffAxis4Left,

    /// <summary>The <c>isometricOffAxis4Right</c> camera.</summary>
    IsometricOffAxis4Right,

    /// <summary>The <c>isometricOffAxis4Bottom</c> camera.</summary>
    IsometricOffAxis4Bottom,

    /// <summary>The <c>obliqueTopLeft</c> camera.</summary>
    ObliqueTopLeft,

    /// <summary>The <c>obliqueTop</c> camera.</summary>
    ObliqueTop,

    /// <summary>The <c>obliqueTopRight</c> camera.</summary>
    ObliqueTopRight,

    /// <summary>The <c>obliqueLeft</c> camera.</summary>
    ObliqueLeft,

    /// <summary>The <c>obliqueRight</c> camera.</summary>
    ObliqueRight,

    /// <summary>The <c>obliqueBottomLeft</c> camera.</summary>
    ObliqueBottomLeft,

    /// <summary>The <c>obliqueBottom</c> camera.</summary>
    ObliqueBottom,

    /// <summary>The <c>obliqueBottomRight</c> camera.</summary>
    ObliqueBottomRight,

    /// <summary>The <c>perspectiveFront</c> camera.</summary>
    PerspectiveFront,

    /// <summary>The <c>perspectiveLeft</c> camera.</summary>
    PerspectiveLeft,

    /// <summary>The <c>perspectiveRight</c> camera.</summary>
    PerspectiveRight,

    /// <summary>The <c>perspectiveAbove</c> camera.</summary>
    PerspectiveAbove,

    /// <summary>The <c>perspectiveBelow</c> camera.</summary>
    PerspectiveBelow,

    /// <summary>The <c>perspectiveAboveLeftFacing</c> camera.</summary>
    PerspectiveAboveLeftFacing,

    /// <summary>The <c>perspectiveAboveRightFacing</c> camera.</summary>
    PerspectiveAboveRightFacing,

    /// <summary>The <c>perspectiveContrastingLeftFacing</c> camera.</summary>
    PerspectiveContrastingLeftFacing,

    /// <summary>The <c>perspectiveContrastingRightFacing</c> camera.</summary>
    PerspectiveContrastingRightFacing,

    /// <summary>The <c>perspectiveHeroicLeftFacing</c> camera.</summary>
    PerspectiveHeroicLeftFacing,

    /// <summary>The <c>perspectiveHeroicRightFacing</c> camera.</summary>
    PerspectiveHeroicRightFacing,

    /// <summary>The <c>perspectiveHeroicExtremeLeftFacing</c> camera.</summary>
    PerspectiveHeroicExtremeLeftFacing,

    /// <summary>The <c>perspectiveHeroicExtremeRightFacing</c> camera.</summary>
    PerspectiveHeroicExtremeRightFacing,

    /// <summary>The <c>perspectiveRelaxed</c> camera.</summary>
    PerspectiveRelaxed,

    /// <summary>The <c>perspectiveRelaxedModerately</c> camera.</summary>
    PerspectiveRelaxedModerately,
}

/// <summary>The lightings of a <see cref="Scene3D"/> (ST_LightRigType), as on Office's Lighting menu.</summary>
public enum LightRigType
{
    /// <summary>The <c>threePt</c> lighting.</summary>
    ThreePoint,

    /// <summary>The <c>legacyFlat1</c> lighting.</summary>
    LegacyFlat1,

    /// <summary>The <c>legacyFlat2</c> lighting.</summary>
    LegacyFlat2,

    /// <summary>The <c>legacyFlat3</c> lighting.</summary>
    LegacyFlat3,

    /// <summary>The <c>legacyFlat4</c> lighting.</summary>
    LegacyFlat4,

    /// <summary>The <c>legacyNormal1</c> lighting.</summary>
    LegacyNormal1,

    /// <summary>The <c>legacyNormal2</c> lighting.</summary>
    LegacyNormal2,

    /// <summary>The <c>legacyNormal3</c> lighting.</summary>
    LegacyNormal3,

    /// <summary>The <c>legacyNormal4</c> lighting.</summary>
    LegacyNormal4,

    /// <summary>The <c>legacyHarsh1</c> lighting.</summary>
    LegacyHarsh1,

    /// <summary>The <c>legacyHarsh2</c> lighting.</summary>
    LegacyHarsh2,

    /// <summary>The <c>legacyHarsh3</c> lighting.</summary>
    LegacyHarsh3,

    /// <summary>The <c>legacyHarsh4</c> lighting.</summary>
    LegacyHarsh4,

    /// <summary>The <c>balanced</c> lighting.</summary>
    Balanced,

    /// <summary>The <c>soft</c> lighting.</summary>
    Soft,

    /// <summary>The <c>harsh</c> lighting.</summary>
    Harsh,

    /// <summary>The <c>flood</c> lighting.</summary>
    Flood,

    /// <summary>The <c>contrasting</c> lighting.</summary>
    Contrasting,

    /// <summary>The <c>morning</c> lighting.</summary>
    Morning,

    /// <summary>The <c>sunrise</c> lighting.</summary>
    Sunrise,

    /// <summary>The <c>sunset</c> lighting.</summary>
    Sunset,

    /// <summary>The <c>chilly</c> lighting.</summary>
    Chilly,

    /// <summary>The <c>freezing</c> lighting.</summary>
    Freezing,

    /// <summary>The <c>flat</c> lighting.</summary>
    Flat,

    /// <summary>The <c>twoPt</c> lighting.</summary>
    TwoPoint,

    /// <summary>The <c>glow</c> lighting.</summary>
    Glow,

    /// <summary>The <c>brightRoom</c> lighting.</summary>
    BrightRoom,
}
