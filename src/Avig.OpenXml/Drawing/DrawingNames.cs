namespace Avig.OpenXml.Drawing;

/// <summary>The names DrawingML stores the enums of this layer as, both ways.</summary>
internal static class DrawingNames
{
    private static readonly (FillPattern Pattern, string Name)[] Patterns =
    [
        (FillPattern.Percent5, "pct5"), (FillPattern.Percent10, "pct10"), (FillPattern.Percent20, "pct20"), (FillPattern.Percent25, "pct25"),
        (FillPattern.Percent30, "pct30"), (FillPattern.Percent40, "pct40"), (FillPattern.Percent50, "pct50"), (FillPattern.Percent60, "pct60"),
        (FillPattern.Percent70, "pct70"), (FillPattern.Percent75, "pct75"), (FillPattern.Percent80, "pct80"), (FillPattern.Percent90, "pct90"),
        (FillPattern.Horizontal, "horz"), (FillPattern.Vertical, "vert"), (FillPattern.LightHorizontal, "ltHorz"), (FillPattern.LightVertical, "ltVert"),
        (FillPattern.DarkHorizontal, "dkHorz"), (FillPattern.DarkVertical, "dkVert"), (FillPattern.NarrowHorizontal, "narHorz"),
        (FillPattern.NarrowVertical, "narVert"), (FillPattern.DashedHorizontal, "dashHorz"), (FillPattern.DashedVertical, "dashVert"),
        (FillPattern.Cross, "cross"), (FillPattern.DownwardDiagonal, "dnDiag"), (FillPattern.UpwardDiagonal, "upDiag"),
        (FillPattern.LightDownwardDiagonal, "ltDnDiag"), (FillPattern.LightUpwardDiagonal, "ltUpDiag"), (FillPattern.DarkDownwardDiagonal, "dkDnDiag"),
        (FillPattern.DarkUpwardDiagonal, "dkUpDiag"), (FillPattern.WideDownwardDiagonal, "wdDnDiag"), (FillPattern.WideUpwardDiagonal, "wdUpDiag"),
        (FillPattern.DashedDownwardDiagonal, "dashDnDiag"), (FillPattern.DashedUpwardDiagonal, "dashUpDiag"), (FillPattern.DiagonalCross, "diagCross"),
        (FillPattern.SmallCheck, "smCheck"), (FillPattern.LargeCheck, "lgCheck"), (FillPattern.SmallGrid, "smGrid"), (FillPattern.LargeGrid, "lgGrid"),
        (FillPattern.DottedGrid, "dotGrid"), (FillPattern.SmallConfetti, "smConfetti"), (FillPattern.LargeConfetti, "lgConfetti"),
        (FillPattern.HorizontalBrick, "horzBrick"), (FillPattern.DiagonalBrick, "diagBrick"), (FillPattern.SolidDiamond, "solidDmnd"),
        (FillPattern.OpenDiamond, "openDmnd"), (FillPattern.DottedDiamond, "dotDmnd"), (FillPattern.Plaid, "plaid"), (FillPattern.Sphere, "sphere"),
        (FillPattern.Weave, "weave"), (FillPattern.Divot, "divot"), (FillPattern.Shingle, "shingle"), (FillPattern.Wave, "wave"),
        (FillPattern.Trellis, "trellis"), (FillPattern.ZigZag, "zigZag"),
    ];

    private static readonly (LineDash Dash, string Name)[] Dashes =
    [
        (LineDash.RoundDot, "sysDot"), (LineDash.SquareDot, "sysDash"), (LineDash.Dash, "dash"), (LineDash.DashDot, "dashDot"),
        (LineDash.LongDash, "lgDash"), (LineDash.LongDashDot, "lgDashDot"), (LineDash.LongDashDotDot, "lgDashDotDot"), (LineDash.Solid, "solid"),
    ];

    private static readonly (MarkerStyle Style, string Name)[] Markers =
    [
        (MarkerStyle.None, "none"), (MarkerStyle.Square, "square"), (MarkerStyle.Diamond, "diamond"), (MarkerStyle.Triangle, "triangle"),
        (MarkerStyle.X, "x"), (MarkerStyle.Star, "star"), (MarkerStyle.Dash, "dash"), (MarkerStyle.Dot, "dot"), (MarkerStyle.Plus, "plus"),
        (MarkerStyle.Circle, "circle"),
    ];

    private static readonly (DataLabelPosition Position, string Name)[] Positions =
    [
        (DataLabelPosition.Center, "ctr"), (DataLabelPosition.InsideEnd, "inEnd"), (DataLabelPosition.InsideBase, "inBase"),
        (DataLabelPosition.OutsideEnd, "outEnd"), (DataLabelPosition.Above, "t"), (DataLabelPosition.Below, "b"),
        (DataLabelPosition.Left, "l"), (DataLabelPosition.Right, "r"), (DataLabelPosition.BestFit, "bestFit"),
    ];

    public static string Pattern(FillPattern pattern)
    {
        return Patterns.First(entry => entry.Pattern == pattern).Name;
    }

    public static FillPattern PatternOf(string? name)
    {
        foreach ((FillPattern pattern, string known) in Patterns)
        {
            if (known == name)
            {
                return pattern;
            }
        }

        return FillPattern.Percent5;
    }

    public static string Dash(LineDash dash)
    {
        foreach ((LineDash known, string name) in Dashes)
        {
            if (known == dash)
            {
                return name;
            }
        }

        return "solid";
    }

    public static LineDash DashOf(string? name)
    {
        foreach ((LineDash dash, string known) in Dashes)
        {
            if (known == name)
            {
                return dash;
            }
        }

        return LineDash.Solid;
    }

    public static string Marker(MarkerStyle style)
    {
        foreach ((MarkerStyle known, string name) in Markers)
        {
            if (known == style)
            {
                return name;
            }
        }

        return "circle";
    }

    public static MarkerStyle MarkerOf(string? name)
    {
        foreach ((MarkerStyle style, string known) in Markers)
        {
            if (known == name)
            {
                return style;
            }
        }

        return MarkerStyle.Circle;
    }

    public static string Position(DataLabelPosition position)
    {
        foreach ((DataLabelPosition known, string name) in Positions)
        {
            if (known == position)
            {
                return name;
            }
        }

        return "bestFit";
    }

    public static DataLabelPosition PositionOf(string? name)
    {
        foreach ((DataLabelPosition position, string known) in Positions)
        {
            if (known == name)
            {
                return position;
            }
        }

        return DataLabelPosition.Default;
    }

    public static string Anchor(TextAnchor anchor)
    {
        return anchor switch
        {
            TextAnchor.Center => "ctr",
            TextAnchor.Bottom => "b",
            _ => "t",
        };
    }

    public static TextAnchor AnchorOf(string? name)
    {
        return name switch
        {
            "ctr" => TextAnchor.Center,
            "b" => TextAnchor.Bottom,
            _ => TextAnchor.Top,
        };
    }

    public static string Alignment(TextAlignment alignment)
    {
        return alignment switch
        {
            TextAlignment.Center => "ctr",
            TextAlignment.Right => "r",
            TextAlignment.Justify => "just",
            TextAlignment.Distributed => "dist",
            _ => "l",
        };
    }

    public static TextAlignment AlignmentOf(string? name)
    {
        return name switch
        {
            "ctr" => TextAlignment.Center,
            "r" => TextAlignment.Right,
            "just" or "justLow" => TextAlignment.Justify,
            "dist" or "thaiDist" => TextAlignment.Distributed,
            _ => TextAlignment.Left,
        };
    }

    public static string Legend(LegendPosition position)
    {
        return position switch
        {
            LegendPosition.Bottom => "b",
            LegendPosition.Top => "t",
            LegendPosition.Left => "l",
            _ => "r",
        };
    }

    public static LegendPosition LegendOf(string? name)
    {
        return name switch
        {
            "b" => LegendPosition.Bottom,
            "t" => LegendPosition.Top,
            "l" => LegendPosition.Left,
            _ => LegendPosition.Right,
        };
    }
}
