namespace Avig.OpenXml.Drawing;

/// <summary>DrawingML's preset geometries (ST_ShapeType) and the names of the ones <see cref="ShapeKind"/> has.</summary>
internal static class PresetGeometry
{
    private static readonly HashSet<string> Known = new(StringComparer.Ordinal)
    {
        "line", "lineInv", "triangle", "rtTriangle", "rect", "diamond", "parallelogram", "trapezoid", "nonIsoscelesTrapezoid",
        "pentagon", "hexagon", "heptagon", "octagon", "decagon", "dodecagon", "star4", "star5", "star6", "star7", "star8",
        "star10", "star12", "star16", "star24", "star32", "roundRect", "round1Rect", "round2SameRect", "round2DiagRect",
        "snipRoundRect", "snip1Rect", "snip2SameRect", "snip2DiagRect", "plaque", "ellipse", "teardrop", "homePlate", "chevron",
        "pieWedge", "pie", "blockArc", "donut", "noSmoking", "rightArrow", "leftArrow", "upArrow", "downArrow", "stripedRightArrow",
        "notchedRightArrow", "bentUpArrow", "leftRightArrow", "upDownArrow", "leftUpArrow", "leftRightUpArrow", "quadArrow",
        "leftArrowCallout", "rightArrowCallout", "upArrowCallout", "downArrowCallout", "leftRightArrowCallout", "upDownArrowCallout",
        "quadArrowCallout", "bentArrow", "uturnArrow", "circularArrow", "leftCircularArrow", "leftRightCircularArrow",
        "curvedRightArrow", "curvedLeftArrow", "curvedUpArrow", "curvedDownArrow", "swooshArrow", "cube", "can", "lightningBolt",
        "heart", "sun", "moon", "smileyFace", "irregularSeal1", "irregularSeal2", "foldedCorner", "bevel", "frame", "halfFrame",
        "corner", "diagStripe", "chord", "arc", "leftBracket", "rightBracket", "leftBrace", "rightBrace", "bracketPair", "bracePair",
        "straightConnector1", "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5", "curvedConnector2",
        "curvedConnector3", "curvedConnector4", "curvedConnector5", "callout1", "callout2", "callout3", "accentCallout1",
        "accentCallout2", "accentCallout3", "borderCallout1", "borderCallout2", "borderCallout3", "accentBorderCallout1",
        "accentBorderCallout2", "accentBorderCallout3", "wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout",
        "cloudCallout", "cloud", "ribbon", "ribbon2", "ellipseRibbon", "ellipseRibbon2", "leftRightRibbon", "verticalScroll",
        "horizontalScroll", "wave", "doubleWave", "plus", "flowChartProcess", "flowChartDecision", "flowChartInputOutput",
        "flowChartPredefinedProcess", "flowChartInternalStorage", "flowChartDocument", "flowChartMultidocument",
        "flowChartTerminator", "flowChartPreparation", "flowChartManualInput", "flowChartManualOperation", "flowChartConnector",
        "flowChartPunchedCard", "flowChartPunchedTape", "flowChartSummingJunction", "flowChartOr", "flowChartCollate",
        "flowChartSort", "flowChartExtract", "flowChartMerge", "flowChartOfflineStorage", "flowChartOnlineStorage",
        "flowChartMagneticTape", "flowChartMagneticDisk", "flowChartMagneticDrum", "flowChartDisplay", "flowChartDelay",
        "flowChartAlternateProcess", "flowChartOffpageConnector", "actionButtonBlank", "actionButtonHome", "actionButtonHelp",
        "actionButtonInformation", "actionButtonForwardNext", "actionButtonBackPrevious", "actionButtonEnd",
        "actionButtonBeginning", "actionButtonReturn", "actionButtonDocument", "actionButtonSound", "actionButtonMovie", "gear6",
        "gear9", "funnel", "mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual", "mathNotEqual", "cornerTabs",
        "squareTabs", "plaqueTabs", "chartX", "chartStar", "chartPlus",
    };

    /// <summary>The names in the schema's order.</summary>
    public static IReadOnlyList<string> Names { get; } = [.. Known];

    /// <summary>True for a name in ST_ShapeType.</summary>
    public static bool IsKnown(string preset)
    {
        return Known.Contains(preset);
    }

    /// <summary>True for the presets Office draws as a connector (cxnSp): the line and the straight, elbow and curved connectors.</summary>
    public static bool IsConnector(string preset)
    {
        return preset is "line" or "straightConnector1" or "bentConnector2" or "bentConnector3" or "bentConnector4" or "bentConnector5"
            or "curvedConnector2" or "curvedConnector3" or "curvedConnector4" or "curvedConnector5";
    }

    /// <summary>The preset a kind is drawn with.</summary>
    public static string NameOf(ShapeKind kind)
    {
        return kind switch
        {
            ShapeKind.RoundedRectangle => "roundRect",
            ShapeKind.Ellipse => "ellipse",
            ShapeKind.Triangle => "triangle",
            ShapeKind.Diamond => "diamond",
            ShapeKind.Hexagon => "hexagon",
            ShapeKind.RightArrow => "rightArrow",
            ShapeKind.LeftArrow => "leftArrow",
            ShapeKind.UpArrow => "upArrow",
            ShapeKind.DownArrow => "downArrow",
            ShapeKind.Chevron => "chevron",
            ShapeKind.Pentagon => "homePlate",
            ShapeKind.Line => "line",
            _ => "rect",
        };
    }

    /// <summary>The kind a preset has a name as; a text box is a rect with txBox="1", which the reader decides.</summary>
    public static bool TryKindOf(string preset, out ShapeKind kind)
    {
        kind = preset switch
        {
            "rect" => ShapeKind.Rectangle,
            "roundRect" => ShapeKind.RoundedRectangle,
            "ellipse" => ShapeKind.Ellipse,
            "triangle" => ShapeKind.Triangle,
            "diamond" => ShapeKind.Diamond,
            "hexagon" => ShapeKind.Hexagon,
            "rightArrow" => ShapeKind.RightArrow,
            "leftArrow" => ShapeKind.LeftArrow,
            "upArrow" => ShapeKind.UpArrow,
            "downArrow" => ShapeKind.DownArrow,
            "chevron" => ShapeKind.Chevron,
            "homePlate" => ShapeKind.Pentagon,
            "line" => ShapeKind.Line,
            _ => ShapeKind.TextBox,
        };

        return kind != ShapeKind.TextBox;
    }

    /// <summary>The name Office gives a new shape of a kind, before its number: "Oval 3".</summary>
    public static string DefaultName(Shape shape)
    {
        if (shape.Geometry is not null)
        {
            return "Freeform: Shape";
        }

        if (shape.Preset.StartsWith("bentConnector", StringComparison.Ordinal) == true)
        {
            return "Connector: Elbow";
        }

        if (shape.Preset.StartsWith("curvedConnector", StringComparison.Ordinal) == true)
        {
            return "Connector: Curved";
        }

        if (shape.Preset == "straightConnector1")
        {
            return "Straight Arrow Connector";
        }

        return shape.Kind switch
        {
            ShapeKind.RoundedRectangle => "Rectangle: Rounded Corners",
            ShapeKind.Ellipse => "Oval",
            ShapeKind.Triangle => "Isosceles Triangle",
            ShapeKind.RightArrow => "Arrow: Right",
            ShapeKind.LeftArrow => "Arrow: Left",
            ShapeKind.UpArrow => "Arrow: Up",
            ShapeKind.DownArrow => "Arrow: Down",
            ShapeKind.Chevron => "Arrow: Chevron",
            ShapeKind.Pentagon => "Arrow: Pentagon",
            ShapeKind.Line => "Straight Connector",
            ShapeKind.TextBox => "TextBox",
            ShapeKind kind => kind.ToString(),
            null => "Shape",
        };
    }
}
