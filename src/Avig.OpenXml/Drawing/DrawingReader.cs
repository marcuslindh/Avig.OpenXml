using System.Xml.Linq;

namespace Avig.OpenXml.Drawing;

/// <summary>
/// Reads DrawingML back into the types <see cref="DrawingWriter"/> writes: shapes, connectors,
/// pictures, groups and graphic frames, whatever namespace their shape tree is in (they are
/// read by local name, so Strict OOXML works too). Values outside what the types take are
/// clamped, never thrown, since the file may come from anyone. What the types cannot hold (a
/// theme colour, a custom geometry, a chart inside a group) is left out.
/// </summary>
public static partial class DrawingReader
{
    /// <summary>The most images one element may refer to, so that a hostile group cannot make the reader fetch without end.</summary>
    private const int MaxImages = 4096;

    /// <summary>Reads one element of a shape tree: sp, cxnSp, pic, grpSp or graphicFrame.</summary>
    /// <param name="element">The element.</param>
    /// <param name="context">The part's relationships and images.</param>
    /// <param name="cancellationToken">Cancels the reading of images.</param>
    /// <returns>The item; Invalid when the element is none of the five, or a picture's image cannot be read.</returns>
    public static async Task<Result<ShapeTreeItem>> ReadAsync(XElement element, DrawingReadContext context, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(context);

        if ((element.Name.LocalName is "sp" or "cxnSp" or "pic" or "grpSp" or "graphicFrame") == false)
        {
            return Result<ShapeTreeItem>.Invalid($"<{element.Name.LocalName}> is not an element of a shape tree.");
        }

        Dictionary<string, Picture> images = new(StringComparer.Ordinal);
        Result loaded = await LoadImagesAsync(element, context, images, cancellationToken);

        if (loaded.IsFailure(out string? error, out ResultStatus status) == true)
        {
            return Result<ShapeTreeItem>.Failure(error, status);
        }

        Reader reader = new(context, images);

        if (reader.TryRead(element, out ShapeTreeItem? item, out string? problem) == false)
        {
            return Result<ShapeTreeItem>.Invalid(problem);
        }

        return item;
    }

    /// <summary>
    /// Reads one element of a shape tree as <see cref="ReadAsync"/> does, for a context made with
    /// images that are already in memory (the constructor that takes a synchronous reader).
    /// </summary>
    /// <param name="element">The element.</param>
    /// <param name="context">The part's relationships and images; made with a synchronous image reader.</param>
    /// <returns>The item; Invalid when the element is none of the five, or a picture's image cannot be read.</returns>
    /// <exception cref="InvalidOperationException">The context reads its images asynchronously.</exception>
    public static Result<ShapeTreeItem> Read(XElement element, DrawingReadContext context)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(context);

        if (context.IsSynchronous == false)
        {
            throw new InvalidOperationException("The context reads its images asynchronously; use ReadAsync, or make the context with a synchronous image reader.");
        }

        if ((element.Name.LocalName is "sp" or "cxnSp" or "pic" or "grpSp" or "graphicFrame") == false)
        {
            return Result<ShapeTreeItem>.Invalid($"<{element.Name.LocalName}> is not an element of a shape tree.");
        }

        Dictionary<string, Picture> images = new(StringComparer.Ordinal);
        Result loaded = LoadImages(element, context, images);

        if (loaded.IsFailure(out string? error, out ResultStatus status) == true)
        {
            return Result<ShapeTreeItem>.Failure(error, status);
        }

        Reader reader = new(context, images);

        if (reader.TryRead(element, out ShapeTreeItem? item, out string? problem) == false)
        {
            return Result<ShapeTreeItem>.Invalid(problem);
        }

        return item;
    }

    /// <summary>Reads a fill from the element that holds one (spPr, a chart's c:spPr, a run's rPr), when it has one.</summary>
    /// <param name="parent">The element.</param>
    /// <param name="fill">The fill; a picture fill is not read here, since it needs the part's images.</param>
    /// <returns>True when the element has a fill this layer can hold.</returns>
    public static bool TryReadFill(XElement parent, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Fill? fill)
    {
        return TryReadFill(parent, null, out fill);
    }

    /// <summary>Reads a fill, with theme colours (a:schemeClr) resolved from <paramref name="schemeColors"/> (see <see cref="DrawingColors"/>).</summary>
    /// <param name="parent">The element.</param>
    /// <param name="schemeColors">The theme's colours by name, or null.</param>
    /// <param name="fill">The fill; a picture fill is not read here, since it needs the part's images.</param>
    /// <returns>True when the element has a fill this layer can hold.</returns>
    public static bool TryReadFill(XElement parent, IReadOnlyDictionary<string, string>? schemeColors, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Fill? fill)
    {
        ArgumentNullException.ThrowIfNull(parent);

        return new Reader(null, [], schemeColors).TryFill(parent, out fill);
    }

    /// <summary>Reads the a:ln of an element, when it has one.</summary>
    /// <param name="parent">The element that holds the a:ln.</param>
    /// <param name="outline">The line.</param>
    /// <returns>True when the element has a line this layer can hold.</returns>
    public static bool TryReadOutline(XElement parent, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Outline? outline)
    {
        return TryReadOutline(parent, null, out outline);
    }

    /// <summary>Reads the a:ln of an element, with theme colours resolved from <paramref name="schemeColors"/>.</summary>
    /// <param name="parent">The element that holds the a:ln.</param>
    /// <param name="schemeColors">The theme's colours by name, or null.</param>
    /// <param name="outline">The line.</param>
    /// <returns>True when the element has a line this layer can hold.</returns>
    public static bool TryReadOutline(XElement parent, IReadOnlyDictionary<string, string>? schemeColors, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Outline? outline)
    {
        ArgumentNullException.ThrowIfNull(parent);

        return new Reader(null, [], schemeColors).TryOutline(parent, out outline);
    }

    /// <summary>Reads a run's look from an a:rPr, a:defRPr or a:endParaRPr.</summary>
    /// <param name="properties">The element.</param>
    /// <param name="font">The look.</param>
    /// <returns>True when the element sets anything a <see cref="TextFont"/> holds.</returns>
    public static bool TryReadFont(XElement properties, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out TextFont? font)
    {
        return TryReadFont(properties, null, out font);
    }

    /// <summary>Reads a run's look, with theme colours resolved from <paramref name="schemeColors"/>.</summary>
    /// <param name="properties">The element.</param>
    /// <param name="schemeColors">The theme's colours by name, or null.</param>
    /// <param name="font">The look.</param>
    /// <returns>True when the element sets anything a <see cref="TextFont"/> holds.</returns>
    public static bool TryReadFont(XElement properties, IReadOnlyDictionary<string, string>? schemeColors, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out TextFont? font)
    {
        ArgumentNullException.ThrowIfNull(properties);

        return new Reader(null, [], schemeColors).TryFont(properties, out font);
    }

    /// <summary>Reads the shadows and soft edges of an a:effectLst.</summary>
    /// <param name="effects">The a:effectLst, or null.</param>
    /// <param name="outer">The shadow behind, or null.</param>
    /// <param name="inner">The shadow inside, or null.</param>
    /// <param name="softEdges">Soft edges in points, or null.</param>
    public static void ReadEffects(XElement? effects, out Shadow? outer, out Shadow? inner, out double? softEdges)
    {
        ReadEffects(effects, null, out outer, out inner, out softEdges);
    }

    /// <summary>Reads the shadows and soft edges of an a:effectLst, with theme colours resolved from <paramref name="schemeColors"/>.</summary>
    /// <param name="effects">The a:effectLst, or null.</param>
    /// <param name="schemeColors">The theme's colours by name, or null.</param>
    /// <param name="outer">The shadow behind, or null.</param>
    /// <param name="inner">The shadow inside, or null.</param>
    /// <param name="softEdges">Soft edges in points, or null.</param>
    public static void ReadEffects(XElement? effects, IReadOnlyDictionary<string, string>? schemeColors, out Shadow? outer, out Shadow? inner, out double? softEdges)
    {
        new Reader(null, [], schemeColors).Effects(effects, out outer, out inner, out softEdges);
    }

    /// <summary>Fetches every image the element refers to, through a context that has them in memory.</summary>
    private static Result LoadImages(XElement element, DrawingReadContext context, Dictionary<string, Picture> images)
    {
        foreach (XElement blip in element.Descendants().Where(child => child.Name.LocalName == "blip"))
        {
            if (images.Count >= MaxImages)
            {
                return Result.Invalid($"{OpcErrors.DamagedPrefix}: a drawing refers to more than {MaxImages} images.");
            }

            string? rasterId = Reader.RelationshipId(blip, "embed");

            if (rasterId is null || images.ContainsKey(rasterId) == true)
            {
                continue;
            }

            Result<Picture> raster = ReadPicture(context, rasterId, false, null);

            if (raster.IsFailure(out string? error, out ResultStatus status, out Picture? picture) == true)
            {
                return Result.Failure(error, status);
            }

            images[rasterId] = picture;
            XElement? svgBlip = blip.Descendants().FirstOrDefault(child => child.Name.LocalName == "svgBlip");
            string? svgId = null;

            if (svgBlip is not null)
            {
                svgId = Reader.RelationshipId(svgBlip, "embed");
            }

            if (svgId is not null && images.ContainsKey(Reader.SvgKey(rasterId, svgId)) == false)
            {
                Result<Picture> svg = ReadPicture(context, svgId, true, picture);

                if (svg.IsFailure(out string? svgError, out ResultStatus svgStatus, out Picture? vector) == true)
                {
                    return Result.Failure(svgError, svgStatus);
                }

                images[Reader.SvgKey(rasterId, svgId)] = vector;
            }
        }

        return Result.Success();
    }

    private static Result<Picture> ReadPicture(DrawingReadContext context, string id, bool svg, Picture? fallback)
    {
        if (context.TryGetRelationship(id, out OpcRelationship? relationship) == false || relationship.IsExternal == true)
        {
            return Result<Picture>.Invalid($"{OpcErrors.DamagedPrefix}: the picture's relationship {id} is missing or points outside the package.");
        }

        Result<byte[]> read = context.ReadImage(relationship);

        if (read.IsFailure(out string? error, out ResultStatus status, out byte[]? bytes) == true)
        {
            return Result<Picture>.Failure(error, status);
        }

        if (svg == true)
        {
            return Picture.FromSvg(bytes, fallback!);
        }

        return Picture.FromBytes(bytes);
    }

    /// <summary>Fetches every image the element refers to (a:blip and asvg:svgBlip) before it is read.</summary>
    private static async Task<Result> LoadImagesAsync(XElement element, DrawingReadContext context, Dictionary<string, Picture> images, CancellationToken cancellationToken)
    {
        foreach (XElement blip in element.Descendants().Where(child => child.Name.LocalName == "blip"))
        {
            if (images.Count >= MaxImages)
            {
                return Result.Invalid($"{OpcErrors.DamagedPrefix}: a drawing refers to more than {MaxImages} images.");
            }

            string? rasterId = Reader.RelationshipId(blip, "embed");

            if (rasterId is null || images.ContainsKey(rasterId) == true)
            {
                continue;
            }

            Result<Picture> raster = await ReadPictureAsync(context, rasterId, false, null, cancellationToken);

            if (raster.IsFailure(out string? error, out ResultStatus status, out Picture? picture) == true)
            {
                return Result.Failure(error, status);
            }

            images[rasterId] = picture;

            XElement? svgBlip = blip.Descendants().FirstOrDefault(child => child.Name.LocalName == "svgBlip");
            string? svgId = null;

            if (svgBlip is not null)
            {
                svgId = Reader.RelationshipId(svgBlip, "embed");
            }

            if (svgId is not null && images.ContainsKey(Reader.SvgKey(rasterId, svgId)) == false)
            {
                Result<Picture> svg = await ReadPictureAsync(context, svgId, true, picture, cancellationToken);

                if (svg.IsFailure(out string? svgError, out ResultStatus svgStatus, out Picture? vector) == true)
                {
                    return Result.Failure(svgError, svgStatus);
                }

                images[Reader.SvgKey(rasterId, svgId)] = vector;
            }
        }

        return Result.Success();
    }

    private static async Task<Result<Picture>> ReadPictureAsync(DrawingReadContext context, string id, bool svg, Picture? fallback, CancellationToken cancellationToken)
    {
        if (context.TryGetRelationship(id, out OpcRelationship? relationship) == false || relationship.IsExternal == true)
        {
            return Result<Picture>.Invalid($"{OpcErrors.DamagedPrefix}: the picture's relationship {id} is missing or points outside the package.");
        }

        Result<byte[]> read = await context.ReadImageAsync(relationship, cancellationToken);

        if (read.IsFailure(out string? error, out ResultStatus status, out byte[]? bytes) == true)
        {
            return Result<Picture>.Failure(error, status);
        }

        if (svg == true)
        {
            return Picture.FromSvg(bytes, fallback!);
        }

        return Picture.FromBytes(bytes);
    }

    /// <summary>The parsing itself, with the images already fetched, so that it needs no I/O.</summary>
    private sealed partial class Reader
    {
        private readonly DrawingReadContext? context;
        private readonly Dictionary<string, Picture> images;
        private readonly IReadOnlyDictionary<string, string>? schemeColors;

        public Reader(DrawingReadContext? context, Dictionary<string, Picture> images, IReadOnlyDictionary<string, string>? schemeColors = null)
        {
            this.context = context;
            this.images = images;
            this.schemeColors = schemeColors ?? context?.SchemeColors;
        }

        public static string SvgKey(string rasterId, string svgId)
        {
            return rasterId + "|" + svgId;
        }

        /// <summary>An r:embed or r:id, by local name in whatever namespace (Transitional or Strict), or null.</summary>
        public static string? RelationshipId(XElement? element, string localName)
        {
            XAttribute? found = element?.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == localName && attribute.Name.NamespaceName.Length > 0);

            return found?.Value;
        }

        public static XElement? Child(XElement? parent, string localName)
        {
            return parent?.Elements().FirstOrDefault(child => child.Name.LocalName == localName);
        }

        public static IEnumerable<XElement> Children(XElement? parent, string localName)
        {
            if (parent is null)
            {
                return [];
            }

            return parent.Elements().Where(child => child.Name.LocalName == localName);
        }

        public bool TryRead(XElement element, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out ShapeTreeItem? item, [System.Diagnostics.CodeAnalysis.NotNullWhen(false)] out string? problem)
        {
            problem = null;
            item = null;

            switch (element.Name.LocalName)
            {
                case "sp" or "cxnSp":
                    item = ReadShape(element);

                    return true;

                case "pic":
                    if (TryReadPicture(element, out item) == false)
                    {
                        problem = $"{OpcErrors.DamagedPrefix}: a picture has no image.";

                        return false;
                    }

                    return true;

                case "grpSp":
                    if (TryReadGroup(element, out item) == false)
                    {
                        problem = $"{OpcErrors.DamagedPrefix}: a group holds nothing that can be read.";

                        return false;
                    }

                    return true;

                default:
                    item = ReadGraphicFrame(element);

                    return true;
            }
        }

        private ShapeTreeItem ReadShape(XElement element)
        {
            XElement? properties = Child(element, "spPr");
            XElement? nonVisual = Child(Child(element, "nvSpPr") ?? Child(element, "nvCxnSpPr"), "cNvPr");
            XElement? geometry = Child(properties, "prstGeom");
            string preset = geometry?.Attribute("prst")?.Value ?? "rect";

            if (PresetGeometry.IsKnown(preset) == false)
            {
                preset = "rect";
            }

            bool textBox = DrawingValues.Boolean(Child(Child(element, "nvSpPr"), "cNvSpPr")?.Attribute("txBox")?.Value, false);
            Shape shape;

            if (textBox == true && preset == "rect")
            {
                shape = new Shape(ShapeKind.TextBox);
            }
            else
            {
                shape = Shape.FromPreset(preset);
            }

            List<ShapeAdjustment> adjustments = [];
            double? cornerRadius = null;

            foreach (XElement guide in Children(Child(geometry, "avLst"), "gd"))
            {
                string? name = guide.Attribute("name")?.Value;
                string formula = guide.Attribute("fmla")?.Value ?? string.Empty;

                if (name is null || name.Length == 0 || name.Length > 64 || name.All(char.IsAsciiLetterOrDigit) == false || adjustments.Count >= 16
                    || formula.StartsWith("val ", StringComparison.Ordinal) == false || DrawingValues.TryLong(formula[4..], out long value) == false)
                {
                    continue;
                }

                if (shape.Kind == ShapeKind.RoundedRectangle && name == "adj" && cornerRadius is null)
                {
                    cornerRadius = Math.Clamp(value / 100_000.0, 0, 0.5);

                    continue;
                }

                adjustments.Add(new ShapeAdjustment(name, value));
            }

            XElement? transform = Child(properties, "xfrm");
            Fill? fill = null;
            Outline? outline = null;

            if (properties is not null && TryFill(properties, out Fill? readFill) == true)
            {
                fill = readFill;
            }

            if (properties is not null && TryOutline(properties, out Outline? readOutline) == true)
            {
                outline = readOutline;
            }

            XElement? effects = Child(properties, "effectLst");
            Effects(effects, out Shadow? outer, out Shadow? inner, out double? softEdges);
            MoreEffects(effects, out GlowEffect? glow, out ReflectionEffect? reflection, out BlurEffect? blur, out PresetShadow? presetShadow, out FillOverlayEffect? overlay);
            ThreeD(properties, out Scene3D? scene, out Shape3D? threeD);
            XElement? bodyProperties = Child(Child(element, "txBody"), "bodyPr");

            if (scene is not null && Child(bodyProperties, "flatTx") is not null)
            {
                scene = scene with { KeepTextFlat = true };
            }

            if (Child(properties, "custGeom") is XElement custom)
            {
                if (TryCustomGeometry(custom, out CustomGeometry? read) == true)
                {
                    shape = Shape.FromGeometry(read);
                    adjustments.Clear();
                    cornerRadius = null;
                }
            }

            if (DrawingValues.Boolean(element.Attribute("useBgFill")?.Value, false) == true)
            {
                fill = Fill.Background;
            }

            XElement? connectorProperties = Child(Child(element, "nvCxnSpPr"), "cNvCxnSpPr");
            ShapeText? text = null;

            if (shape.IsConnector == false && TryText(Child(element, "txBody"), out ShapeText? readText) == true)
            {
                text = readText;
            }

            TryStyle(Child(element, "style"), out ShapeStyle? style);

            shape = shape with
            {
                Glow = glow,
                Reflection = reflection,
                Blur = blur,
                PresetShadow = presetShadow,
                FillOverlay = overlay,
                Scene3D = scene,
                Shape3D = threeD,
                StartConnection = Connection(Child(connectorProperties, "stCxn")),
                EndConnection = Connection(Child(connectorProperties, "endCxn")),
                Locks = Locks(Child(Child(Child(element, "nvSpPr"), "cNvSpPr"), "spLocks") ?? Child(connectorProperties, "cxnSpLocks")),
                IsHidden = DrawingValues.Boolean(nonVisual?.Attribute("hidden")?.Value, false),
                IsDecorative = IsDecorative(nonVisual),
                Title = DrawingText(nonVisual?.Attribute("title")?.Value),
            };

            // The style Office gives a new shape of the kind is no style of its own.
            if (style is not null && style != DrawingWriter.DefaultStyle(shape))
            {
                shape = shape with { Style = style };
            }

            shape = shape with
            {
                CornerRadius = cornerRadius,
                Adjustments = adjustments,
                Fill = fill,
                Outline = outline,
                Shadow = outer,
                InnerShadow = inner,
                SoftEdges = softEdges,
                Text = text,
                Rotation = Rotation(transform),
                FlipHorizontal = DrawingValues.Boolean(transform?.Attribute("flipH")?.Value, false),
                FlipVertical = DrawingValues.Boolean(transform?.Attribute("flipV")?.Value, false),
                Link = Link(nonVisual),
                Name = DrawingText(nonVisual?.Attribute("name")?.Value),
                Description = DrawingText(nonVisual?.Attribute("descr")?.Value),
            };

            return new ShapeTreeItem(ShapeTreeItemKind.Shape, Id(nonVisual), nonVisual?.Attribute("name")?.Value ?? string.Empty, Bounds(transform))
            {
                Shape = shape,
                Description = shape.Description,
            };
        }

        private bool TryReadPicture(XElement element, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out ShapeTreeItem? item)
        {
            item = null;
            XElement? nonVisual = Child(Child(element, "nvPicPr"), "cNvPr");
            XElement? blip = Child(Child(element, "blipFill"), "blip");

            if (TryBlip(blip, out Picture? picture) == false)
            {
                return false;
            }

            XElement? properties = Child(element, "spPr");
            double transparency = 0;

            if (DrawingValues.TryLong(Child(blip, "alphaModFix")?.Attribute("amt")?.Value, out long amount) == true)
            {
                transparency = Math.Clamp(1 - amount / 100_000.0, 0, 1);
            }

            Effects(Child(properties, "effectLst"), out Shadow? outer, out _, out _);
            picture = picture with
            {
                Description = DrawingText(nonVisual?.Attribute("descr")?.Value),
                Link = Link(nonVisual),
                Transparency = transparency,
                Shadow = outer,
            };

            item = new ShapeTreeItem(ShapeTreeItemKind.Picture, Id(nonVisual), nonVisual?.Attribute("name")?.Value ?? string.Empty, Bounds(Child(properties, "xfrm")))
            {
                Picture = picture,
                Description = picture.Description,
            };

            return true;
        }

        private bool TryReadGroup(XElement element, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out ShapeTreeItem? item)
        {
            item = null;
            XElement? nonVisual = Child(Child(element, "nvGrpSpPr"), "cNvPr");
            List<GroupItem> items = [];

            foreach (XElement child in element.Elements())
            {
                if (items.Count >= 1000)
                {
                    break;
                }

                if ((child.Name.LocalName is "sp" or "cxnSp" or "pic" or "grpSp") == false || TryRead(child, out ShapeTreeItem? read, out _) == false)
                {
                    continue;
                }

                if (read.Shape is not null)
                {
                    items.Add(GroupItem.Of(read.Shape, read.Bounds));
                }
                else if (read.Picture is not null)
                {
                    items.Add(GroupItem.Of(read.Picture, read.Bounds));
                }
                else if (read.Group is not null)
                {
                    items.Add(GroupItem.Of(read.Group, read.Bounds));
                }
            }

            if (items.Count == 0)
            {
                return false;
            }

            XElement? groupTransform = Child(Child(element, "grpSpPr"), "xfrm");
            ShapeGroup group = new(items)
            {
                Link = Link(nonVisual),
                Name = DrawingText(nonVisual?.Attribute("name")?.Value),
                Description = DrawingText(nonVisual?.Attribute("descr")?.Value),
                Rotation = Rotation(groupTransform),
                FlipHorizontal = DrawingValues.Boolean(groupTransform?.Attribute("flipH")?.Value, false),
                FlipVertical = DrawingValues.Boolean(groupTransform?.Attribute("flipV")?.Value, false),
                Locks = Locks(Child(Child(Child(element, "nvGrpSpPr"), "cNvGrpSpPr"), "grpSpLocks")),
                IsHidden = DrawingValues.Boolean(nonVisual?.Attribute("hidden")?.Value, false),
            };

            item = new ShapeTreeItem(ShapeTreeItemKind.Group, Id(nonVisual), nonVisual?.Attribute("name")?.Value ?? string.Empty, Bounds(Child(Child(element, "grpSpPr"), "xfrm")))
            {
                Group = group,
                Description = group.Description,
            };

            return true;
        }

        private static ShapeTreeItem ReadGraphicFrame(XElement element)
        {
            XElement? nonVisual = Child(Child(element, "nvGraphicFramePr"), "cNvPr");
            XElement? data = Child(Child(element, "graphic"), "graphicData");
            string? uri = data?.Attribute("uri")?.Value;
            string? chartId = null;

            if (uri == OpcNamespaces.DrawingChart)
            {
                chartId = RelationshipId(Child(data, "chart"), "id");
            }

            return new ShapeTreeItem(ShapeTreeItemKind.GraphicFrame, Id(nonVisual), nonVisual?.Attribute("name")?.Value ?? string.Empty, Bounds(Child(element, "xfrm")))
            {
                Description = DrawingText(nonVisual?.Attribute("descr")?.Value),
                GraphicDataUri = uri,
                ChartRelationshipId = chartId,
            };
        }

        private bool TryBlip(XElement? blip, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Picture? picture)
        {
            picture = null;

            if (blip is null)
            {
                return false;
            }

            string? rasterId = RelationshipId(blip, "embed");

            if (rasterId is null || images.TryGetValue(rasterId, out Picture? raster) == false)
            {
                return false;
            }

            picture = raster;
            XElement? svgBlip = blip.Descendants().FirstOrDefault(child => child.Name.LocalName == "svgBlip");
            string? svgId = null;

            if (svgBlip is not null)
            {
                svgId = RelationshipId(svgBlip, "embed");
            }

            if (svgId is not null && images.TryGetValue(SvgKey(rasterId, svgId), out Picture? vector) == true)
            {
                picture = vector;
            }

            return true;
        }

        public bool TryFill(XElement parent, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Fill? fill)
        {
            fill = null;

            foreach (XElement child in parent.Elements())
            {
                switch (child.Name.LocalName)
                {
                    case "noFill":
                        fill = Fill.None;

                        return true;

                    case "solidFill":
                        if (TryColor(child, out string? color, out double transparency) == true)
                        {
                            fill = Fill.Solid(color, transparency);

                            return true;
                        }

                        return false;

                    case "gradFill":
                        return TryGradient(child, out fill);

                    case "grpFill":
                        fill = Fill.Group;

                        return true;

                    case "pattFill":
                        if (TryColor(Child(child, "fgClr"), out string? foreground, out _) == true && TryColor(Child(child, "bgClr"), out string? background, out _) == true)
                        {
                            fill = Fill.Pattern(DrawingNames.PatternOf(child.Attribute("prst")?.Value), foreground, background);

                            return true;
                        }

                        return false;

                    case "blipFill":
                        if (TryBlip(Child(child, "blip"), out Picture? image) == true)
                        {
                            fill = PictureDetails(child, Fill.Picture(image));

                            return true;
                        }

                        return false;
                }
            }

            return false;
        }

        private bool TryGradient(XElement element, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Fill? fill)
        {
            fill = null;
            List<GradientStop> stops = [];

            foreach (XElement stop in Children(Child(element, "gsLst"), "gs"))
            {
                if (stops.Count >= 10)
                {
                    break;
                }

                if (DrawingValues.TryLong(stop.Attribute("pos")?.Value, out long position) == true && TryColor(stop, out string? color, out double transparency) == true)
                {
                    stops.Add(new GradientStop(Math.Clamp(position / 100_000.0, 0, 1), color, transparency));
                }
            }

            if (stops.Count < 2)
            {
                return false;
            }

            double angle = 0;

            if (DrawingValues.TryLong(Child(element, "lin")?.Attribute("ang")?.Value, out long rawAngle) == true)
            {
                angle = Math.Clamp(rawAngle / 60_000.0, -360, 360);
            }

            fill = GradientDetails(element, Fill.Gradient(angle, [.. stops]));

            return true;
        }

        /// <summary>An a:srgbClr (or a:sysClr with its last colour) inside the element, with its alpha as a transparency.</summary>
        public bool TryColor(XElement? parent, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? color, out double transparency)
        {
            if (schemeColors is not null)
            {
                return DrawingColors.TryRead(parent, schemeColors, out color, out transparency);
            }

            color = null;
            transparency = 0;
            XElement? element = Child(parent, "srgbClr");
            string? value = element?.Attribute("val")?.Value;

            if (element is null)
            {
                element = Child(parent, "sysClr");
                value = element?.Attribute("lastClr")?.Value;
            }

            if (element is null || value is null || value.Length != 6 || value.All(char.IsAsciiHexDigit) == false)
            {
                return false;
            }

            color = "#" + value.ToUpperInvariant();

            if (DrawingValues.TryLong(Child(element, "alpha")?.Attribute("val")?.Value, out long alpha) == true)
            {
                transparency = Math.Clamp(1 - alpha / 100_000.0, 0, 1);
            }

            return true;
        }

        public bool TryOutline(XElement parent, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Outline? outline)
        {
            outline = null;
            XElement? line = Child(parent, "ln");

            if (line is null)
            {
                return false;
            }

            if (Child(line, "noFill") is not null)
            {
                outline = Outline.None;

                return true;
            }

            double width = 0.75;

            if (DrawingValues.TryLong(line.Attribute("w")?.Value, out long emu) == true)
            {
                width = Math.Clamp(DrawingValues.Points(emu), 0, 1584);
            }

            if (TryColor(Child(line, "solidFill"), out string? color, out double transparency) == true)
            {
                outline = new Outline(color, width) { Transparency = transparency };
            }
            else if ((Child(line, "gradFill") is not null || Child(line, "pattFill") is not null) && TryFill(line, out Fill? lineFill) == true && lineFill.Kind is FillKind.Gradient or FillKind.Pattern)
            {
                outline = new Outline(lineFill, width);
            }
            else
            {
                return false;
            }

            outline = LineDetails(line, outline with { Dash = DrawingNames.DashOf(Child(line, "prstDash")?.Attribute("val")?.Value) });

            return true;
        }

        public void Effects(XElement? effects, out Shadow? outer, out Shadow? inner, out double? softEdges)
        {
            outer = null;
            inner = null;
            softEdges = null;

            if (Child(effects, "outerShdw") is XElement outerShadow && TryShadow(outerShadow, out Shadow? readOuter) == true)
            {
                outer = ShadowDetails(outerShadow, readOuter);
            }

            if (Child(effects, "innerShdw") is XElement innerShadow && TryShadow(innerShadow, out Shadow? readInner) == true)
            {
                inner = readInner;
            }

            if (DrawingValues.TryLong(Child(effects, "softEdge")?.Attribute("rad")?.Value, out long radius) == true)
            {
                softEdges = Math.Clamp(DrawingValues.Points(radius), 0, 100);
            }
        }

        private bool TryShadow(XElement element, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Shadow? shadow)
        {
            shadow = null;

            if (TryColor(element, out string? color, out double transparency) == false)
            {
                return false;
            }

            shadow = new Shadow
            {
                Color = color,
                Transparency = transparency,
                Blur = Math.Clamp(DrawingValues.Points(Long(element.Attribute("blurRad")?.Value)), 0, 100),
                Distance = Math.Clamp(DrawingValues.Points(Long(element.Attribute("dist")?.Value)), 0, 200),
                Angle = Math.Clamp(Long(element.Attribute("dir")?.Value) / 60_000.0, -360, 360),
            };

            return true;
        }

        /// <summary>
        /// The text of a txBody. The paragraphs become runs with a <c>\n</c> between them, each
        /// in the look of the end of the paragraph before; runs that look the same are joined.
        /// No text at all, and no wrap set, is no text: that is how an empty shape is written.
        /// </summary>
        private bool TryText(XElement? body, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out ShapeText? text)
        {
            text = null;

            if (body is null)
            {
                return false;
            }

            XElement? bodyProperties = Child(body, "bodyPr");
            List<TextRun> runs = [];
            TextAlignment alignment = TextAlignment.Left;
            bool first = true;
            TextFont? previousEnd = null;
            int length = 0;

            foreach (XElement paragraph in Children(body, "p"))
            {
                if (first == true)
                {
                    alignment = DrawingNames.AlignmentOf(Child(paragraph, "pPr")?.Attribute("algn")?.Value);
                }
                else
                {
                    Append(runs, "\n", previousEnd, ref length);
                }

                TextFont? last = null;

                foreach (XElement run in paragraph.Elements().Where(child => child.Name.LocalName is "r" or "fld" or "br"))
                {
                    TextFont? font = null;

                    if (Child(run, "rPr") is XElement properties && TryFont(properties, out TextFont? readFont) == true)
                    {
                        font = readFont;
                    }

                    // A line break inside a paragraph becomes a new paragraph, the only break the types have.
                    string value = "\n";

                    if (run.Name.LocalName != "br")
                    {
                        value = Child(run, "t")?.Value ?? string.Empty;
                    }

                    Append(runs, value, font, ref length);
                    last = font;
                }

                TextFont? end = last;

                if (Child(paragraph, "endParaRPr") is XElement endProperties && TryFont(endProperties, out TextFont? endFont) == true)
                {
                    end = endFont;
                }

                previousEnd = end;
                first = false;
            }

            bool hasText = runs.Exists(run => run.Text.Length > 0 && run.Text != "\n");

            if (hasText == false && bodyProperties?.Attribute("wrap") is null)
            {
                return false;
            }

            if (runs.Count == 0)
            {
                runs.Add(new TextRun(string.Empty));
            }

            text = new ShapeText(runs)
            {
                Horizontal = alignment,
                Vertical = DrawingNames.AnchorOf(bodyProperties?.Attribute("anchor")?.Value),
                Wrap = bodyProperties?.Attribute("wrap")?.Value != "none",
                MarginLeft = Inset(bodyProperties, "lIns"),
                MarginTop = Inset(bodyProperties, "tIns"),
                MarginRight = Inset(bodyProperties, "rIns"),
                MarginBottom = Inset(bodyProperties, "bIns"),
            };

            return true;
        }

        /// <summary>Adds text to the last run when it looks the same, or as a new run; stops at the most a run holds.</summary>
        private static void Append(List<TextRun> runs, string value, TextFont? font, ref int length)
        {
            if (value.Length == 0 && runs.Count > 0)
            {
                return;
            }

            string kept = value[..Math.Min(value.Length, Math.Max(0, DrawingValues.MaxTextLength - length))];
            kept = new string(kept.Where(System.Xml.XmlConvert.IsXmlChar).ToArray());
            length += kept.Length;

            if (runs.Count > 0 && runs[^1].Font == font)
            {
                runs[^1] = new TextRun(runs[^1].Text + kept, font);

                return;
            }

            runs.Add(new TextRun(kept, font));
        }

        private static double? Inset(XElement? bodyProperties, string name)
        {
            double? points = null;

            if (DrawingValues.TryLong(bodyProperties?.Attribute(name)?.Value, out long emu) == true)
            {
                points = Math.Clamp(DrawingValues.Points(emu), 0, 1000);
            }

            return points;
        }

        public bool TryFont(XElement properties, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out TextFont? font)
        {
            font = null;
            string? size = properties.Attribute("sz")?.Value;
            string? bold = properties.Attribute("b")?.Value;
            string? italic = properties.Attribute("i")?.Value;
            string? typeface = Child(properties, "latin")?.Attribute("typeface")?.Value;
            bool colored = TryColor(Child(properties, "solidFill"), out string? color, out _);

            if (size is null && bold is null && italic is null && typeface is null && colored == false)
            {
                return false;
            }

            double? points = null;

            if (DrawingValues.TryLong(size, out long hundredths) == true)
            {
                points = Math.Clamp(hundredths / 100.0, 1, 400);
            }

            if (typeface is not null && (typeface.Length == 0 || typeface.Length > 31))
            {
                typeface = null;
            }

            font = new TextFont
            {
                Size = points,
                Bold = DrawingValues.Boolean(bold, false),
                Italic = DrawingValues.Boolean(italic, false),
                Color = color,
                Name = typeface,
            };

            return true;
        }

        /// <summary>The a:hlinkClick of a cNvPr, or null when there is none or its relationship is missing.</summary>
        private Hyperlink? Link(XElement? nonVisual)
        {
            XElement? click = Child(nonVisual, "hlinkClick");
            Hyperlink? link = null;

            if (click is null || context is null)
            {
                return link;
            }

            string? action = click.Attribute("action")?.Value;
            string? tooltip = click.Attribute("tooltip")?.Value;

            if (context.TryGetRelationship(RelationshipId(click, "id"), out OpcRelationship? relationship) == true)
            {
                if (relationship.IsExternal == true)
                {
                    link = Hyperlink.Read(relationship.Target, null);
                }
                else
                {
                    link = Hyperlink.Read(null, relationship.Target.TrimStart('#'));
                }
            }
            else if (action is not null)
            {
                // An action without a target, such as "next slide".
                link = Hyperlink.Read(null, string.Empty);
            }

            if (link is not null && tooltip is not null && tooltip.Length <= 255)
            {
                link = link with { Tooltip = Clean(tooltip) };
            }

            if (link is not null && action is not null && action.Length > 0 && action.Length <= 2079)
            {
                link = link with { Action = Clean(action) };
            }

            if (link is not null)
            {
                link = link with
                {
                    HighlightClick = DrawingValues.Boolean(click.Attribute("highlightClick")?.Value, false),
                    StopsSound = DrawingValues.Boolean(click.Attribute("endSnd")?.Value, false),
                };
            }

            return link;
        }

        private static int Id(XElement? nonVisual)
        {
            if (int.TryParse(nonVisual?.Attribute("id")?.Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int id) == true)
            {
                return id;
            }

            return 0;
        }

        private static double Rotation(XElement? transform)
        {
            return Math.Clamp(Long(transform?.Attribute("rot")?.Value) / 60_000.0, -360, 360);
        }

        public static DrawingBounds Bounds(XElement? transform)
        {
            XElement? offset = Child(transform, "off");
            XElement? extent = Child(transform, "ext");

            return DrawingBounds.Read(
                Coordinate(offset?.Attribute("x")?.Value),
                Coordinate(offset?.Attribute("y")?.Value),
                Coordinate(extent?.Attribute("cx")?.Value),
                Coordinate(extent?.Attribute("cy")?.Value));
        }

        private static long Coordinate(string? text)
        {
            return Math.Clamp(Long(text), -DrawingValues.MaxCoordinate, DrawingValues.MaxCoordinate);
        }

        private static long Long(string? text)
        {
            if (DrawingValues.TryLong(text, out long value) == true)
            {
                return value;
            }

            return 0;
        }

        /// <summary>A name or alternative text as the types take it: at most 1000 characters, nothing XML cannot hold.</summary>
        private static string? DrawingText(string? value)
        {
            string? text = value;

            if (value is not null)
            {
                text = Clean(value[..Math.Min(value.Length, 1000)]);
            }

            return text;
        }

        private static string Clean(string value)
        {
            return new string(value.Where(System.Xml.XmlConvert.IsXmlChar).ToArray());
        }
    }
}
