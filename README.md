# Avig.OpenXml

The package layer under [Avig.Kalkyl](https://github.com/marcuslindh/Avig.Kalkyl) (Excel), Avig.Ord
(Word) and Avig.Bildspel (PowerPoint). It reads and writes Open Packaging Conventions files — the
zip container that .xlsx, .docx and .pptx files are — and Flat OPC, the same package as one XML
file.

It is **not** related to Microsoft's `DocumentFormat.OpenXml` (the Open XML SDK). That package
gives you every element of every Office schema as a class. This one knows nothing about
spreadsheets, documents or presentations: it gives you parts, content types and relationships,
and makes sure a file from outside cannot run your process out of memory while you read them. On
top of that it has what the three packages share: `Length`, the PresentationML and DrawingML
names in `OpcNamespaces`, and DrawingML itself, the shapes, pictures and charts that Office
writes the same way everywhere.

- No dependencies besides Avig.Result; trimming and Native AOT safe.
- Failures you should expect come back as `Result`, not exceptions.
- Every XML part is read without DTD processing and with a limit on how deep elements nest.
- The zip file must say one thing only: every local header as the central directory has it, and
  no bytes before, between or after the entries, so that a scanner and Office cannot be shown two
  different documents (`OpcOpenOptions.StrictZip`, on by default).
- The zip directory is checked before anything is read: the total unpacked size, the number of
  parts, duplicate names (case and Unicode forms folded), names no package may have (`..`, `:`,
  device names such as `CON`, percent-encoded separators), and entries that are encrypted or
  cannot be unpacked. Each part is then held to the size its zip header gave.

## Reading a package

```csharp
Result<OpcPackage> opened = await OpcPackage.OpenAsync("report.docx");

if (opened.IsFailure(out string? error, out _, out OpcPackage? package) == true)
{
    Console.WriteLine(error);

    return;
}

await using (package)
{
    if ((await package.FindMainPartAsync()).IsSuccess(out string? main) == true)
    {
        Result<Stream> part = await package.OpenPartAsync(main);
        // ... read it with package.CreateXmlReader(stream)
    }
}
```

`OpcPackage.OpenAsync` takes an `OpcOpenOptions` with the limits (`MaxUncompressedBytes`,
`MaxParts`, `MaxXmlDepth`) and the wording of the messages (`FormatName`, `LimitsName`), so a
package built on top can give messages that read as its own:

```csharp
OpcOpenOptions options = new() { FormatName = "a .docx document", LimitsName = "OrdLimits" };
// "The file is not a .docx document." / "... the limit set in OrdLimits.MaxUncompressedBytes."
```

Every message about a damaged file starts with `OpcErrors.DamagedPrefix`, and every message about
a limit contains `OpcErrors.LimitMarker`, so the caller can tell the kinds apart.

## What a file points to outside itself

`GetExternalRelationshipsAsync` lists every relationship with `TargetMode="External"`, with the
part it belongs to. Office follows all but hyperlinks on its own when it opens a file: an attached
template on a server, a linked picture or OLE object, a frame, an external workbook. This package
never follows any.

```csharp
foreach (OpcExternalRelationship link in await package.GetExternalRelationshipsAsync())
{
    if (link.IsHyperlink == false)
    {
        Console.WriteLine($"{link.SourcePart}: {link.Kind} -> {link.Target}");
    }
}
```

## Writing a package

```csharp
await using OpcPackageWriter writer = await OpcPackageWriter.CreateAsync(output);

await writer.WritePartAsync("word/document.xml", DocumentContentType, async stream => await WriteBodyAsync(stream));
writer.AddRelationship(string.Empty, OpcNamespaces.OfficeDocumentRelationship, "word/document.xml");

await writer.CompleteAsync();
```

The .rels parts and `[Content_Types].xml` are written by `CompleteAsync`. Pass
`OpcPackageKind.FlatOpc` to `CreateAsync` to write one XML file instead of a zip.

## Lengths

`Length` holds a length exactly, as EMUs (914 400 to the inch), whatever unit it came in:
`Length.Cm(2.5)`, `Length.Pt(11)`, `Length.Inch(1)`, `Length.Px(96)`, `Length.FromEmu(...)`, and
Word's twips and half-points (`FromTwips`, `ToHalfPoints`). Avig.Ord, Avig.Kalkyl and
Avig.Bildspel all use it.

## DrawingML: shapes, pictures and charts

The namespace `Avig.OpenXml.Drawing` holds what Excel, PowerPoint and Word write the same way:

- shapes with a preset geometry (`Shape`, `ShapeKind`, or any of the 187 presets with
  `Shape.FromPreset("star5")` and `Adjustments`);
- fills (`Fill.Solid`, `Fill.Gradient`, `Fill.Pattern`, `Fill.Picture`, `Fill.None`), lines
  (`Outline`, `LineDash`), shadows and soft edges (`Shadow`);
- text in shapes: paragraphs and runs with their own font (`ShapeText`, `TextRun`, `TextFont`,
  `TextAlignment`, `TextAnchor`, margins and wrapping);
- groups, also inside groups (`ShapeGroup`, `GroupItem`), pictures (`Picture`: PNG, JPEG, GIF, BMP and TIFF,
  with the size read from the file's header, and SVG with a PNG for older versions), and a link on
  a shape, a group or a picture (`Hyperlink`);
- places and sizes in `Length` (`DrawingBounds`);
- chart parts, `c:chartSpace` (`ChartSpace`, `ChartSpaceSeries`, `ChartData`, `ChartType`,
  `ChartAxis`, `DataLabels`, `MarkerStyle`, `LegendPosition`): column, bar, line, area, pie,
  doughnut and scatter, stacked, combined and on a second axis, with the series' formulas and the
  values last read from them.

The layer does not know where it is. A spreadsheet drawing puts its shapes in `xdr:` elements
inside cell anchors, a slide in `p:` elements in its `p:spTree`: the package that owns the part
says which with a `ShapeTreeNamespace`, writes what goes around the shapes, and writes the
relationships the shapes need, which a `DrawingContext` collects while they are written.

```csharp
ShapeTreeNamespace slide = new("p", OpcNamespaces.PresentationMain) { ApplicationProperties = true };
DrawingContext context = new(slide, usedRelationshipIds, firstShapeId: 2);

Shape card = new(ShapeKind.RoundedRectangle)
{
    Fill = Fill.Gradient("#2384F9", "#605776", 45),
    Outline = Outline.None,
    Shadow = new Shadow(),
    Text = new ShapeText("Projects", new TextFont { Size = 14, Bold = true, Color = "#FFFFFF" }),
    Link = Hyperlink.ToUrl("https://example.com/projects"),
};

await DrawingWriter.WriteShapeAsync(xml, card, context, new DrawingBounds(Length.Cm(2), Length.Cm(2), Length.Cm(6), Length.Cm(3)));

foreach (DrawingTarget target in context.Targets)
{
    // An image, a link or a chart: the owner writes the part or the relationship with target.Id.
}
```

`DrawingReader.ReadAsync` reads an `sp`, `cxnSp`, `pic`, `grpSp` or `graphicFrame` back into the
same types, whatever its namespace, with a `DrawingReadContext` that gives it the part's
relationships and a way to read an image; `ChartReader.Read` does the same for a chart part.
What is written and read back writes the same XML again. Everything is read by local name, so
Strict OOXML works, and values from a file outside what the types take are clamped, not thrown.
What the types cannot hold, such as a custom geometry, is left out. A theme colour (`a:schemeClr`)
is read when the context is given the theme's colours (`DrawingReadContext.SchemeColors`);
`DrawingColors.TryRead` resolves any DrawingML colour with its modifiers (`lumMod`, `tint`, `shade` …).
A package that has the part and its images in memory reads with `DrawingReader.Read`, without I/O.

The writer is synchronous: `DrawingWriter.WriteShape`, `WritePicture`, `WriteGroup`,
`WriteChartFrame`, `WriteFill`, `WriteOutline`, `WriteEffects`, `WriteRunProperties` and
`ChartWriter.WriteChartSpace` write into any `XmlWriter`, also one without `Async` or one made
with `XContainer.CreateWriter()` to edit a part in memory. The `…Async` forms do the same into
the writer's buffer and return a completed task; the owner of the part flushes it.

## Lower level

`OpcZip`, `OpcXml` and `OpcPath` are the helpers the two classes above use, for a package that
works on a `ZipArchive` itself, as Avig.Kalkyl does.

## License

MIT
