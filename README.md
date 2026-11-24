# Avig.OpenXml

The package layer under [Avig.Kalkyl](https://github.com/marcuslindh/Avig.Kalkyl) (Excel) and
Avig.Ord (Word). It reads and writes Open Packaging Conventions files — the zip container that
.xlsx, .docx and .pptx files are — and Flat OPC, the same package as one XML file.

It is **not** related to Microsoft's `DocumentFormat.OpenXml` (the Open XML SDK). That package
gives you every element of every Office schema as a class. This one knows nothing about
spreadsheets or documents: it gives you parts, content types and relationships, and makes sure a
file from outside cannot run your process out of memory while you read them.

- No dependencies besides Avig.Result; trimming and Native AOT safe.
- Failures you should expect come back as `Result`, not exceptions.
- Every XML part is read without DTD processing and with a limit on how deep elements nest.
- The zip directory is checked before anything is read: the total unpacked size, the number of
  parts, duplicate names, names that point outside the package, and entries that cannot be
  unpacked. Each part is then held to the size its zip header gave.

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

## Writing a package

```csharp
await using OpcPackageWriter writer = await OpcPackageWriter.CreateAsync(output);

await writer.WritePartAsync("word/document.xml", DocumentContentType, async stream => await WriteBodyAsync(stream));
writer.AddRelationship(string.Empty, OpcNamespaces.OfficeDocumentRelationship, "word/document.xml");

await writer.CompleteAsync();
```

The .rels parts and `[Content_Types].xml` are written by `CompleteAsync`. Pass
`OpcPackageKind.FlatOpc` to `CreateAsync` to write one XML file instead of a zip.

## Lower level

`OpcZip`, `OpcXml` and `OpcPath` are the helpers the two classes above use, for a package that
works on a `ZipArchive` itself, as Avig.Kalkyl does.

## License

MIT
