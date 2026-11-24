namespace Avig.OpenXml;

/// <summary>Namespaces, relationship types and content types that every OPC package shares.</summary>
public static class OpcNamespaces
{
    /// <summary>The namespace of r:id and other relationship attributes in transitional OOXML.</summary>
    public const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>The namespace of the elements in a .rels part.</summary>
    public const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>The namespace of [Content_Types].xml.</summary>
    public const string ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>The namespace of Flat OPC (pkg:package), as Word writes it in "Word XML Document".</summary>
    public const string FlatOpc = "http://schemas.microsoft.com/office/2006/xmlPackage";

    /// <summary>The namespace of mc:Ignorable and mc:AlternateContent.</summary>
    public const string MarkupCompatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    /// <summary>The relationship from the package to its main part.</summary>
    public const string OfficeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    /// <summary>The relationship from the package to docProps/core.xml.</summary>
    public const string CorePropertiesRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

    /// <summary>The relationship from the package to docProps/app.xml.</summary>
    public const string ExtendedPropertiesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

    /// <summary>The relationship from a part to an image.</summary>
    public const string ImageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    /// <summary>The relationship from a part to a URL (with TargetMode="External").</summary>
    public const string HyperlinkRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    /// <summary>The content type of a .rels part.</summary>
    public const string RelationshipsContentType = "application/vnd.openxmlformats-package.relationships+xml";

    /// <summary>The content type of docProps/core.xml.</summary>
    public const string CorePropertiesContentType = "application/vnd.openxmlformats-package.core-properties+xml";

    /// <summary>The content type of docProps/app.xml.</summary>
    public const string ExtendedPropertiesContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";

    /// <summary>The content type of a VBA project, which only macro-enabled files have.</summary>
    public const string VbaProjectContentType = "application/vnd.ms-office.vbaProject";
}
