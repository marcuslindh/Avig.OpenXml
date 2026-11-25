namespace Avig.OpenXml;

/// <summary>
/// A relationship that points out of the package (TargetMode="External"), with the part it
/// belongs to. Office follows some of these when a file is opened, before anyone clicks
/// anything: a template attached from a server, a linked picture or OLE object, a frame, an
/// external workbook or data source. That is how remote template injection and Follina reach a
/// document, and how a file makes Windows send its user's credentials to a server.
/// <see cref="IsHyperlink"/> tells the links a reader has to click apart from those.
/// </summary>
/// <param name="SourcePart">The part the relationship belongs to; the empty name for the package itself.</param>
/// <param name="Relationship">The relationship; <see cref="OpcRelationship.Target"/> is the URL.</param>
public sealed record OpcExternalRelationship(string SourcePart, OpcRelationship Relationship)
{
    /// <summary>The URL it points to.</summary>
    public string Target => Relationship.Target;

    /// <summary>The relationship type's last segment, such as attachedTemplate, image, oleObject, frame or hyperlink.</summary>
    public string Kind => Relationship.Type[(Relationship.Type.LastIndexOf('/') + 1)..];

    /// <summary>True for a hyperlink, which is only followed when someone clicks it.</summary>
    public bool IsHyperlink => OpcPath.IsRelationship(Relationship.Type, "hyperlink");
}
