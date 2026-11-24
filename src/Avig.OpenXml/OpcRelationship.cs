namespace Avig.OpenXml;

/// <summary>
/// One relationship from a .rels file. <see cref="Target"/> is as written, relative to the
/// part that owns the relationship; resolve it with <see cref="OpcPath.ResolveTarget"/> unless
/// <see cref="IsExternal"/> is true, when it is a URL.
/// </summary>
/// <param name="Id">The id the part refers to it by, such as rId3.</param>
/// <param name="Type">The relationship type, a URI. Compare with <see cref="OpcPath.IsRelationship"/>.</param>
/// <param name="Target">Where it points.</param>
/// <param name="IsExternal">True for TargetMode="External": a URL outside the package.</param>
public sealed record OpcRelationship(string Id, string Type, string Target, bool IsExternal)
{
    /// <summary>The relationships in a .rels part, in file order. Ones without Id, Type or Target are left out.</summary>
    public static async Task<IReadOnlyList<OpcRelationship>> ReadAsync(Stream relsPart, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        using System.Xml.XmlReader reader = OpcXml.CreateReader(relsPart);

        return await OpcZip.ReadRelationshipsFromAsync(reader);
    }
}
