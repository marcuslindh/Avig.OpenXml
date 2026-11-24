using System.Diagnostics.CodeAnalysis;
using System.Xml;

namespace Avig.OpenXml;

/// <summary>
/// Reading the XML parts of a package safely. Elements are compared by local name, not
/// namespace, so that both transitional and Strict OOXML can be read with the same code.
/// </summary>
public static class OpcXml
{
    /// <summary>How deep elements may be nested in a part read through <see cref="CreateReader(Stream, int)"/>.</summary>
    public const int DefaultMaxDepth = DepthLimitedXmlReader.DefaultMaxDepth;

    /// <summary>
    /// Settings for every XML part read from a file: async, and no DTD and no resolver. A file
    /// often comes from outside, and a DTD is the way in for XXE and "billion laughs".
    /// </summary>
    public static XmlReaderSettings Settings { get; } = new()
    {
        Async = true,
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        CloseInput = false,
    };

    /// <summary>A reader for a part of a file, with <see cref="Settings"/> and a limit on how deep elements may be nested.</summary>
    public static XmlReader CreateReader(Stream stream, int maxDepth = DefaultMaxDepth)
    {
        return new DepthLimitedXmlReader(XmlReader.Create(stream, Settings), maxDepth);
    }

    /// <summary>As <see cref="CreateReader(Stream, int)"/>, for a fragment such as one element's XML kept from an earlier read.</summary>
    public static XmlReader CreateReader(TextReader text, XmlReaderSettings settings, XmlParserContext? context, int maxDepth = DefaultMaxDepth)
    {
        return new DepthLimitedXmlReader(XmlReader.Create(text, settings, context), maxDepth);
    }

    /// <summary>True when the reader is on the start of an element with this local name, in any namespace.</summary>
    public static bool IsElement(XmlReader reader, string localName)
    {
        return reader.NodeType == XmlNodeType.Element && string.Equals(reader.LocalName, localName, StringComparison.Ordinal) == true;
    }

    /// <summary>True when the reader is on the end of the element that started at <paramref name="depth"/>.</summary>
    public static bool IsEnd(XmlReader reader, int depth)
    {
        return reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth;
    }

    /// <summary>The r:id attribute, whose namespace differs between transitional and Strict OOXML.</summary>
    public static bool TryGetRelationshipId(XmlReader reader, [NotNullWhen(true)] out string? id)
    {
        return TryGetRelationshipAttribute(reader, "id", out id);
    }

    /// <summary>
    /// An attribute in the relationships namespace (r:id, r:embed, r:link …), found by local
    /// name, since the namespace differs between transitional and Strict OOXML.
    /// </summary>
    public static bool TryGetRelationshipAttribute(XmlReader reader, string localName, [NotNullWhen(true)] out string? value)
    {
        value = null;

        if (reader.MoveToFirstAttribute() == true)
        {
            do
            {
                if (string.Equals(reader.LocalName, localName, StringComparison.Ordinal) == true
                    && reader.NamespaceURI.Contains("relationships", StringComparison.Ordinal) == true)
                {
                    value = reader.Value;
                }
            }
            while (reader.MoveToNextAttribute() == true);

            reader.MoveToElement();
        }

        return value is not null;
    }

    /// <summary>An attribute found by local name in any namespace, such as w:val in transitional or Strict WordprocessingML.</summary>
    public static bool TryGetAttribute(XmlReader reader, string localName, [NotNullWhen(true)] out string? value)
    {
        value = null;

        if (reader.MoveToFirstAttribute() == false)
        {
            return false;
        }

        do
        {
            if (string.Equals(reader.LocalName, localName, StringComparison.Ordinal) == true
                && string.Equals(reader.Prefix, "xmlns", StringComparison.Ordinal) == false)
            {
                value = reader.Value;

                break;
            }
        }
        while (reader.MoveToNextAttribute() == true);

        reader.MoveToElement();

        return value is not null;
    }
}
