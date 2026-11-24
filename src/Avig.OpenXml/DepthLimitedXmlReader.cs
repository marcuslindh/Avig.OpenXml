using System.Xml;

namespace Avig.OpenXml;

/// <summary>
/// An XmlReader that refuses elements nested deeper than a maximum depth,
/// <see cref="DefaultMaxDepth"/> unless told otherwise. Office XML never goes past a few dozen
/// levels, but 200 000 nested &lt;r&gt; in one spreadsheet cell (1.4 MB of file) took 28 s and
/// 107 GB of allocations in Avig.Kalkyl's document model, because copying an element's XML costs
/// time in proportion to its depth at every level. The error is an XmlException, which every
/// caller already turns into a damaged file.
/// </summary>
internal sealed class DepthLimitedXmlReader : XmlReader, IXmlLineInfo
{
    /// <summary>Ten times what Excel writes, and far below where depth starts to cost.</summary>
    public const int DefaultMaxDepth = 100;

    private readonly int maxDepth;

    private readonly XmlReader inner;

    /// <summary>
    /// True when the inner reader is synchronous. Its async methods then answer from the
    /// synchronous ones, so code written against the async API can read an in-memory fragment
    /// with a synchronous reader, whose buffer is 4 kB rather than the 64 kB an async reader takes.
    /// </summary>
    private readonly bool synchronous;

    private readonly XmlReaderSettings? settings;

    public DepthLimitedXmlReader(XmlReader inner, int maxDepth = DefaultMaxDepth)
    {
        this.inner = inner;
        this.maxDepth = maxDepth;
        settings = inner.Settings;
        synchronous = inner.Settings?.Async != true;

        if (synchronous == true && settings is not null)
        {
            settings = settings.Clone();
            settings.Async = true;
        }
    }

    public override int AttributeCount => inner.AttributeCount;

    public override string BaseURI => inner.BaseURI;

    public override int Depth => inner.Depth;

    public override bool EOF => inner.EOF;

    public override bool HasValue => inner.HasValue;

    public override bool IsDefault => inner.IsDefault;

    public override bool IsEmptyElement => inner.IsEmptyElement;

    public override string LocalName => inner.LocalName;

    public override string Name => inner.Name;

    public override string NamespaceURI => inner.NamespaceURI;

    public override XmlNameTable NameTable => inner.NameTable;

    public override XmlNodeType NodeType => inner.NodeType;

    public override string Prefix => inner.Prefix;

    public override ReadState ReadState => inner.ReadState;

    public override XmlReaderSettings? Settings => settings;

    // The value of a text node is decoded when it is asked for, so it can meet the same defect as Read.
    public override string Value
    {
        get
        {
            try
            {
                return inner.Value;
            }
            catch (Exception exception) when (IsReaderDefect(exception) == true)
            {
                throw Broken(exception);
            }
        }
    }

    public override string XmlLang => inner.XmlLang;

    public override XmlSpace XmlSpace => inner.XmlSpace;

    /// <summary>The line in the part, for error messages; 0 when the inner reader has none.</summary>
    public int LineNumber
    {
        get
        {
            if (inner is IXmlLineInfo info)
            {
                return info.LineNumber;
            }

            return 0;
        }
    }

    public int LinePosition
    {
        get
        {
            if (inner is IXmlLineInfo info)
            {
                return info.LinePosition;
            }

            return 0;
        }
    }

    public bool HasLineInfo()
    {
        return inner is IXmlLineInfo info && info.HasLineInfo();
    }

    public override string GetAttribute(int index)
    {
        return inner.GetAttribute(index);
    }

    public override string? GetAttribute(string name)
    {
        return inner.GetAttribute(name);
    }

    public override string? GetAttribute(string name, string? namespaceURI)
    {
        return inner.GetAttribute(name, namespaceURI);
    }

    public override string? LookupNamespace(string prefix)
    {
        return inner.LookupNamespace(prefix);
    }

    public override bool MoveToAttribute(string name)
    {
        return inner.MoveToAttribute(name);
    }

    public override bool MoveToAttribute(string name, string? ns)
    {
        return inner.MoveToAttribute(name, ns);
    }

    public override void MoveToAttribute(int index)
    {
        inner.MoveToAttribute(index);
    }

    public override bool MoveToElement()
    {
        return inner.MoveToElement();
    }

    public override bool MoveToFirstAttribute()
    {
        return inner.MoveToFirstAttribute();
    }

    public override bool MoveToNextAttribute()
    {
        return inner.MoveToNextAttribute();
    }

    public override bool Read()
    {
        bool more;

        try
        {
            more = inner.Read();
        }
        catch (Exception exception) when (IsReaderDefect(exception) == true)
        {
            throw Broken(exception);
        }

        Check();

        return more;
    }

    public override async Task<bool> ReadAsync()
    {
        if (synchronous == true)
        {
            return Read();
        }

        bool more;

        try
        {
            more = await inner.ReadAsync();
        }
        catch (Exception exception) when (IsReaderDefect(exception) == true)
        {
            throw Broken(exception);
        }

        Check();

        return more;
    }

    public override async Task<string> GetValueAsync()
    {
        if (synchronous == true)
        {
            return Value;
        }

        try
        {
            return await inner.GetValueAsync();
        }
        catch (Exception exception) when (IsReaderDefect(exception) == true)
        {
            throw Broken(exception);
        }
    }

    // Straight to the inner reader for speed: an element read as text has no child elements, so
    // there is no depth to check. Skip stays on this reader: a deep subtree skipped when the file
    // is opened would otherwise be found only when the file is saved and the subtree copied.
    public override string ReadElementContentAsString()
    {
        try
        {
            return inner.ReadElementContentAsString();
        }
        catch (Exception exception) when (IsReaderDefect(exception) == true)
        {
            throw Broken(exception);
        }
    }

    public override async Task<string> ReadElementContentAsStringAsync()
    {
        if (synchronous == true)
        {
            return ReadElementContentAsString();
        }

        try
        {
            return await inner.ReadElementContentAsStringAsync();
        }
        catch (Exception exception) when (IsReaderDefect(exception) == true)
        {
            throw Broken(exception);
        }
    }

    public override bool ReadAttributeValue()
    {
        return inner.ReadAttributeValue();
    }

    public override void ResolveEntity()
    {
        inner.ResolveEntity();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing == true)
        {
            inner.Dispose();
        }

        base.Dispose(disposing);
    }

    /// <summary>
    /// An exception the XML reader should not throw for any input. XmlTextReaderImpl in .NET 10
    /// throws ArgumentOutOfRangeException from Buffer.BlockCopy for a malformed XML declaration
    /// read from a deflate stream (found by libFuzzer, fixed in .NET 11). Any input that does it
    /// is broken XML, and is reported as such.
    /// </summary>
    private static bool IsReaderDefect(Exception exception)
    {
        return exception is ArgumentException or IndexOutOfRangeException;
    }

    private XmlException Broken(Exception exception)
    {
        return new XmlException($"The XML could not be read: {exception.Message}", exception, LineNumber, LinePosition);
    }

    private void Check()
    {
        if (inner.Depth > maxDepth)
        {
            throw new XmlException($"Elements are nested more than {maxDepth} levels deep, which no office program writes.", null, LineNumber, LinePosition);
        }
    }
}
