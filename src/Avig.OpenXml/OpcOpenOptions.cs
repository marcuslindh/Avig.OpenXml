namespace Avig.OpenXml;

/// <summary>
/// How <see cref="OpcPackage.OpenAsync(Stream, OpcOpenOptions?, bool, CancellationToken)"/> opens a file: the limits that
/// keep a file from outside from running the process out of memory, and how the messages name
/// the format and the limits, so that they read as the calling package's own.
/// </summary>
/// <example>
/// <code>
/// OpcOpenOptions options = new() { FormatName = "a .docx document", LimitsName = "OrdLimits" };
/// </code>
/// </example>
public sealed record OpcOpenOptions
{
    /// <summary>The options used when none are given.</summary>
    public static OpcOpenOptions Default { get; } = new();

    /// <summary>
    /// How many bytes the parts may unpack to, together. 1 GiB by default. Each part of a zip
    /// file is also held to the size its zip header gives, so a header cannot lie its way past
    /// this. For Flat OPC it limits the size of the file.
    /// </summary>
    public long MaxUncompressedBytes { get; init; } = 1L << 30;

    /// <summary>How many parts the package may have. 100 000 by default; Office files have tens.</summary>
    public int MaxParts { get; init; } = 100_000;

    /// <summary>How deep elements may be nested in the XML parts read through <see cref="OpcPackage.CreateXmlReader"/>.</summary>
    public int MaxXmlDepth { get; init; } = OpcXml.DefaultMaxDepth;

    /// <summary>
    /// The type the limits are set in, as the messages name it: "the limit set in
    /// OrdLimits.MaxUncompressedBytes". "OpcOpenOptions" by default.
    /// </summary>
    public string LimitsName { get; init; } = "OpcOpenOptions";

    /// <summary>The format in the message for a file that is not one: "The file is not a .docx document.".</summary>
    public string FormatName { get; init; } = "an Open XML package";

    /// <summary>
    /// Added to the message for a file in the old binary compound format, which is both what
    /// .doc/.xls files and password-protected Open XML files are.
    /// </summary>
    public string CompoundFileHint { get; init; } = "It is either a file in the old binary format or a password-protected file; neither can be read.";

    /// <summary>Whether a Flat OPC file (the whole package as one XML file) is opened too. True by default.</summary>
    public bool AllowFlatOpc { get; init; } = true;

    /// <summary>
    /// Whether a zip file must say one thing only (see <see cref="OpcZip.CheckStructureAsync"/>):
    /// each local header as the central directory has it, and no bytes before, between or after
    /// the entries. True by default; office programs never write anything else, and a file that
    /// does is one that can look like two different documents to two different readers.
    /// </summary>
    public bool StrictZip { get; init; } = true;

    /// <summary>How every message about a file that is not of the format starts.</summary>
    public string NotFormatPrefix => $"The file is not {FormatName}";
}
