namespace Avig.OpenXml;

/// <summary>
/// How the messages of this package start, so that the packages built on it can tell what kind
/// of failure a message is (as KalkylError.ErrorKind does) without a type of its own in Result.
/// </summary>
public static class OpcErrors
{
    /// <summary>How every message about a damaged file starts.</summary>
    public const string DamagedPrefix = "The file is damaged";

    /// <summary>How every message about a file that goes past a limit names the limit: "the limit set in X.Y".</summary>
    public const string LimitMarker = "the limit set in ";

    /// <summary>The message for a file that unpacks to more than <paramref name="maxBytes"/>.</summary>
    public static string TooLarge(long maxBytes, string limitName)
    {
        return $"The file unpacks to more than {maxBytes:N0} bytes, {LimitMarker}{limitName}. It may be a zip bomb; if it is a file you trust, open it with larger limits.";
    }
}
