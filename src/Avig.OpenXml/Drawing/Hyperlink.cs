namespace Avig.OpenXml.Drawing;

/// <summary>
/// Where a click on a shape, a group or a picture goes: a web address (<see cref="ToUrl(string)"/>)
/// or a place inside the document (<see cref="ToLocation(string)"/>), such as a cell in a workbook
/// or a slide in a presentation. What a location means is up to the package that writes it:
/// Avig.Kalkyl's <c>Hyperlink.ToCell</c> makes one that names a sheet and a cell.
/// </summary>
public sealed record Hyperlink
{
    private readonly string? tooltip;
    private readonly string? action;

    private Hyperlink(string? url, string? location)
    {
        Url = url;
        Location = location;
    }

    /// <summary>The web address, or null for a link inside the document.</summary>
    public string? Url { get; }

    /// <summary>
    /// The place inside the document, as the package that owns the drawing stores it (in a
    /// workbook <c>'Sales 2026'!A1</c>), or null for a web address.
    /// </summary>
    public string? Location { get; }

    /// <summary>True when the link goes to a web address, and so needs an external relationship.</summary>
    public bool IsExternal => Url is not null;

    /// <summary>The text shown when the pointer rests on the link, or null.</summary>
    public string? Tooltip
    {
        get => tooltip;
        init
        {
            if (value is not null && value.Length > 255)
            {
                throw new ArgumentException("A tooltip holds at most 255 characters.", nameof(Tooltip));
            }

            if (value is not null)
            {
                DrawingValues.XmlText(value, nameof(Tooltip));
            }

            tooltip = value;
        }
    }

    /// <summary>
    /// The action of the click (a:hlinkClick action), such as PowerPoint's
    /// <c>ppaction://hlinksldjump</c> or <c>ppaction://hlinkshowjump?jump=nextslide</c>, or null
    /// for an ordinary link.
    /// </summary>
    public string? Action
    {
        get => action;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 2079))
            {
                throw new ArgumentException("An action holds 1 to 2079 characters.", nameof(Action));
            }

            if (value is not null)
            {
                DrawingValues.XmlText(value, nameof(Action));
            }

            action = value;
        }
    }

    /// <summary>A link to a web page or a mail address.</summary>
    /// <param name="url">An absolute address: <c>https://</c>, <c>http://</c> or <c>mailto:</c>.</param>
    /// <returns>The link.</returns>
    public static Hyperlink ToUrl(string url)
    {
        ArgumentNullException.ThrowIfNull(url);

        if (url.Length > 2079 || Uri.TryCreate(url, UriKind.Absolute, out Uri? parsed) == false
            || (parsed.Scheme != Uri.UriSchemeHttps && parsed.Scheme != Uri.UriSchemeHttp && parsed.Scheme != Uri.UriSchemeMailto))
        {
            throw new ArgumentException($"'{url}' is not an absolute https, http or mailto address of at most 2079 characters.", nameof(url));
        }

        return new Hyperlink(url, null);
    }

    /// <summary>
    /// A link to a place inside the document. The package that writes the drawing decides what
    /// the text means and how it is stored; Avig.Kalkyl writes it as the fragment of the
    /// relationship (<c>#'Sheet'!A1</c>).
    /// </summary>
    /// <param name="location">The place, 1 to 2079 characters.</param>
    /// <returns>The link.</returns>
    public static Hyperlink ToLocation(string location)
    {
        ArgumentNullException.ThrowIfNull(location);

        if (location.Length == 0 || location.Length > 2079)
        {
            throw new ArgumentException("A location holds 1 to 2079 characters.", nameof(location));
        }

        DrawingValues.XmlText(location, nameof(location));

        return new Hyperlink(null, location);
    }

    /// <summary>
    /// A link read from a file: the relationship's target as it is, without the checks of
    /// <see cref="ToUrl(string)"/>, since a file may hold any scheme. Nothing is ever followed.
    /// </summary>
    internal static Hyperlink Read(string? url, string? location)
    {
        return new Hyperlink(url, location);
    }
}
