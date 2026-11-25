namespace Avig.OpenXml.Drawing;

/// <summary>How text in a shape, a chart title, a legend, an axis or a data label looks. Every property left unset is the application's own.</summary>
public sealed record TextFont
{
    private readonly double? size;
    private readonly string? color;
    private readonly string? name;

    /// <summary>The typeface, such as <c>Segoe UI</c>, or null for the theme's.</summary>
    public string? Name
    {
        get => name;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 31))
            {
                throw new ArgumentException("A font name holds 1 to 31 characters.", nameof(Name));
            }

            name = value;
        }
    }

    /// <summary>The size in points, 1 to 400, or null for the default (11 in an Excel shape, 10 in a chart, 18 on a slide).</summary>
    public double? Size
    {
        get => size;
        init
        {
            if (value is double points)
            {
                DrawingValues.Range(points, 1, 400, nameof(Size), "A font size");
            }

            size = value;
        }
    }

    /// <summary>Bold.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic.</summary>
    public bool Italic { get; init; }

    /// <summary>The colour as <c>#RRGGBB</c>, or null for the default.</summary>
    public string? Color
    {
        get => color;
        init
        {
            string? normalized = null;

            if (value is not null)
            {
                normalized = DrawingValues.Color(value, nameof(Color));
            }

            color = normalized;
        }
    }
}

/// <summary>A piece of text with one look. A <c>\n</c> in the text starts a new paragraph.</summary>
public sealed record TextRun
{
    /// <summary>Creates a run.</summary>
    /// <param name="text">The text.</param>
    /// <param name="font">How it looks, or null for the shape's default.</param>
    public TextRun(string text, TextFont? font = null)
    {
        ArgumentNullException.ThrowIfNull(text);

        if (text.Length > DrawingValues.MaxTextLength)
        {
            throw new ArgumentException($"A text run holds at most {DrawingValues.MaxTextLength} characters.", nameof(text));
        }

        Text = DrawingValues.XmlText(text, nameof(text));
        Font = font;
    }

    /// <summary>The text.</summary>
    public string Text { get; }

    /// <summary>How it looks, or null for the shape's default.</summary>
    public TextFont? Font { get; }
}

/// <summary>Where the lines of a <see cref="ShapeText"/> sit across the shape (a:pPr algn).</summary>
public enum TextAlignment
{
    /// <summary>Along the left edge.</summary>
    Left,

    /// <summary>In the middle.</summary>
    Center,

    /// <summary>Along the right edge.</summary>
    Right,

    /// <summary>Stretched to both edges, except the last line.</summary>
    Justify,

    /// <summary>Spread over the whole line, the last line too.</summary>
    Distributed,
}

/// <summary>Where the text of a <see cref="ShapeText"/> sits from top to bottom (a:bodyPr anchor).</summary>
public enum TextAnchor
{
    /// <summary>At the top.</summary>
    Top,

    /// <summary>In the middle.</summary>
    Center,

    /// <summary>At the bottom.</summary>
    Bottom,
}

/// <summary>
/// The text inside a <see cref="Shape"/>: one or more runs, aligned and wrapped in the shape. A
/// <c>\n</c> in a run starts a new paragraph (a:p); each run keeps its look on both sides of it.
/// </summary>
public sealed record ShapeText
{
    private readonly double? marginLeft;
    private readonly double? marginTop;
    private readonly double? marginRight;
    private readonly double? marginBottom;

    /// <summary>Text with one look.</summary>
    /// <param name="text">The text; <c>\n</c> starts a new paragraph.</param>
    /// <param name="font">How it looks, or null for the shape's default.</param>
    public ShapeText(string text, TextFont? font = null)
        : this([new TextRun(text, font)])
    {
    }

    /// <summary>Text made of runs, such as a word in another colour.</summary>
    /// <param name="runs">The runs, at least one.</param>
    public ShapeText(IReadOnlyList<TextRun> runs)
    {
        ArgumentNullException.ThrowIfNull(runs);

        if (runs.Count == 0)
        {
            throw new ArgumentException("A shape's text has at least one run.", nameof(runs));
        }

        foreach (TextRun run in runs)
        {
            ArgumentNullException.ThrowIfNull(run, nameof(runs));
        }

        Runs = [.. runs];
    }

    /// <summary>The runs.</summary>
    public IReadOnlyList<TextRun> Runs { get; }

    /// <summary>The text of all the runs, with <c>\n</c> between paragraphs.</summary>
    public string PlainText => string.Concat(Runs.Select(run => run.Text));

    /// <summary>Where the lines sit across the shape.</summary>
    public TextAlignment Horizontal { get; init; } = TextAlignment.Center;

    /// <summary>Where the text sits from top to bottom.</summary>
    public TextAnchor Vertical { get; init; } = TextAnchor.Center;

    /// <summary>Break lines at the shape's edge. Off, a line runs on past it.</summary>
    public bool Wrap { get; init; } = true;

    /// <summary>The space between the shape's left edge and the text, in points, or null for Office's 7.2.</summary>
    public double? MarginLeft
    {
        get => marginLeft;
        init => marginLeft = Margin(value, nameof(MarginLeft));
    }

    /// <summary>The space above the text, in points, or null for Office's 3.6.</summary>
    public double? MarginTop
    {
        get => marginTop;
        init => marginTop = Margin(value, nameof(MarginTop));
    }

    /// <summary>The space right of the text, in points, or null for Office's 7.2.</summary>
    public double? MarginRight
    {
        get => marginRight;
        init => marginRight = Margin(value, nameof(MarginRight));
    }

    /// <summary>The space below the text, in points, or null for Office's 3.6.</summary>
    public double? MarginBottom
    {
        get => marginBottom;
        init => marginBottom = Margin(value, nameof(MarginBottom));
    }

    private static double? Margin(double? value, string parameter)
    {
        if (value is double points)
        {
            DrawingValues.Range(points, 0, 1000, parameter, "A margin");
        }

        return value;
    }
}
