using System.Text;
using System.Xml;

namespace Avig.OpenXml.Drawing;

/// <summary>A cell of a <see cref="ChartWorkbook"/>: a text, a number or nothing.</summary>
/// <param name="Text">The text, or null.</param>
/// <param name="Number">The number, or null.</param>
public readonly record struct ChartCell(string? Text, double? Number)
{
    /// <summary>An empty cell.</summary>
    public static ChartCell Empty => default;

    /// <summary>True when the cell holds neither text nor a number.</summary>
    public bool IsEmpty => Text is null && Number is null;
}

/// <summary>
/// The workbook a chart in a presentation or a document keeps its data in, so that the
/// application's Edit Data opens it: one sheet with the categories down the first columns and a
/// column per series, the names in the first row. <see cref="From(ChartSpace)"/> lays out a
/// chart's stored values and gives back the chart with its formulas pointing at the cells
/// (<see cref="Chart"/>); <see cref="ToArray"/> writes the .xlsx, the smallest one Excel opens
/// without repair. The package that owns the chart stores it as a part of its own, with a
/// <see cref="PackageRelationship"/> from the chart part, whose id goes in
/// <see cref="ChartSpace.ExternalData"/>.
/// </summary>
public sealed class ChartWorkbook
{
    /// <summary>The name of the sheet.</summary>
    public const string SheetName = "Sheet1";

    /// <summary>The content type of the workbook part.</summary>
    public const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /// <summary>The relationship from a chart part to its workbook.</summary>
    public const string PackageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";

    /// <summary>The most rows a chart's workbook is laid out with, so that a hostile point count cannot make one without end.</summary>
    public const int MaxRows = 1_048_576;

    private const string SpreadsheetMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    private const string StylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    private const string WorksheetRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    private const string StylesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

    private readonly List<List<ChartCell>> rows;

    private ChartWorkbook(List<List<ChartCell>> rows)
    {
        this.rows = rows;
    }

    /// <summary>The cells, row by row from row 1; a row may be shorter than the others.</summary>
    public IReadOnlyList<IReadOnlyList<ChartCell>> Rows => rows;

    /// <summary>The chart <see cref="From(ChartSpace)"/> was given, with its formulas pointing at the cells; null for a workbook of a <see cref="ChartExSpace"/>.</summary>
    public ChartSpace? Chart { get; private init; }

    /// <summary>The chart <see cref="From(ChartExSpace)"/> was given, with its formulas pointing at the cells; null for a workbook of a <see cref="ChartSpace"/>.</summary>
    public ChartExSpace? ExtendedChart { get; private init; }

    /// <summary>
    /// Lays out a chart's stored values: the categories (the x values of a scatter or bubble
    /// chart) in column A when the series share them, otherwise in a column before each series'
    /// values; a column per series, and one more for a bubble series' sizes. Values a series has
    /// only as a formula, without stored values, leave their cells empty.
    /// </summary>
    /// <param name="chart">The chart.</param>
    /// <returns>The workbook, with <see cref="Chart"/> bound to it.</returns>
    public static ChartWorkbook From(ChartSpace chart)
    {
        ArgumentNullException.ThrowIfNull(chart);
        Sheet sheet = new();
        sheet.Set(0, 0, new ChartCell(" ", null));
        ChartData? shared = SharedCategories(chart.Series);
        int column = 0;

        if (shared is not null)
        {
            sheet.Column(column, shared);
            column++;
        }

        List<ChartSpaceSeries> bound = [];

        for (int index = 0; index < chart.Series.Count; index++)
        {
            ChartSpaceSeries series = chart.Series[index];
            ChartData? categories = series.Categories;

            if (shared is not null)
            {
                categories = shared with { Formula = Range(0, 1, Count(shared)) };
            }
            else if (series.Categories is not null)
            {
                sheet.Set(0, column, new ChartCell("X", null));
                sheet.Column(column, series.Categories);
                categories = series.Categories with { Formula = Range(column, 1, Count(series.Categories)) };

                column++;
            }

            string name = series.Name ?? "Series " + DrawingValues.Text(index + 1);
            sheet.Set(0, column, new ChartCell(name, null));
            sheet.Column(column, series.Values);
            ChartSpaceSeries made = series with
            {
                Name = name,
                NameFormula = Cell(column, 0),
                Values = series.Values with { Formula = Range(column, 1, Count(series.Values)) },
                Categories = categories,
            };

            column++;

            if (series.BubbleSizes is ChartData sizes)
            {
                sheet.Set(0, column, new ChartCell("Size", null));
                sheet.Column(column, sizes);
                made = made with { BubbleSizes = sizes with { Formula = Range(column, 1, Count(sizes)) } };

                column++;
            }

            bound.Add(made);
        }

        return new ChartWorkbook(sheet.Rows) { Chart = chart with { Series = bound } };
    }

    /// <summary>
    /// Lays out the data of a chart of the newer kinds: the category levels in the first columns,
    /// the outermost first, then a column per series.
    /// </summary>
    /// <param name="chart">The chart.</param>
    /// <returns>The workbook, with <see cref="ExtendedChart"/> bound to it.</returns>
    public static ChartWorkbook From(ChartExSpace chart)
    {
        ArgumentNullException.ThrowIfNull(chart);
        Sheet sheet = new();
        ChartExCategories? categories = chart.Categories;
        int column = 0;
        int count = chart.Series.Select(item => Count(item.Values)).DefaultIfEmpty(0).Max();

        if (categories is not null)
        {
            count = Math.Max(count, categories.Levels.Select(level => level.Count).DefaultIfEmpty(0).Max());

            for (int level = 0; level < categories.Levels.Count; level++)
            {
                sheet.Set(0, column + level, new ChartCell(" ", null));
                sheet.Column(column + level, ChartData.FromTexts(categories.Levels[level]));
            }

            int levels = Math.Max(1, categories.Levels.Count);
            categories = categories with { Formula = SheetName + "!" + Address(0, 1) + ":" + Address(levels - 1, Math.Max(1, count)) };

            column += levels;
        }

        List<ChartExSeries> bound = [];

        for (int index = 0; index < chart.Series.Count; index++)
        {
            ChartExSeries series = chart.Series[index];
            string name = series.Name ?? "Series " + DrawingValues.Text(index + 1);
            sheet.Set(0, column, new ChartCell(name, null));
            sheet.Column(column, series.Values);
            bound.Add(series with
            {
                Name = name,
                NameFormula = Cell(column, 0),
                Values = series.Values with { Formula = Range(column, 1, Math.Max(count, Count(series.Values))) },
            });

            column++;
        }

        return new ChartWorkbook(sheet.Rows) { ExtendedChart = chart with { Series = bound, Categories = categories } };
    }

    /// <summary>Writes the workbook as an .xlsx.</summary>
    /// <param name="output">Where it goes; left open.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>A task that completes when it is written.</returns>
    public async Task WriteAsync(Stream output, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(output);

        await using OpcPackageWriter package = await OpcPackageWriter.CreateAsync(output, OpcPackageKind.Zip, true, cancellationToken);
        await package.WritePartAsync("xl/workbook.xml", WorkbookContentType, Xml(WriteWorkbook), cancellationToken);
        await package.WritePartAsync("xl/worksheets/sheet1.xml", WorksheetContentType, Xml(WriteSheet), cancellationToken);
        await package.WritePartAsync("xl/styles.xml", StylesContentType, Xml(WriteStyles), cancellationToken);
        package.AddRelationship(string.Empty, OpcNamespaces.OfficeDocumentRelationship, "xl/workbook.xml");
        package.AddRelationship("xl/workbook.xml", WorksheetRelationship, "xl/worksheets/sheet1.xml");
        package.AddRelationship("xl/workbook.xml", StylesRelationship, "xl/styles.xml");
        await package.CompleteAsync(cancellationToken);
    }

    /// <summary>The .xlsx as bytes, written without I/O.</summary>
    /// <returns>The bytes.</returns>
    public byte[] ToArray()
    {
        using MemoryStream output = new();
        Task written = WriteAsync(output);

        // Into memory every step completes at once; a task that did not would mean a change in the package writer.
        if (written.IsCompleted == false)
        {
            throw new InvalidOperationException("The workbook was not written synchronously into memory.");
        }

        if (written.IsFaulted == true)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(written.Exception!.GetBaseException()).Throw();
        }

        return output.ToArray();
    }

    /// <summary>A cell's address, such as <c>Sheet1!$B$1</c>.</summary>
    /// <param name="column">The column, from 0.</param>
    /// <param name="row">The row, from 0.</param>
    /// <returns>The reference.</returns>
    public static string Cell(int column, int row)
    {
        return SheetName + "!" + Address(column, row);
    }

    private static string Range(int column, int firstRow, int count)
    {
        return SheetName + "!" + Address(column, firstRow) + ":" + Address(column, firstRow + Math.Max(1, count) - 1);
    }

    private static string Address(int column, int row)
    {
        return "$" + ColumnName(column) + "$" + DrawingValues.Text(row + 1);
    }

    private static string ColumnName(int column)
    {
        StringBuilder name = new();
        int remaining = column + 1;

        while (remaining > 0)
        {
            int digit = (remaining - 1) % 26;
            name.Insert(0, (char)('A' + digit));
            remaining = (remaining - 1) / 26;
        }

        return name.ToString();
    }

    private static int Count(ChartData data)
    {
        int count = Math.Max(data.Numbers?.Count ?? 0, data.Texts?.Count ?? 0);

        return Math.Min(count, MaxRows - 1);
    }

    /// <summary>The categories every series has the same of, or null.</summary>
    private static ChartData? SharedCategories(IReadOnlyList<ChartSpaceSeries> series)
    {
        ChartData? first = null;
        bool same = true;

        foreach (ChartSpaceSeries item in series)
        {
            if (item.Categories is null)
            {
                continue;
            }

            if (first is null)
            {
                first = item.Categories;

                continue;
            }

            same = same == true && SameValues(first, item.Categories) == true;
        }

        if (same == false)
        {
            first = null;
        }

        return first;
    }

    private static bool SameValues(ChartData first, ChartData second)
    {
        bool numbers = (first.Numbers ?? []).SequenceEqual(second.Numbers ?? []);
        bool texts = (first.Texts ?? []).SequenceEqual(second.Texts ?? [], StringComparer.Ordinal);

        return numbers == true && texts == true && (first.Formula is null || second.Formula is null || first.Formula == second.Formula);
    }

    private static byte[] Xml(Action<XmlWriter> write)
    {
        using MemoryStream stream = new();
        XmlWriterSettings settings = new() { Encoding = new UTF8Encoding(false) };

        using (XmlWriter xml = XmlWriter.Create(stream, settings))
        {
            xml.WriteStartDocument(true);
            write(xml);
            xml.WriteEndDocument();
        }

        return stream.ToArray();
    }

    private static void WriteWorkbook(XmlWriter xml)
    {
        xml.WriteStartElement(null, "workbook", SpreadsheetMain);
        xml.WriteAttributeString("xmlns", "r", null, OpcNamespaces.Relationships);
        xml.WriteStartElement(null, "sheets", SpreadsheetMain);
        xml.WriteStartElement(null, "sheet", SpreadsheetMain);
        xml.WriteAttributeString(null, "name", null, SheetName);
        xml.WriteAttributeString(null, "sheetId", null, "1");
        xml.WriteAttributeString("r", "id", OpcNamespaces.Relationships, "rId1");
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    private void WriteSheet(XmlWriter xml)
    {
        xml.WriteStartElement(null, "worksheet", SpreadsheetMain);
        xml.WriteStartElement(null, "sheetData", SpreadsheetMain);

        for (int row = 0; row < rows.Count; row++)
        {
            xml.WriteStartElement(null, "row", SpreadsheetMain);
            xml.WriteAttributeString(null, "r", null, DrawingValues.Text(row + 1));

            for (int column = 0; column < rows[row].Count; column++)
            {
                ChartCell cell = rows[row][column];

                if (cell.IsEmpty == true)
                {
                    continue;
                }

                xml.WriteStartElement(null, "c", SpreadsheetMain);
                xml.WriteAttributeString(null, "r", null, ColumnName(column) + DrawingValues.Text(row + 1));

                if (cell.Number is double number)
                {
                    xml.WriteElementString(null, "v", SpreadsheetMain, DrawingValues.Text(number));
                }
                else
                {
                    xml.WriteAttributeString(null, "t", null, "inlineStr");
                    xml.WriteStartElement(null, "is", SpreadsheetMain);
                    xml.WriteStartElement(null, "t", SpreadsheetMain);

                    if (cell.Text!.Length > 0 && (char.IsWhiteSpace(cell.Text[0]) == true || char.IsWhiteSpace(cell.Text[^1]) == true))
                    {
                        xml.WriteAttributeString("xml", "space", null, "preserve");
                    }

                    xml.WriteString(cell.Text);
                    xml.WriteEndElement();
                    xml.WriteEndElement();
                }

                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    private static void WriteStyles(XmlWriter xml)
    {
        xml.WriteStartElement(null, "styleSheet", SpreadsheetMain);
        xml.WriteStartElement(null, "fonts", SpreadsheetMain);
        xml.WriteAttributeString(null, "count", null, "1");
        xml.WriteStartElement(null, "font", SpreadsheetMain);
        Value(xml, "sz", "11");
        Value(xml, "name", "Calibri");
        Value(xml, "family", "2");
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteStartElement(null, "fills", SpreadsheetMain);
        xml.WriteAttributeString(null, "count", null, "2");

        foreach (string pattern in new[] { "none", "gray125" })
        {
            xml.WriteStartElement(null, "fill", SpreadsheetMain);
            xml.WriteStartElement(null, "patternFill", SpreadsheetMain);
            xml.WriteAttributeString(null, "patternType", null, pattern);
            xml.WriteEndElement();
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        xml.WriteStartElement(null, "borders", SpreadsheetMain);
        xml.WriteAttributeString(null, "count", null, "1");
        xml.WriteStartElement(null, "border", SpreadsheetMain);

        foreach (string side in new[] { "left", "right", "top", "bottom", "diagonal" })
        {
            xml.WriteElementString(null, side, SpreadsheetMain, null);
        }

        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteStartElement(null, "cellStyleXfs", SpreadsheetMain);
        xml.WriteAttributeString(null, "count", null, "1");
        Format(xml, false);
        xml.WriteEndElement();
        xml.WriteStartElement(null, "cellXfs", SpreadsheetMain);
        xml.WriteAttributeString(null, "count", null, "1");
        Format(xml, true);
        xml.WriteEndElement();
        xml.WriteStartElement(null, "cellStyles", SpreadsheetMain);
        xml.WriteAttributeString(null, "count", null, "1");
        xml.WriteStartElement(null, "cellStyle", SpreadsheetMain);
        xml.WriteAttributeString(null, "name", null, "Normal");
        xml.WriteAttributeString(null, "xfId", null, "0");
        xml.WriteAttributeString(null, "builtinId", null, "0");
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    private static void Value(XmlWriter xml, string name, string value)
    {
        xml.WriteStartElement(null, name, SpreadsheetMain);
        xml.WriteAttributeString(null, "val", null, value);
        xml.WriteEndElement();
    }

    private static void Format(XmlWriter xml, bool cell)
    {
        xml.WriteStartElement(null, "xf", SpreadsheetMain);
        xml.WriteAttributeString(null, "numFmtId", null, "0");
        xml.WriteAttributeString(null, "fontId", null, "0");
        xml.WriteAttributeString(null, "fillId", null, "0");
        xml.WriteAttributeString(null, "borderId", null, "0");

        if (cell == true)
        {
            xml.WriteAttributeString(null, "xfId", null, "0");
        }

        xml.WriteEndElement();
    }

    /// <summary>The cells as they are laid out, growing as they are set.</summary>
    private sealed class Sheet
    {
        public List<List<ChartCell>> Rows { get; } = [];

        public void Set(int row, int column, ChartCell cell)
        {
            if (row >= MaxRows)
            {
                return;
            }

            while (Rows.Count <= row)
            {
                Rows.Add([]);
            }

            List<ChartCell> cells = Rows[row];

            while (cells.Count <= column)
            {
                cells.Add(ChartCell.Empty);
            }

            cells[column] = cell;
        }

        /// <summary>The stored values of a series down a column from row 2.</summary>
        public void Column(int column, ChartData data)
        {
            int count = Count(data);

            for (int point = 0; point < count; point++)
            {
                double? number = null;
                string? text = null;

                if (data.Numbers is not null && point < data.Numbers.Count)
                {
                    number = data.Numbers[point];
                }

                if (number is null && data.Texts is not null && point < data.Texts.Count)
                {
                    text = data.Texts[point];
                }

                if (number is not null || text is not null)
                {
                    Set(point + 1, column, new ChartCell(text, number));
                }
            }
        }
    }
}
