using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

if (args.Length < 3)
{
    Console.Error.WriteLine("Usage: PreviewEngine <input.pptx> <output.pptx> <payload.json> [summary.json]");
    Environment.Exit(2);
}

var inputPath   = args[0];
var outputPath  = args[1];
var payloadPath = args[2];
var summaryPath = args.Length > 3 ? args[3] : string.Empty;

var jsonOpts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

var payloadJson = await File.ReadAllTextAsync(payloadPath);
var payload = JsonSerializer.Deserialize<PreviewPayload>(payloadJson, jsonOpts) ?? new PreviewPayload();

File.Copy(inputPath, outputPath, overwrite: true);

var replacements = payload.Replacements ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
var tableVars = (payload.TableVariables ?? new List<TableVariable>())
    .Where(v => !string.IsNullOrWhiteSpace(v.Name))
    .ToDictionary(v => v.Name!, v => v, StringComparer.OrdinalIgnoreCase);

var stats = new PreviewStats();

using (var doc = PresentationDocument.Open(outputPath, true))
{
    var presPart = doc.PresentationPart;
    if (presPart?.Presentation?.SlideIdList == null)
    {
        throw new InvalidOperationException("Presentation has no slides.");
    }

    var slideIds = presPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
    stats.SlideCount = slideIds.Count;

    for (var i = 0; i < slideIds.Count; i++)
    {
        var relId = slideIds[i].RelationshipId?.Value;
        if (string.IsNullOrWhiteSpace(relId)) continue;

        var slidePart = (SlidePart)presPart.GetPartById(relId);
        var slideKey  = (i + 1).ToString();

        // Prefer payload (from JS-side Office.js tag scan) over PPTX XML tag scanning
        var dynTables = payload.DynamicTables?.GetValueOrDefault(slideKey) ?? new List<DynamicTableConfig>();
        var dynImages = payload.DynamicImages?.GetValueOrDefault(slideKey) ?? new List<DynamicImageConfig>();
        var dynCharts = payload.DynamicCharts?.GetValueOrDefault(slideKey) ?? new List<DynamicChartConfig>();

        if (dynTables.Count == 0 && dynImages.Count == 0 && dynCharts.Count == 0)
        {
            var scanned = ScanSlideForTags(slidePart, jsonOpts);
            dynTables = scanned.Tables;
            dynImages = scanned.Images;
            dynCharts = scanned.Charts;
        }

        foreach (var cfg in dynTables)
        {
            if (ApplyDynamicTable(slidePart, cfg, tableVars))
                stats.ExpandedTables++;
        }

        foreach (var cfg in dynImages)
        {
            if (await ApplyDynamicImageAsync(slidePart, cfg))
                stats.UpdatedImages++;
        }

        foreach (var cfg in dynCharts)
        {
            if (ApplyDynamicChart(slidePart, cfg, tableVars, replacements))
                stats.UpdatedCharts++;
        }

        stats.ReplacedTokens += ReplaceTokensInSlide(slidePart, replacements);
        slidePart.Slide.Save();
    }
}

if (!string.IsNullOrWhiteSpace(summaryPath))
{
    var summaryJson = JsonSerializer.Serialize(stats, new JsonSerializerOptions { WriteIndented = true });
    await File.WriteAllTextAsync(summaryPath, summaryJson);
}

Console.WriteLine(JsonSerializer.Serialize(stats));

// ── Tag scanning ───────────────────────────────────────────────────────────────

static string? GetNvPrTagValue(DocumentFormat.OpenXml.OpenXmlElement? nvPr, string tagKey)
{
    if (nvPr == null) return null;
    XNamespace ns = "http://schemas.openxmlformats.org/presentationml/2006/main";
    try
    {
        var el  = XElement.Parse(nvPr.OuterXml);
        var tag = el.Descendants(ns + "tag")
            .FirstOrDefault(t => (string?)t.Attribute("name") == tagKey);
        return (string?)tag?.Attribute("val");
    }
    catch { return null; }
}

static (List<DynamicTableConfig> Tables, List<DynamicImageConfig> Images, List<DynamicChartConfig> Charts)
    ScanSlideForTags(SlidePart slidePart, JsonSerializerOptions jsonOpts)
{
    var tables = new List<DynamicTableConfig>();
    var images = new List<DynamicImageConfig>();
    var charts = new List<DynamicChartConfig>();

    // Graphic frames host tables and charts
    foreach (var frame in slidePart.Slide.Descendants<P.GraphicFrame>())
    {
        var shapeName = frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value;
        if (string.IsNullOrWhiteSpace(shapeName)) continue;

        var nvPr = frame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties;

        var tableTagVal = GetNvPrTagValue(nvPr, "LIVEDOC_DYN_TABLE");
        if (tableTagVal != null)
        {
            try
            {
                var cfg = JsonSerializer.Deserialize<DynamicTableConfig>(tableTagVal, jsonOpts);
                if (cfg != null) { cfg.ShapeName = shapeName; tables.Add(cfg); }
            }
            catch { }
        }

        var chartTagVal = GetNvPrTagValue(nvPr, "LIVEDOC_DYN_CHART");
        if (chartTagVal != null)
        {
            try
            {
                var cfg = JsonSerializer.Deserialize<DynamicChartConfig>(chartTagVal, jsonOpts);
                if (cfg != null) { cfg.ShapeName = shapeName; charts.Add(cfg); }
            }
            catch { }
        }
    }

    // Shapes host dynamic images (skip indicator shapes)
    foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
    {
        var shapeName = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
        if (string.IsNullOrWhiteSpace(shapeName)) continue;
        if (shapeName.StartsWith("__LIVEDOC_IND_", StringComparison.Ordinal)) continue;

        var nvPr = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;
        var imageTagVal = GetNvPrTagValue(nvPr, "LIVEDOC_DYN_IMAGE");
        if (imageTagVal != null)
        {
            try
            {
                var cfg = JsonSerializer.Deserialize<DynamicImageConfig>(imageTagVal, jsonOpts);
                if (cfg != null) { cfg.ShapeName = shapeName; images.Add(cfg); }
            }
            catch { }
        }
    }

    return (tables, images, charts);
}

// ── Apply handlers ─────────────────────────────────────────────────────────────

static bool ApplyDynamicTable(SlidePart slidePart, DynamicTableConfig cfg, Dictionary<string, TableVariable> tableVars)
{
    if (string.IsNullOrWhiteSpace(cfg.ShapeName) || string.IsNullOrWhiteSpace(cfg.VariableName)) return false;
    if (!tableVars.TryGetValue(cfg.VariableName, out var tableVar)) return false;
    if (tableVar.Columns == null || tableVar.Rows == null || tableVar.Columns.Count == 0) return false;
    if (tableVar.Rows.Count == 0 && !cfg.HideWhenEmpty) return false;

    var allTableFrames = slidePart.Slide.Descendants<P.GraphicFrame>()
        .Where(f => f.Graphic?.GraphicData?.GetFirstChild<A.Table>() != null)
        .ToList();

    var frame = allTableFrames.FirstOrDefault(f =>
        string.Equals(
            (f.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty).Trim(),
            cfg.ShapeName.Trim(),
            StringComparison.OrdinalIgnoreCase));

    if (frame == null) return false;

    var table = frame.Graphic?.GraphicData?.GetFirstChild<A.Table>();
    if (table == null) return false;

    var rows = table.Elements<A.TableRow>().ToList();
    if (rows.Count == 0) return false;

    var fromIdx = Math.Max(0, cfg.FromRow - 1);
    var toIdx   = Math.Max(0, cfg.ToRow   - 1);
    if (fromIdx >= rows.Count || toIdx >= rows.Count || fromIdx > toIdx) return false;

    var templateRows = rows.Skip(fromIdx).Take(toIdx - fromIdx + 1).ToList();
    if (templateRows.Count == 0) return false;

    // AutoTokenize: fill the header row (last row before fromRow) with column names
    if (cfg.AutoTokenize && fromIdx > 0)
    {
        var headerRow = rows[fromIdx - 1];
        FillRowByIndex(headerRow, tableVar.Columns.Select(c => c ?? string.Empty).ToList());
    }

    // HideWhenEmpty: remove template rows and stop — slide content is preserved but table shrinks
    if (tableVar.Rows.Count == 0 && cfg.HideWhenEmpty)
    {
        foreach (var tr in templateRows) tr.Remove();
        return true;
    }

    // Sort data rows according to sort rules (multi-key, stable)
    IEnumerable<List<string>> orderedRows = tableVar.Rows;
    foreach (var sortRule in Enumerable.Reverse(cfg.SortRules))
    {
        var colIdx = tableVar.Columns.IndexOf(sortRule.Column);
        if (colIdx < 0) continue;
        orderedRows = string.Equals(sortRule.Direction, "desc", StringComparison.OrdinalIgnoreCase)
            ? orderedRows.OrderByDescending(r => colIdx < r.Count ? r[colIdx] : string.Empty, StringComparer.OrdinalIgnoreCase)
            : orderedRows.OrderBy(r => colIdx < r.Count ? r[colIdx] : string.Empty, StringComparer.OrdinalIgnoreCase);
    }
    var sortedRows = orderedRows.ToList();

    var insertBefore = templateRows[0];
    var expandedRows = new List<A.TableRow>();

    // Step 1: expand with conditional filtering and token replacement
    foreach (var dataRow in sortedRows)
    {
        var rowMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var c = 0; c < tableVar.Columns.Count; c++)
        {
            var colName = tableVar.Columns[c] ?? string.Empty;
            var val     = c < dataRow.Count ? dataRow[c] : string.Empty;
            rowMap[colName] = val ?? string.Empty;
        }

        for (var ti = 0; ti < templateRows.Count; ti++)
        {
            var absRowIndex = fromIdx + ti + 1;
            var rule = cfg.ConditionalRows.FirstOrDefault(r => r.RowIndex == absRowIndex);
            if (rule != null && !EvaluateCondition(rule, rowMap)) continue;

            var clone = (A.TableRow)templateRows[ti].CloneNode(true);
            if (cfg.AutoTokenize)
            {
                // Fill cells positionally — no tokens needed in slide cells
                var values = tableVar.Columns.Select((col, i) =>
                    rowMap.TryGetValue(col ?? string.Empty, out var v) ? v : string.Empty).ToList();
                FillRowByIndex(clone, values);
            }
            else
            {
                ReplaceTokensInRow(clone, rowMap);
            }
            table.InsertBefore(clone, insertBefore);
            expandedRows.Add(clone);
        }
    }

    foreach (var tr in templateRows)
        tr.Remove();

    // Step 2: apply vertical merge rules across all expanded rows
    foreach (var mergeRule in cfg.MergeRules)
        ApplyVerticalMerge(expandedRows, mergeRule.ColumnIndex - 1, mergeRule.Strategy);

    return true;
}

static void FillRowByIndex(A.TableRow row, List<string> values)
{
    var cells = row.Elements<A.TableCell>().ToList();
    for (var ci = 0; ci < Math.Min(cells.Count, values.Count); ci++)
    {
        var val   = values[ci];
        var texts = cells[ci].Descendants<A.Text>().ToList();
        if (texts.Count > 0)
        {
            texts[0].Text = val;
            for (var i = 1; i < texts.Count; i++) texts[i].Text = string.Empty;
        }
        else
        {
            var para = cells[ci].Descendants<A.Paragraph>().FirstOrDefault();
            if (para != null)
                para.AppendChild(new A.Run(new A.Text(val)));
        }
    }
}

static bool EvaluateCondition(ConditionalRowRule rule, Dictionary<string, string> rowMap)
{
    var val = rowMap.TryGetValue(rule.Column, out var v) ? v : string.Empty;
    return rule.Operator switch
    {
        "empty"     => string.IsNullOrEmpty(val),
        "equals"    => string.Equals(val, rule.Value, StringComparison.OrdinalIgnoreCase),
        "notEquals" => !string.Equals(val, rule.Value, StringComparison.OrdinalIgnoreCase),
        _           => !string.IsNullOrEmpty(val),   // "notEmpty" (default)
    };
}

static void ReplaceTokensInRow(A.TableRow row, Dictionary<string, string> rowMap)
{
    foreach (var para in row.Descendants<A.Paragraph>())
    {
        var texts = para.Descendants<A.Text>().ToList();
        if (texts.Count == 0) continue;

        var combined = string.Concat(texts.Select(t => t.Text ?? string.Empty));
        if (!combined.Contains("{{", StringComparison.Ordinal)) continue;

        var replaced = ReplaceTokenString(combined, name => rowMap.TryGetValue(name, out var v) ? v : null);
        texts[0].Text = replaced;
        for (var i = 1; i < texts.Count; i++)
            texts[i].Text = string.Empty;
    }
}

static void ApplyVerticalMerge(List<A.TableRow> rows, int colIdx, string strategy)
{
    if (colIdx < 0 || rows.Count < 2) return;

    var cells = rows.Select(r => r.Elements<A.TableCell>().ElementAtOrDefault(colIdx)).ToList();
    if (cells.Any(c => c == null)) return;

    static string GetText(A.TableCell? cell) =>
        string.Concat(cell?.Descendants<A.Text>().Select(t => t.Text ?? string.Empty) ?? Enumerable.Empty<string>());

    var i = 0;
    while (i < cells.Count)
    {
        var anchor = cells[i]!;
        var anchorText = GetText(anchor);

        bool ShouldMerge(A.TableCell? next) => strategy == "empty"
            ? string.IsNullOrEmpty(GetText(next))
            : GetText(next) == anchorText;

        var j = i + 1;
        while (j < cells.Count && ShouldMerge(cells[j])) j++;

        if (j - i > 1)
        {
            anchor.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("rowSpan", string.Empty, (j - i).ToString()));

            for (var k = i + 1; k < j; k++)
            {
                var cont = cells[k]!;
                cont.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("vMerge", string.Empty, "1"));
                foreach (var run in cont.Descendants<A.Run>())
                {
                    var textEl = run.GetFirstChild<A.Text>();
                    if (textEl != null) textEl.Text = string.Empty;
                }
            }
        }

        i = j;
    }
}

static bool ApplyDynamicChart(
    SlidePart slidePart,
    DynamicChartConfig cfg,
    Dictionary<string, TableVariable> tableVars,
    Dictionary<string, string> replacements)
{
    if (string.IsNullOrWhiteSpace(cfg.ShapeName) || string.IsNullOrWhiteSpace(cfg.TableVar)) return false;
    if (!tableVars.TryGetValue(cfg.TableVar, out var tableVar)) return false;
    if (tableVar.Columns.Count == 0 || tableVar.Rows.Count == 0) return false;

    var chartFrames = slidePart.Slide.Descendants<P.GraphicFrame>()
        .Where(f => f.Graphic?.GraphicData?.Descendants<C.ChartReference>().Any() == true)
        .ToList();

    var frame = chartFrames.FirstOrDefault(f =>
        string.Equals(
            (f.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty).Trim(),
            cfg.ShapeName.Trim(),
            StringComparison.OrdinalIgnoreCase));

    if (frame == null) return false;

    var chartRef   = frame.Graphic?.GraphicData?.Descendants<C.ChartReference>().FirstOrDefault();
    var chartRelId = chartRef?.Id?.Value;
    if (string.IsNullOrWhiteSpace(chartRelId)) return false;

    if (!slidePart.TryGetPartById(chartRelId!, out var part)) return false;
    if (part is not ChartPart chartPart || chartPart.ChartSpace == null) return false;

    var labelIdx = tableVar.Columns.FindIndex(c => string.Equals(c, cfg.LabelCol, StringComparison.OrdinalIgnoreCase));
    var valueIdx = tableVar.Columns.FindIndex(c => string.Equals(c, cfg.ValueCol, StringComparison.OrdinalIgnoreCase));
    if (labelIdx < 0 || valueIdx < 0) return false;

    var labels = tableVar.Rows.Select(r => labelIdx < r.Count ? (r[labelIdx] ?? string.Empty) : string.Empty).ToList();
    var values = tableVar.Rows.Select(r => valueIdx < r.Count ? (r[valueIdx] ?? "0") : "0").ToList();

    var cat = chartPart.ChartSpace.Descendants<C.CategoryAxisData>().FirstOrDefault();
    if (cat != null)
    {
        var strRef = cat.GetFirstChild<C.StringReference>();
        if (strRef == null)
        {
            strRef = new C.StringReference();
            cat.RemoveAllChildren();
            cat.Append(strRef);
        }
        SetStringCache(strRef, labels);
    }

    var vals = chartPart.ChartSpace.Descendants<C.Values>().FirstOrDefault();
    if (vals != null)
    {
        var numRef = vals.GetFirstChild<C.NumberReference>();
        if (numRef == null)
        {
            numRef = new C.NumberReference();
            vals.RemoveAllChildren();
            vals.Append(numRef);
        }
        SetNumberCache(numRef, values);
    }

    if (!string.IsNullOrWhiteSpace(cfg.TitleVar)
        && replacements.TryGetValue(cfg.TitleVar, out var titleValue)
        && !string.IsNullOrWhiteSpace(titleValue))
    {
        var titleText = chartPart.ChartSpace.Descendants<C.Title>().SelectMany(t => t.Descendants<A.Text>()).FirstOrDefault();
        if (titleText != null)
            titleText.Text = titleValue;
    }

    chartPart.ChartSpace.Save();
    return true;
}

static async Task<bool> ApplyDynamicImageAsync(SlidePart slidePart, DynamicImageConfig cfg)
{
    if (string.IsNullOrWhiteSpace(cfg.ShapeName) || string.IsNullOrWhiteSpace(cfg.SourceUrl)) return false;

    var shape = slidePart.Slide.Descendants<P.Shape>().FirstOrDefault(s =>
        string.Equals(
            (s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty).Trim(),
            cfg.ShapeName.Trim(),
            StringComparison.OrdinalIgnoreCase));
    if (shape == null) return false;

    var xfrm = shape.ShapeProperties?.GetFirstChild<A.Transform2D>();
    var off  = xfrm?.Offset;
    var ext  = xfrm?.Extents;
    if (off == null || ext == null) return false;

    var (bytes, mimeType) = await FetchImageBytesAsync(cfg.SourceUrl);
    if (bytes.Length == 0) return false;

    var imagePartType = ToImagePartType(mimeType, cfg.SourceUrl);
    var imagePart = slidePart.AddImagePart(imagePartType);
    using (var ms = new MemoryStream(bytes))
        imagePart.FeedData(ms);

    var relId = slidePart.GetIdOfPart(imagePart);

    var nextId  = GetNextShapeId(slidePart);
    var picture = new P.Picture(
        new P.NonVisualPictureProperties(
            new P.NonVisualDrawingProperties { Id = nextId, Name = "DynImg_" + cfg.ShapeName },
            new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.BlipFill(
            new A.Blip { Embed = relId },
            new A.Stretch(new A.FillRectangle())),
        new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = off.X, Y = off.Y },
                new A.Extents { Cx = ext.Cx, Cy = ext.Cy }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

    shape.InsertAfterSelf(picture);
    shape.Remove();
    return true;
}

// ── Chart helpers ──────────────────────────────────────────────────────────────

static void SetStringCache(C.StringReference strRef, List<string> values)
{
    var cache = strRef.StringCache ?? strRef.AppendChild(new C.StringCache());
    cache.RemoveAllChildren<C.StringPoint>();
    cache.PointCount = new C.PointCount { Val = (uint)values.Count };
    for (var i = 0; i < values.Count; i++)
    {
        cache.Append(new C.StringPoint
        {
            Index        = (uint)i,
            NumericValue = new C.NumericValue(values[i] ?? string.Empty),
        });
    }
}

static void SetNumberCache(C.NumberReference numRef, List<string> values)
{
    var cache = numRef.NumberingCache ?? numRef.AppendChild(new C.NumberingCache());
    cache.RemoveAllChildren<C.NumericPoint>();
    cache.PointCount = new C.PointCount { Val = (uint)values.Count };
    for (var i = 0; i < values.Count; i++)
    {
        var parsed = double.TryParse(values[i], out var n) ? n : 0;
        cache.Append(new C.NumericPoint
        {
            Index        = (uint)i,
            NumericValue = new C.NumericValue(parsed.ToString(System.Globalization.CultureInfo.InvariantCulture)),
        });
    }
}

// ── Image helpers ──────────────────────────────────────────────────────────────

static async Task<(byte[] Bytes, string MimeType)> FetchImageBytesAsync(string sourceUrl)
{
    if (sourceUrl.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
    {
        var comma = sourceUrl.IndexOf(',');
        if (comma > 0)
        {
            var header = sourceUrl.Substring(5, comma - 5);
            var mime   = header.Split(';')[0];
            var body   = sourceUrl.Substring(comma + 1);
            return (Convert.FromBase64String(body), mime);
        }
    }

    using var http = new HttpClient();
    using var resp = await http.GetAsync(sourceUrl);
    resp.EnsureSuccessStatusCode();
    var bytesOut    = await resp.Content.ReadAsByteArrayAsync();
    var contentType = resp.Content.Headers.ContentType?.MediaType ?? "image/png";
    return (bytesOut, contentType);
}

static PartTypeInfo ToImagePartType(string mimeType, string sourceUrl)
{
    var m = (mimeType ?? string.Empty).ToLowerInvariant();
    if (m.Contains("jpeg") || m.Contains("jpg")
        || sourceUrl.EndsWith(".jpg",  StringComparison.OrdinalIgnoreCase)
        || sourceUrl.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
        return ImagePartType.Jpeg;
    if (m.Contains("gif") || sourceUrl.EndsWith(".gif", StringComparison.OrdinalIgnoreCase))
        return ImagePartType.Gif;
    if (m.Contains("bmp") || sourceUrl.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase))
        return ImagePartType.Bmp;
    return ImagePartType.Png;
}

static uint GetNextShapeId(SlidePart slidePart)
{
    var ids = slidePart.Slide.Descendants<P.NonVisualDrawingProperties>()
        .Select(nv => nv.Id?.Value ?? 0U);
    return ids.Any() ? ids.Max() + 1U : 2U;
}

// ── Token replacement ──────────────────────────────────────────────────────────

static int ReplaceTokensInSlide(SlidePart slidePart, Dictionary<string, string> replacements)
{
    var replacedCount = 0;

    foreach (var para in slidePart.Slide.Descendants<A.Paragraph>())
    {
        var texts = para.Descendants<A.Text>().ToList();
        if (texts.Count == 0) continue;

        var combined = string.Concat(texts.Select(t => t.Text ?? string.Empty));
        if (!combined.Contains("{{", StringComparison.Ordinal)) continue;

        var replacedResult = ReplaceTokenStringWithCount(combined, name => replacements.TryGetValue(name, out var v) ? v : null);
        var replaced       = replacedResult.Text;
        replacedCount     += replacedResult.Replaced;
        if (replaced == combined) continue;

        texts[0].Text = replaced;
        for (var i = 1; i < texts.Count; i++)
            texts[i].Text = string.Empty;
    }

    return replacedCount;
}

static string ReplaceTokenString(string input, Func<string, string?> resolver)
    => ReplaceTokenStringWithCount(input, resolver).Text;

static (string Text, int Replaced) ReplaceTokenStringWithCount(string input, Func<string, string?> resolver)
{
    var localCount = 0;
    var text = Regex.Replace(input, "\\{\\{(\\w+)\\}\\}", match =>
    {
        var key   = match.Groups[1].Value;
        var value = resolver(key);
        if (value == null) return match.Value;
        localCount++;
        return value;
    });
    return (text, localCount);
}

// ── Model ──────────────────────────────────────────────────────────────────────

internal sealed class PreviewPayload
{
    public Dictionary<string, string>?                          Replacements    { get; set; }
    public List<TableVariable>?                                 TableVariables  { get; set; }
    // Keyed by 1-based slide index string ("1", "2", …)
    public Dictionary<string, List<DynamicTableConfig>>?        DynamicTables   { get; set; }
    public Dictionary<string, List<DynamicImageConfig>>?        DynamicImages   { get; set; }
    public Dictionary<string, List<DynamicChartConfig>>?        DynamicCharts   { get; set; }
}

internal sealed class DynamicTableConfig
{
    public string ShapeName    { get; set; } = string.Empty;
    public string VariableName { get; set; } = string.Empty;
    public int    FromRow      { get; set; }
    public int    ToRow        { get; set; }
    public bool   AutoTokenize  { get; set; }   // engine fills header+cells by column index
    public bool   HideWhenEmpty { get; set; }
    public List<SortRule>           SortRules       { get; set; } = new();
    public List<ConditionalRowRule> ConditionalRows { get; set; } = new();
    public List<MergeRule>          MergeRules      { get; set; } = new();
}

internal sealed class SortRule
{
    public string Column    { get; set; } = string.Empty;
    public string Direction { get; set; } = "asc";   // "asc" | "desc"
}

internal sealed class ConditionalRowRule
{
    public int    RowIndex  { get; set; }   // 1-indexed, absolute table row
    public string Column    { get; set; } = string.Empty;
    public string Operator  { get; set; } = "notEmpty";
    public string Value     { get; set; } = string.Empty;
}

internal sealed class MergeRule
{
    public int    ColumnIndex { get; set; }   // 1-indexed
    public string Strategy    { get; set; } = "sameValue";
}

internal sealed class TableVariable
{
    public string?             Name    { get; set; }
    public List<string>        Columns { get; set; } = new();
    public List<List<string>>  Rows    { get; set; } = new();
}

internal sealed class DynamicImageConfig
{
    public string ShapeName { get; set; } = string.Empty;
    public string FitMode   { get; set; } = string.Empty;
    public string SourceUrl { get; set; } = string.Empty;
}

internal sealed class DynamicChartConfig
{
    public string ShapeName { get; set; } = string.Empty;
    public string TitleVar  { get; set; } = string.Empty;
    public string TableVar  { get; set; } = string.Empty;
    public string LabelCol  { get; set; } = string.Empty;
    public string ValueCol  { get; set; } = string.Empty;
}

internal sealed class PreviewStats
{
    public int SlideCount     { get; set; }
    public int ExpandedTables { get; set; }
    public int UpdatedImages  { get; set; }
    public int UpdatedCharts  { get; set; }
    public int ReplacedTokens { get; set; }
}
