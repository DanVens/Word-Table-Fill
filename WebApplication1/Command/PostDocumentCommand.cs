using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WebApplication1.Command
{
    public class PostDocumentCommand
    {
        // =========================================
        // Entry point
        // =========================================
        public MemoryStream Execute(Stream templateStream, FillDocumentDto dto)
        {
            var mem = new MemoryStream();
            templateStream.CopyTo(mem);
            mem.Position = 0;

            using var doc = WordprocessingDocument.Open(mem, true);
            var main = doc.MainDocumentPart ?? throw new InvalidOperationException("No MainDocumentPart.");

            // 0) Optional scalars (simple tag -> string)
            if (dto.scalars is { Count: > 0 })
                FillScalarsEverywhere(main, dto.scalars);    // <— use this

            // 1) Tables: prefer by tag; else by order
            if (dto.rowsByTag is { Count: > 0 })
            {
                PopulateTableByTag(main, dto.tableTag, dto.rowsByTag);
            }
            else if (dto.rowsByOrder is { Count: > 0 })
            {
                PopulateTableByOrder(main, dto.tableTag, dto.rowsByOrder);
            }

            main.Document.Save();
            mem.Position = 0;
            return mem;
        }

        // =========================================
        // FILL BY ORDER (values in array order)
        // =========================================
        public static void PopulateTableByOrder(
            MainDocumentPart main,
            string? tableTag,
            IEnumerable<IList<string>> rows,
            int? templateRowIndex = null)
        {
            var (table, tplRow1, _, tplRow2) = FindTemplateRecord(main, tableTag);

            // 1) Cache a clean copy of the template record (row or row-pair)
            var rec1 = (TableRow)tplRow1.CloneNode(true);
            TableRow? rec2 = tplRow2 != null ? (TableRow)tplRow2.CloneNode(true) : null;
            ClearHeaderFlag(rec1);
            if (rec2 != null) ClearHeaderFlag(rec2);
            
            RemoveRowsAfterTemplateRecord(table, tplRow1, tplRow2);

            // 3) Insert new rows immediately after the template record
            OpenXmlElement anchor = (OpenXmlElement)(tplRow2 ?? tplRow1);

            foreach (var values in rows)
            {
                var newR1 = (TableRow)rec1.CloneNode(true);
                TableRow? newR2 = rec2 != null ? (TableRow)rec2.CloneNode(true) : null;

                int v = 0; // next value from the JSON array
                foreach (var sdt in EnumerateSdts(newR1, newR2))
                {
                    ReplaceSdtTextPreserveFormatting(sdt, v < values.Count ? values[v++] : string.Empty);
                }

                anchor = table.InsertAfter(newR1, anchor);
                if (newR2 != null) anchor = table.InsertAfter(newR2, anchor);
            }

            // 4) Finally remove the original template record
            tplRow2?.Remove();
            tplRow1.Remove();
        }


        // =========================================
        // FILL BY TAG (dictionary per row {tag: value})
        // =========================================
        public static void PopulateTableByTag(
            MainDocumentPart main,
            string? tableTag,
            IEnumerable<IDictionary<string, string>> rows)
        {
            var (table, tplRow1, _, tplRow2) = FindTemplateRecord(main, tableTag);

            var rec1 = (TableRow)tplRow1.CloneNode(true);
            TableRow? rec2 = tplRow2 != null ? (TableRow)tplRow2.CloneNode(true) : null;
            ClearHeaderFlag(rec1);
            if (rec2 != null) ClearHeaderFlag(rec2);

            RemoveRowsAfterTemplateRecord(table, tplRow1, tplRow2);
            OpenXmlElement anchor = (OpenXmlElement)(tplRow2 ?? tplRow1);

            foreach (var inputMap in rows)
            {
                var map = new Dictionary<string, string>(inputMap, StringComparer.OrdinalIgnoreCase);

                var newR1 = (TableRow)rec1.CloneNode(true);
                FillContentControlsIn(newR1, map);

                TableRow? newR2 = null;
                if (rec2 != null)
                {
                    newR2 = (TableRow)rec2.CloneNode(true);
                    FillContentControlsIn(newR2, map);
                }

                anchor = table.InsertAfter(newR1, anchor);
                if (newR2 != null) anchor = table.InsertAfter(newR2, anchor);
            }

            tplRow2?.Remove();
            tplRow1.Remove();
        }




        // =========================================
        // Helpers (record detection, cleanup, filling)
        // =========================================
        
    private static (Table table, TableRow tplRow1, List<string?> firstRowTags, TableRow? tplRow2)
    FindTemplateRecord(MainDocumentPart main, string? tableTag)
{
    if (main?.Document?.Body == null)
        throw new InvalidOperationException("No document body.");

    // 1) Resolve the table
    Table table;
    if (!string.IsNullOrWhiteSpace(tableTag))
    {
        var wraps = main.Document.Body
            .Descendants<SdtElement>()
            .Where(s => string.Equals(GetTagOrAlias(s), tableTag, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (wraps.Count == 0)
            throw new InvalidOperationException(
                $"No content control wrapper with tag '{tableTag}' was found. Wrap *one* table with that tag.");
        if (wraps.Count > 1)
            throw new InvalidOperationException(
                $"More than one table wrapper tagged '{tableTag}' found ({wraps.Count}). Keep exactly one.");

        table = wraps[0].Descendants<Table>().FirstOrDefault()
             ?? throw new InvalidOperationException($"Wrapper '{tableTag}' did not contain a table.");
    }
    else
    {
        var allTables = main.Document.Body.Elements<Table>().ToList();
        if (allTables.Count == 0)
            throw new InvalidOperationException("No table found.");
        if (allTables.Count > 1)
            throw new InvalidOperationException(
                $"Document contains {allTables.Count} tables. Wrap the intended one and set tableTag.");
        table = allTables[0];
    }

    // 2) Pick the template row automatically
    var allRows = table.Elements<TableRow>().ToList();
    if (allRows.Count == 0)
        throw new InvalidOperationException("Table has no rows.");

    var tplRow1 = allRows
        .Select(r => new
        {
            Row = r,
            SdtCells = r.Elements<TableCell>().Count(c => c.Descendants<SdtElement>().Any()),
            Cells = r.Elements<TableCell>().Count()
        })
        .OrderByDescending(x => x.SdtCells)
        .ThenByDescending(x => x.Cells)
        .FirstOrDefault(x => x.SdtCells > 0)?.Row
        ?? throw new InvalidOperationException(
            "No row with content controls found in the chosen table.");

    // 3) If the next row also has SDTs, treat it as the second half of a two-row record
    TableRow? tplRow2 = tplRow1.NextSibling<TableRow>();
    if (tplRow2 != null && !tplRow2.Descendants<SdtElement>().Any())
        tplRow2 = null;

    // 4) Extract the tags for each cell in the chosen row (null if a cell has no CC)
    var firstRowTags = tplRow1.Elements<TableCell>()
        .Select(cell =>
            cell.Descendants<SdtElement>()
                .Select(GetTagOrAlias)
                .FirstOrDefault(t => !string.IsNullOrWhiteSpace(t)))
        .ToList();

    return (table, tplRow1, firstRowTags, tplRow2);
}

        // Delete every row that comes AFTER the template record (row or row-pair)
        private static void RemoveRowsAfterTemplateRecord(
            Table table,
            TableRow tplRow1,
            TableRow? tplRow2)
        {
            var toRemove = table.Elements<TableRow>()
                .SkipWhile(r => r != tplRow1)
                .Skip(1 + (tplRow2 != null ? 1 : 0))
                .ToList();

            foreach (var r in toRemove)
                r.Remove();
        }

        // Enumerate SDTs across a record (row1 and optional row2), in visual order.
        private static IEnumerable<SdtElement> EnumerateSdts(TableRow row1, TableRow? row2)
        {
            foreach (var sdt in row1.Descendants<SdtElement>()) yield return sdt;
            if (row2 != null)
                foreach (var sdt in row2.Descendants<SdtElement>()) yield return sdt;
        }

        // Remove header flag from a row (to prevent repeated headers for data rows).
        private static void ClearHeaderFlag(TableRow row)
        {
            var pr = row.GetFirstChild<TableRowProperties>();
            pr?.GetFirstChild<TableHeader>()?.Remove();
        }
        private static string? GetTagOrAlias(SdtElement sdt)
        {
            var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(tag)) return tag;
            return sdt.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.Value;
        }

        static void ReplaceSdtText(SdtElement sdt, string? value)
        {
            value ??= string.Empty;

            // Find the actual content container inside the SDT
            OpenXmlElement? content =
                sdt.GetFirstChild<SdtContentRun>() ??
                (OpenXmlElement?)sdt.GetFirstChild<SdtContentBlock>() ??
                sdt.GetFirstChild<SdtContentCell>();

            if (content == null) return;

            // Remove everything inside the control
            content.RemoveAllChildren();

            // Create one run with one text node, with spaces preserved
            var t = new Text(value) { Space = SpaceProcessingModeValues.Preserve };
            var r = new Run(t);

            // For block/cell SDTs we must wrap the run in a paragraph
            if (content is SdtContentRun)
            {
                content.Append(r);
            }
            else
            {
                content.Append(new Paragraph(r));
            }
        }
static void ReplaceSdtTextPreserveFormatting(SdtElement sdt, string? value)
{
    value ??= string.Empty;

    // Run-level SDT
    if (sdt.GetFirstChild<SdtContentRun>() is SdtContentRun runContent)
    {
        var firstRun = runContent.Descendants<Run>().FirstOrDefault();
        var rPr = firstRun?.RunProperties?.CloneNode(true) as RunProperties;

        runContent.RemoveAllChildren();

        var r = new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
        if (rPr != null) r.RunProperties = rPr;
        runContent.Append(r);
        return;
    }

    // Block-level SDT (paragraphs)
    if (sdt.GetFirstChild<SdtContentBlock>() is SdtContentBlock blockContent)
    {
        var firstPara = blockContent.Descendants<Paragraph>().FirstOrDefault();
        var pPr = firstPara?.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
        var firstRun = firstPara?.Descendants<Run>().FirstOrDefault();
        var rPr = firstRun?.RunProperties?.CloneNode(true) as RunProperties;

        blockContent.RemoveAllChildren();

        var p = new Paragraph();
        if (pPr != null) p.ParagraphProperties = pPr;

        var r = new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
        if (rPr != null) r.RunProperties = rPr;

        p.Append(r);
        blockContent.Append(p);
        return;
    }

    // Cell-level SDT (table cell content)
    if (sdt.GetFirstChild<SdtContentCell>() is SdtContentCell cellContent)
    {
        var firstPara = cellContent.Descendants<Paragraph>().FirstOrDefault();
        var pPr = firstPara?.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
        var firstRun = firstPara?.Descendants<Run>().FirstOrDefault();
        var rPr = firstRun?.RunProperties?.CloneNode(true) as RunProperties;

        cellContent.RemoveAllChildren();

        var p = new Paragraph();
        if (pPr != null) p.ParagraphProperties = pPr;

        var r = new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
        if (rPr != null) r.RunProperties = rPr;

        p.Append(r);
        cellContent.Append(p);
    }
}

        private static void FillContentControlsIn(OpenXmlElement scope, IDictionary<string, string> map)
        {
            foreach (var sdt in scope.Descendants<SdtElement>())
            {
                var tag = GetTagOrAlias(sdt);
                if (string.IsNullOrWhiteSpace(tag)) continue;
                if (!map.TryGetValue(tag, out var val)) continue;
                ReplaceSdtTextPreserveFormatting(sdt, val);
            }
        }
        private static void FillScalarsEverywhere(MainDocumentPart main, IDictionary<string,string> map)
        {
            if (map == null || map.Count == 0) return;

            // Body
            if (main.Document?.Body != null)
                FillContentControlsIn(main.Document.Body, map);

            // All headers
            foreach (var hp in main.HeaderParts)
                if (hp.Header != null)
                    FillContentControlsIn(hp.Header, map);

            // All footers
            foreach (var fp in main.FooterParts)
                if (fp.Footer != null)
                    FillContentControlsIn(fp.Footer, map);
        }

    }
}
