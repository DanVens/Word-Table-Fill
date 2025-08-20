using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WebApplication1
{
    public static class DocumentGenerator
    {
        public static void CreateDocument(Stream outputStream)
        {
            // Create the document in the provided stream
            using var wordDoc = WordprocessingDocument.Create(
                outputStream,
                WordprocessingDocumentType.Document,
                true
            );

            // Set up parts
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            // --- Call our static helpers ---
            // 1) Numbering (for lists)
            EnsureNumberingPart(mainPart);

            // 2) Table
            body.Append(new Paragraph(new Run(new Text("— Table Example —"))));
            AddSimpleTable(body);

            // 3) Bullets
            body.Append(new Paragraph(new Run(new Text("— Bulleted List —"))));
            AddBulletList(body);

            // 4) Hyperlink
            body.Append(new Paragraph(new Run(new Text("— Hyperlink —"))));
            AddHyperlink(mainPart, body, "https://example.com", "Go to example.com");

            // 5) Image (change path as needed)
            body.Append(new Paragraph(new Run(new Text("— Image —"))));
            AddImage(mainPart, body, @"C:/Users/Praktika/Pictures/hacker.jpg");

            // Save
            mainPart.Document.Save();
        }

        static void EnsureNumberingPart(MainDocumentPart mainPart)
        {
            if (mainPart.NumberingDefinitionsPart != null) return;
            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering(
                new AbstractNum(
                    new Level(
                        new NumberingFormat { Val = NumberFormatValues.Bullet },
                        new LevelText      { Val = "•" },
                        new ParagraphProperties(new Indentation { Left = "720", Hanging = "360" })
                    )
                ) { AbstractNumberId = 1 },
                new NumberingInstance(
                    new AbstractNumId { Val = 1 }
                ) { NumberID = 1 }
            );
        }

        static void AddSimpleTable(Body body)
        {
            var table = new Table(
                new TableProperties(
                    new TableBorders(
                        new TopBorder    { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new LeftBorder   { Val = BorderValues.Single, Size = 4 },
                        new RightBorder  { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder   { Val = BorderValues.Single, Size = 4 }
                    )
                )
            );

            // Header row
            var headerRow = new TableRow();
            headerRow.Append(
                MakeCell("Header A", true),
                MakeCell("Header B", true)
            );
            table.Append(headerRow);

            // Data row
            var dataRow = new TableRow();
            dataRow.Append(
                MakeCell("Cell A2"),
                MakeCell("Cell B2")
            );
            table.Append(dataRow);

            body.Append(table);
        }

        static TableCell MakeCell(string text, bool isBold = false)
        {
            var run = new Run();
            if (isBold)
                run.Append(new RunProperties(new Bold()));
            run.Append(new Text(text));

            var cellProps = new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2400" }
            );

            return new TableCell(new Paragraph(run), cellProps);
        }

        static void AddBulletList(Body body)
        {
            foreach (var item in new[] { "First item", "Second item", "Third item" })
            {
                // Create new numbering props for *each* paragraph
                var numberingProps = new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId             { Val = 1 }
                );

                var pPr = new ParagraphProperties(numberingProps);
                var run = new Run(new Text(item));
                var p   = new Paragraph(pPr, run);

                body.Append(p);
            }
        }


        static void AddHyperlink(MainDocumentPart mainPart, Body body, string url, string displayText)
        {
            var linkRel = mainPart.AddHyperlinkRelationship(new Uri(url), true);
            var hyperlink = new Hyperlink(
                new Run(
                    new RunProperties(
                        new RunStyle { Val = "Hyperlink" },
                        new Color    { ThemeColor = ThemeColorValues.Hyperlink }
                    ),
                    new Text(displayText) { Space = SpaceProcessingModeValues.Preserve }
                )
            ) { Id = linkRel.Id };

            body.Append(new Paragraph(hyperlink));
        }

        static void AddImage(MainDocumentPart mainPart, Body body, string imagePath)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using var stream = File.OpenRead(imagePath);
            imagePart.FeedData(stream);

            string rId = mainPart.GetIdOfPart(imagePart);

            var drawing = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = 990000L, Cy = 792000L },
                    new DW.DocProperties { Id = 1U, Name = Path.GetFileName(imagePath) },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }
                    ),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = Path.GetFileName(imagePath) },
                                    new PIC.NonVisualPictureDrawingProperties()
                                ),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = rId },
                                    new A.Stretch(new A.FillRectangle())
                                ),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = 990000L, Cy = 792000L }
                                    ),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                )
                            )
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U }
            );

            body.Append(new Paragraph(new Run(drawing)));
        }
    }
}
