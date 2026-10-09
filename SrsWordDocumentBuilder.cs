using System;
using System.IO;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SrsAi.Functions
{
    public static class SrsWordDocumentBuilder
    {
        private const string PrimaryColorHex = "1F4E78"; // Dark Navy Blue
        private const string SecondaryColorHex = "2F5597"; // Slate Blue
        private const string TableHeaderBgHex = "1F4E78"; // Table Header Fill
        private const string LightBgHex = "F2F4F7"; // Callout / Alternate Shading

        public static byte[] BuildWordDocument(string jsonContent, string meetingId)
        {
            using MemoryStream ms = new MemoryStream();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                // Add Document Title Banner
                AddTitle(body, "SOFTWARE REQUIREMENTS SPECIFICATION");
                AddSubtitle(body, $"Project Meeting ID: {meetingId}  |  Draft Version for Human Review");

                AddDividerLine(body);

                // Clean raw JSON string (strip markdown fences or prose)
                string cleanedJson = ExtractJsonPayload(jsonContent);

                string project = "Unknown";
                JsonElement root = default;
                bool isParsed = false;

                try
                {
                    using JsonDocument jsonDoc = JsonDocument.Parse(cleanedJson);
                    root = jsonDoc.RootElement.Clone();
                    isParsed = true;

                    if (root.TryGetProperty("project", out var pElem))
                    {
                        project = pElem.GetString() ?? "Unknown";
                    }
                }
                catch
                {
                    // Fallback if parsing fails
                }

                // Executive Summary Metadata Table
                AddMetadataTable(body, project, meetingId, DateTime.UtcNow.ToString("dd-MMM-yyyy HH:mm UTC"));

                if (isParsed)
                {
                    // Section 1: Extracted Functional Requirements
                    AddHeading1(body, "1. Functional Requirements & System Impact Analysis");

                    if (root.TryGetProperty("requirements", out var reqsElem) && reqsElem.ValueKind == JsonValueKind.Array)
                    {
                        int reqIndex = 1;
                        foreach (var req in reqsElem.EnumerateArray())
                        {
                            string id = req.TryGetProperty("id", out var idElem) ? idElem.GetString() ?? $"REQ-{reqIndex:D3}" : $"REQ-{reqIndex:D3}";
                            string title = req.TryGetProperty("title", out var titleElem) ? titleElem.GetString() ?? "Requirement" : "Requirement";
                            string desiredBehavior = req.TryGetProperty("desiredBehavior", out var dbElem) ? dbElem.GetString() ?? "" : "";
                            string sourceMeeting = req.TryGetProperty("sourceMeeting", out var smElem) ? smElem.GetString() ?? "" : "";
                            string classification = req.TryGetProperty("classification", out var cElem) ? cElem.GetString() ?? "NEW" : "NEW";
                            string existingSystemComparison = req.TryGetProperty("existingSystemComparison", out var escElem) ? escElem.GetString() ?? "" : "";

                            AddHeading2(body, $"{id}: {title}  [{classification.ToUpper()}]");

                            if (!string.IsNullOrWhiteSpace(desiredBehavior))
                            {
                                AddCalloutBox(body, "Desired Behavior", desiredBehavior);
                            }

                            if (!string.IsNullOrWhiteSpace(sourceMeeting))
                            {
                                AddLabeledParagraph(body, "Source / Transcript Evidence: ", sourceMeeting);
                            }

                            if (!string.IsNullOrWhiteSpace(existingSystemComparison))
                            {
                                AddLabeledParagraph(body, "Existing System Impact & Analysis: ", existingSystemComparison);
                            }

                            // Acceptance Criteria
                            if (req.TryGetProperty("acceptanceCriteria", out var acElem) && acElem.ValueKind == JsonValueKind.Array && acElem.GetArrayLength() > 0)
                            {
                                AddBoldLabel(body, "Acceptance Criteria:");
                                foreach (var criterion in acElem.EnumerateArray())
                                {
                                    AddBulletPoint(body, criterion.GetString() ?? "");
                                }
                            }

                            // Component Impacts
                            if (req.TryGetProperty("impacts", out var impElem) && impElem.ValueKind == JsonValueKind.Object)
                            {
                                AddImpactsTable(body, impElem);
                            }

                            AddSpacing(body);
                            reqIndex++;
                        }
                    }
                    else
                    {
                        AddParagraph(body, "No explicit requirements extracted from this transcript.");
                    }

                    // Section 2: Open Questions
                    AddHeading1(body, "2. Open Questions & Clarifications Needed");
                    if (root.TryGetProperty("openQuestions", out var oqElem) && oqElem.ValueKind == JsonValueKind.Array && oqElem.GetArrayLength() > 0)
                    {
                        foreach (var question in oqElem.EnumerateArray())
                        {
                            AddBulletPoint(body, question.GetString() ?? "");
                        }
                    }
                    else
                    {
                        AddParagraph(body, "No open questions identified.");
                    }

                    // Section 3: Technical Assumptions
                    AddHeading1(body, "3. Technical Assumptions");
                    if (root.TryGetProperty("assumptions", out var assElem) && assElem.ValueKind == JsonValueKind.Array && assElem.GetArrayLength() > 0)
                    {
                        foreach (var assumption in assElem.EnumerateArray())
                        {
                            AddBulletPoint(body, assumption.GetString() ?? "");
                        }
                    }
                    else
                    {
                        AddParagraph(body, "No explicit technical assumptions logged.");
                    }
                }
                else
                {
                    // Clean human-readable fallback if raw JSON parsing failed
                    AddHeading1(body, "1. Extracted Requirements Summary");
                    string cleanText = FormatTextFallback(jsonContent);
                    AddParagraph(body, cleanText);
                }

                mainPart.Document.Save();
            }

            return ms.ToArray();
        }

        private static string ExtractJsonPayload(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "{}";

            int firstBrace = input.IndexOf('{');
            int lastBrace = input.LastIndexOf('}');

            if (firstBrace >= 0 && lastBrace > firstBrace)
            {
                return input.Substring(firstBrace, lastBrace - firstBrace + 1).Trim();
            }

            return input.Trim();
        }

        private static string FormatTextFallback(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return raw;

            // Strip raw JSON formatting characters for clean document reading
            return raw.Replace("{", "")
                      .Replace("}", "")
                      .Replace("[", "")
                      .Replace("]", "")
                      .Replace("\"", "")
                      .Replace("  ", " ")
                      .Trim();
        }

        #region OpenXML Helper Styling Methods

        private static void AddTitle(Body body, string text)
        {
            Paragraph p = body.AppendChild(new Paragraph());
            Run r = p.AppendChild(new Run());
            r.AppendChild(new Text(text));

            RunProperties rProp = r.GetTypeProperties();
            rProp.AppendChild(new Bold());
            rProp.AppendChild(new FontSize() { Val = "36" }); // 18pt
            rProp.AppendChild(new Color() { Val = PrimaryColorHex });
            rProp.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            ParagraphProperties pProp = p.GetTypeProperties();
            pProp.AppendChild(new SpacingBetweenLines() { After = "60" });
        }

        private static void AddSubtitle(Body body, string text)
        {
            Paragraph p = body.AppendChild(new Paragraph());
            Run r = p.AppendChild(new Run());
            r.AppendChild(new Text(text));

            RunProperties rProp = r.GetTypeProperties();
            rProp.AppendChild(new FontSize() { Val = "22" }); // 11pt
            rProp.AppendChild(new Color() { Val = "595959" }); // Muted grey
            rProp.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            ParagraphProperties pProp = p.GetTypeProperties();
            pProp.AppendChild(new SpacingBetweenLines() { After = "200" });
        }

        private static void AddHeading1(Body body, string text)
        {
            Paragraph p = body.AppendChild(new Paragraph());
            Run r = p.AppendChild(new Run());
            r.AppendChild(new Text(text));

            RunProperties rProp = r.GetTypeProperties();
            rProp.AppendChild(new Bold());
            rProp.AppendChild(new FontSize() { Val = "28" }); // 14pt
            rProp.AppendChild(new Color() { Val = PrimaryColorHex });
            rProp.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            ParagraphProperties pProp = p.GetTypeProperties();
            pProp.AppendChild(new SpacingBetweenLines() { Before = "300", After = "120" });
        }

        private static void AddHeading2(Body body, string text)
        {
            Paragraph p = body.AppendChild(new Paragraph());
            Run r = p.AppendChild(new Run());
            r.AppendChild(new Text(text));

            RunProperties rProp = r.GetTypeProperties();
            rProp.AppendChild(new Bold());
            rProp.AppendChild(new FontSize() { Val = "24" }); // 12pt
            rProp.AppendChild(new Color() { Val = SecondaryColorHex });
            rProp.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            ParagraphProperties pProp = p.GetTypeProperties();
            pProp.AppendChild(new SpacingBetweenLines() { Before = "200", After = "100" });
        }

        private static void AddParagraph(Body body, string text)
        {
            Paragraph p = body.AppendChild(new Paragraph());
            Run r = p.AppendChild(new Run());
            r.AppendChild(new Text(text));

            RunProperties rProp = r.GetTypeProperties();
            rProp.AppendChild(new FontSize() { Val = "22" }); // 11pt
            rProp.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            ParagraphProperties pProp = p.GetTypeProperties();
            pProp.AppendChild(new SpacingBetweenLines() { After = "120" });
        }

        private static void AddBoldLabel(Body body, string text)
        {
            Paragraph p = body.AppendChild(new Paragraph());
            Run r = p.AppendChild(new Run());
            r.AppendChild(new Text(text));

            RunProperties rProp = r.GetTypeProperties();
            rProp.AppendChild(new Bold());
            rProp.AppendChild(new FontSize() { Val = "22" });
            rProp.AppendChild(new Color() { Val = "262626" });
            rProp.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            ParagraphProperties pProp = p.GetTypeProperties();
            pProp.AppendChild(new SpacingBetweenLines() { Before = "100", After = "60" });
        }

        private static void AddLabeledParagraph(Body body, string label, string content)
        {
            Paragraph p = body.AppendChild(new Paragraph());

            Run rLabel = p.AppendChild(new Run());
            rLabel.AppendChild(new Text(label));
            RunProperties rProp1 = rLabel.GetTypeProperties();
            rProp1.AppendChild(new Bold());
            rProp1.AppendChild(new FontSize() { Val = "22" });
            rProp1.AppendChild(new Color() { Val = "262626" });
            rProp1.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            Run rContent = p.AppendChild(new Run());
            rContent.AppendChild(new Text(content));
            RunProperties rProp2 = rContent.GetTypeProperties();
            rProp2.AppendChild(new FontSize() { Val = "22" });
            rProp2.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            ParagraphProperties pProp = p.GetTypeProperties();
            pProp.AppendChild(new SpacingBetweenLines() { After = "120" });
        }

        private static void AddBulletPoint(Body body, string text)
        {
            Paragraph p = body.AppendChild(new Paragraph());
            Run r = p.AppendChild(new Run());
            r.AppendChild(new Text($"•  {text}"));

            RunProperties rProp = r.GetTypeProperties();
            rProp.AppendChild(new FontSize() { Val = "22" });
            rProp.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            ParagraphProperties pProp = p.GetTypeProperties();
            pProp.AppendChild(new Indentation() { Left = "360" });
            pProp.AppendChild(new SpacingBetweenLines() { After = "60" });
        }

        private static void AddCalloutBox(Body body, string title, string text)
        {
            Table table = body.AppendChild(new Table());

            TableProperties tblProp = new TableProperties(
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 100% width
                new TableBorders(
                    new LeftBorder() { Val = BorderValues.Single, Size = 24, Color = PrimaryColorHex },
                    new TopBorder() { Val = BorderValues.None },
                    new RightBorder() { Val = BorderValues.None },
                    new BottomBorder() { Val = BorderValues.None }
                )
            );
            table.AppendChild(tblProp);

            TableRow row = table.AppendChild(new TableRow());
            TableCell cell = row.AppendChild(new TableCell());

            TableCellProperties cellProp = new TableCellProperties(
                new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = LightBgHex }
            );
            cell.AppendChild(cellProp);

            Paragraph pTitle = cell.AppendChild(new Paragraph());
            Run rTitle = pTitle.AppendChild(new Run());
            rTitle.AppendChild(new Text(title.ToUpper()));
            RunProperties rPropTitle = rTitle.GetTypeProperties();
            rPropTitle.AppendChild(new Bold());
            rPropTitle.AppendChild(new FontSize() { Val = "20" });
            rPropTitle.AppendChild(new Color() { Val = PrimaryColorHex });
            rPropTitle.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            Paragraph pText = cell.AppendChild(new Paragraph());
            Run rText = pText.AppendChild(new Run());
            rText.AppendChild(new Text(text));
            RunProperties rPropText = rText.GetTypeProperties();
            rPropText.AppendChild(new FontSize() { Val = "22" });
            rPropText.AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            AddSpacing(body);
        }

        private static void AddMetadataTable(Body body, string project, string meetingId, string dateStr)
        {
            Table table = body.AppendChild(new Table());

            TableProperties tblProp = new TableProperties(
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 4, Color = "D9D9D9" },
                    new BottomBorder() { Val = BorderValues.Single, Size = 4, Color = "D9D9D9" },
                    new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4, Color = "E5E5E5" },
                    new LeftBorder() { Val = BorderValues.None },
                    new RightBorder() { Val = BorderValues.None },
                    new InsideVerticalBorder() { Val = BorderValues.None }
                )
            );
            table.AppendChild(tblProp);

            AddTableRow(table, "Project Name", project, true);
            AddTableRow(table, "Meeting Reference ID", meetingId, false);
            AddTableRow(table, "Generated Date", dateStr, false);
            AddTableRow(table, "Document Stage", "DRAFT PROPOSAL (Awaiting Review)", false);

            AddSpacing(body);
        }

        private static void AddImpactsTable(Body body, JsonElement impactsElem)
        {
            AddBoldLabel(body, "System Component Impacts:");

            Table table = body.AppendChild(new Table());

            TableProperties tblProp = new TableProperties(
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 6, Color = PrimaryColorHex },
                    new BottomBorder() { Val = BorderValues.Single, Size = 6, Color = PrimaryColorHex },
                    new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4, Color = "E5E5E5" },
                    new LeftBorder() { Val = BorderValues.None },
                    new RightBorder() { Val = BorderValues.None },
                    new InsideVerticalBorder() { Val = BorderValues.None }
                )
            );
            table.AppendChild(tblProp);

            // Header Row
            TableRow headerRow = table.AppendChild(new TableRow());
            TableCell c1 = headerRow.AppendChild(new TableCell());
            c1.AppendChild(new TableCellProperties(new Shading() { Fill = TableHeaderBgHex }));
            Paragraph p1 = c1.AppendChild(new Paragraph());
            Run r1 = p1.AppendChild(new Run(new Text("Component Layer")));
            r1.GetTypeProperties().AppendChild(new Bold());
            r1.GetTypeProperties().AppendChild(new Color() { Val = "FFFFFF" });
            r1.GetTypeProperties().AppendChild(new FontSize() { Val = "20" });

            TableCell c2 = headerRow.AppendChild(new TableCell());
            c2.AppendChild(new TableCellProperties(new Shading() { Fill = TableHeaderBgHex }));
            Paragraph p2 = c2.AppendChild(new Paragraph());
            Run r2 = p2.AppendChild(new Run(new Text("Impact Summary & Recommendations")));
            r2.GetTypeProperties().AppendChild(new Bold());
            r2.GetTypeProperties().AppendChild(new Color() { Val = "FFFFFF" });
            r2.GetTypeProperties().AppendChild(new FontSize() { Val = "20" });

            foreach (var prop in impactsElem.EnumerateObject())
            {
                AddTableRow(table, Capitalize(prop.Name), prop.Value.GetString() ?? "N/A", false);
            }

            AddSpacing(body);
        }

        private static void AddTableRow(Table table, string col1, string col2, bool isHeaderBg)
        {
            TableRow row = table.AppendChild(new TableRow());

            TableCell c1 = row.AppendChild(new TableCell());
            if (isHeaderBg) c1.AppendChild(new TableCellProperties(new Shading() { Fill = LightBgHex }));
            Paragraph p1 = c1.AppendChild(new Paragraph());
            Run r1 = p1.AppendChild(new Run(new Text(col1)));
            r1.GetTypeProperties().AppendChild(new Bold());
            r1.GetTypeProperties().AppendChild(new FontSize() { Val = "22" });
            r1.GetTypeProperties().AppendChild(new RunFonts() { Ascii = "Segoe UI" });

            TableCell c2 = row.AppendChild(new TableCell());
            if (isHeaderBg) c2.AppendChild(new TableCellProperties(new Shading() { Fill = LightBgHex }));
            Paragraph p2 = c2.AppendChild(new Paragraph());
            Run r2 = p2.AppendChild(new Run(new Text(col2)));
            r2.GetTypeProperties().AppendChild(new FontSize() { Val = "22" });
            r2.GetTypeProperties().AppendChild(new RunFonts() { Ascii = "Segoe UI" });
        }

        private static void AddDividerLine(Body body)
        {
            Paragraph p = body.AppendChild(new Paragraph());
            ParagraphProperties pProp = p.GetTypeProperties();
            pProp.AppendChild(new ParagraphBorders(new BottomBorder() { Val = BorderValues.Single, Size = 12, Color = PrimaryColorHex }));
            pProp.AppendChild(new SpacingBetweenLines() { After = "200" });
        }

        private static void AddSpacing(Body body)
        {
            Paragraph p = body.AppendChild(new Paragraph());
            p.GetTypeProperties().AppendChild(new SpacingBetweenLines() { After = "160" });
        }

        private static RunProperties GetTypeProperties(this Run run)
        {
            if (run.RunProperties == null)
            {
                run.RunProperties = new RunProperties();
            }
            return run.RunProperties;
        }

        private static ParagraphProperties GetTypeProperties(this Paragraph paragraph)
        {
            if (paragraph.ParagraphProperties == null)
            {
                paragraph.ParagraphProperties = new ParagraphProperties();
            }
            return paragraph.ParagraphProperties;
        }

        private static string Capitalize(string str)
        {
            if (string.IsNullOrEmpty(str)) return str;
            return char.ToUpper(str[0]) + str.Substring(1);
        }

        #endregion
    }
}
