using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Fonts;
using OpenXmlDoc = DocumentFormat.OpenXml.Wordprocessing;
using MigraDocDoc = MigraDoc.DocumentObjectModel;

namespace net.nick4name.MergeService.Docx {

   public class MigraConverterHelper {
      private readonly string _fileToMerge;
      private readonly string _fileMerged;

      public MigraConverterHelper(string fileToMerge, string fileMerged) {
         _fileToMerge = fileToMerge;
         _fileMerged = fileMerged;
      }

      public MigraDocDoc.Document ConvertWordToPdf<T>(List<Placeholder> placeholders, T context) {
         GlobalFontSettings.FontResolver = new DejaVuFontResolver();

         var doc = new MigraDocDoc.Document();
         var section = doc.AddSection();
         section.PageSetup.PageFormat = PageFormat.A4;

         var style = doc.Styles["Normal"]!;
         style.Font.Name = "DejaVuSans";
         style.Font.Bold = true;

         using var wordDoc = WordprocessingDocument.Open(_fileToMerge, false);
         var body = wordDoc.MainDocumentPart!.Document.Body;

         ApplyPageSetup(wordDoc, section);

         foreach (var para in body!.Elements<OpenXmlDoc.Paragraph>()) {
            var migraPara = TryConvertListParagraph(wordDoc, para, section) ?? section.AddParagraph();
            ApplyParagraphFormatting(para, migraPara, section);
            ConvertRuns(para, migraPara, context);
         }

         var pdfRenderer = new PdfDocumentRenderer(true) { Document = doc };
         pdfRenderer.RenderDocument();
         pdfRenderer.PdfDocument.Save(_fileMerged);

         return doc;
      }

      private MigraDocDoc.Paragraph? TryConvertListParagraph(WordprocessingDocument wordDoc, OpenXmlDoc.Paragraph para, Section section) {
         var numberingProps = para.GetFirstChild<NumberingProperties>();
         if (numberingProps?.NumberingId?.Val == null)
            return null;

         int numId = numberingProps.NumberingId.Val.Value;
         int ilvl = numberingProps.NumberingLevelReference?.Val?.Value ?? 0;

         var numberingPart = wordDoc.MainDocumentPart?.NumberingDefinitionsPart;
         if (numberingPart == null)
            return null;

         var numberingInstance = numberingPart.Numbering.Elements<NumberingInstance>()
             .FirstOrDefault(n => n.NumberID?.Value == numId);
         if (numberingInstance?.AbstractNumId?.Val == null)
            return null;

         int abstractNumId = numberingInstance.AbstractNumId.Val.Value;

         var abstractNum = numberingPart.Numbering.Elements<AbstractNum>()
             .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
         if (abstractNum == null)
            return null;

         var level = abstractNum.Elements<Level>()
             .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);
         if (level?.NumberingFormat?.Val == null)
            return null;

         var format = level.NumberingFormat.Val.Value;

         var listType = format == NumberFormatValues.Bullet
             ? ilvl switch {
                0 => ListType.BulletList1,
                1 => ListType.BulletList2,
                _ => ListType.BulletList3
             }
             : ilvl switch {
                0 => ListType.NumberList1,
                1 => ListType.NumberList2,
                _ => ListType.NumberList3
             };

         var migraPara = section.AddParagraph();
         migraPara.Style = "List";
         migraPara.Format.ListInfo = new ListInfo { ListType = listType };
         return migraPara;
      }

      private void ConvertRuns<T>(OpenXmlDoc.Paragraph para, MigraDoc.DocumentObjectModel.Paragraph migraPara, T context) {
         var runs = para.Elements<Run>().ToList();
         int i = 0;

         while (i < runs.Count) {
            var run = runs[i];
            var fldChar = run.GetFirstChild<FieldChar>();

            if (fldChar?.FieldCharType! == FieldCharValues.Begin) {
               int startIndex = i;
               int endIndex = -1;
               string? fieldName = null;

               for (int j = i + 1; j < runs.Count; j++) {
                  var r = runs[j];
                  var fieldCode = r.GetFirstChild<FieldCode>();
                  if (fieldCode?.Text.Contains("MERGEFIELD") == true) {
                     var tokens = fieldCode.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                     if (tokens.Length >= 2)
                        fieldName = tokens[1];
                  }

                  var endChar = r.GetFirstChild<FieldChar>();
                  if (endChar?.FieldCharType! == FieldCharValues.End) {
                     endIndex = j;
                     break;
                  }
               }

               if (fieldName != null && endIndex > startIndex) {
                  var prop = typeof(T).GetProperty(fieldName);
                  string value = prop?.GetValue(context)?.ToString() ?? "";

                  for (int j = startIndex; j <= endIndex; j++)
                     para.RemoveChild(runs[j]);

                  para.InsertAt(new Run(new OpenXmlDoc.Text(value)), startIndex);

                  var text = migraPara.AddFormattedText(value);
                  ApplyTextFormatting(run, text);

                  runs = para.Elements<Run>().ToList();
                  i = startIndex + 1;
                  continue;
               }
            }

            if (!run.InnerText.Contains("MERGEFIELD")) {
               var text = migraPara.AddFormattedText(run.InnerText);
               ApplyTextFormatting(run, text);
            }

            i++;
         }
      }

      // Applies character formatting from a Word Run to a MigraDoc FormattedText object
      static void ApplyTextFormatting(Run run, FormattedText text) {
         var props = run.RunProperties;
         if (props != null) {
            text.Bold = props.Bold != null;
            text.Italic = props.Italic != null;
            text.Underline = props.Underline != null ? MigraDoc.DocumentObjectModel.Underline.Single : MigraDoc.DocumentObjectModel.Underline.None;
            text.Size = props.FontSize != null ? Convert.ToDouble(props!.FontSize.Val!) / 2 : 12;

            if (props.Color != null && !string.IsNullOrEmpty(props!.Color.Val!)) {
               text.Color = ConvertHexToColor(props!.Color.Val!);
            }

            if (props.RunFonts != null && props!.RunFonts.Ascii! != null) {
               text.Font.Name = props!.RunFonts.Ascii!?.Value ?? "Times New Roman";
            }
         }
      }

      // Applies page margins from the Word document to the MigraDoc section
      static void ApplyPageSetup(WordprocessingDocument wordDoc, Section section) {
         // Set margins to 1 inch by default, or use the values from Word if available
         var margins = wordDoc.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()?.GetFirstChild<PageMargin>();
         if (margins != null) {
            section.PageSetup.LeftMargin = Unit.FromPoint(margins.Left! / 20);
            section.PageSetup.RightMargin = Unit.FromPoint(margins.Right! / 20);
            section.PageSetup.TopMargin = Unit.FromPoint(margins.Top! / 20);
            section.PageSetup.BottomMargin = Unit.FromPoint(margins.Bottom! / 20);
         } else {
            section.PageSetup.LeftMargin = Unit.FromInch(1);   // Default 1 inch margin
            section.PageSetup.RightMargin = Unit.FromInch(1);  // Default 1 inch margin
            section.PageSetup.TopMargin = Unit.FromInch(1);    // Default 1 inch margin
            section.PageSetup.BottomMargin = Unit.FromInch(1); // Default 1 inch margin
         }
      }

      static void ApplyParagraphFormatting(DocumentFormat.OpenXml.Wordprocessing.Paragraph para, MigraDoc.DocumentObjectModel.Paragraph migraPara, Section section) {
         // Get the paragraph properties.
         var props = para.ParagraphProperties;
         if (props != null) {
            // Apply justification (alignments)
            if (props.Justification != null) {
               if (props!.Justification.Val!.Value == JustificationValues.Center) {
                  migraPara.Format.Alignment = ParagraphAlignment.Center;
               } else if (props!.Justification.Val!.Value == JustificationValues.Right) {
                  migraPara.Format.Alignment = ParagraphAlignment.Right;
               } else if (props!.Justification.Val!.Value == JustificationValues.Both) {
                  migraPara.Format.Alignment = ParagraphAlignment.Justify;
               } else {
                  migraPara.Format.Alignment = ParagraphAlignment.Left;
               }
            }

            // Apply spacing (before, after, and line spacing)
            if (props.SpacingBetweenLines != null) {
               migraPara.Format.SpaceBefore = Unit.FromPoint(props!.SpacingBetweenLines.Before! != null ? Convert.ToDouble(props!.SpacingBetweenLines.Before!) / 20.0 : 0);
               migraPara.Format.SpaceAfter = Unit.FromPoint(props!.SpacingBetweenLines.After! != null ? Convert.ToDouble(props!.SpacingBetweenLines.After!) / 20.0 : 0);
               migraPara.Format.LineSpacing = Unit.FromPoint(props!.SpacingBetweenLines.Line! != null ? Convert.ToDouble(props!.SpacingBetweenLines.Line!) / 20.0 : 0);
            }

            // Apply indentation (respect page width and margins)
            if (props.Indentation != null) {
               double firstLineIndent = props!.Indentation.FirstLine! != null ? Convert.ToDouble(props!.Indentation.FirstLine!) : 0;

               // Ensure the left indent doesn't exceed the available space (set to 0 because indentation is already applied at document level apart from first line).
               migraPara.Format.LeftIndent = Unit.FromPoint(0);
               migraPara.Format.RightIndent = Unit.FromPoint(0);
               migraPara.Format.FirstLineIndent = Unit.FromPoint(firstLineIndent);
            }

            // Apply tab stops
            if (props.Tabs != null) {
               foreach (var tab in props.Tabs.Elements<DocumentFormat.OpenXml.Wordprocessing.TabStop>()) {
                  // For each tab, align according to its property.
                  TabAlignment align;
                  if (tab.Val! == TabStopValues.Center) {
                     align = TabAlignment.Center;
                  } else if (tab.Val! == TabStopValues.Right) {
                     align = TabAlignment.Right;
                  } else {
                     align = TabAlignment.Left;
                  }
                  // Add a tab stop.
                  migraPara.Format.TabStops.AddTabStop(Unit.FromPoint(tab.Position!), align);
               }
            }

            // Apply shading (background color)
            if (props.Shading != null && !string.IsNullOrEmpty(props!.Shading.Fill!)) {
               migraPara.Format.Shading.Color = ConvertHexToColor(props!.Shading.Fill!);
            }

            // Apply paragraph borders
            if (props.ParagraphBorders != null) {
               var border = props!.ParagraphBorders.TopBorder!;
               if (border != null && border.Color != null) {
                  migraPara.Format.Borders.Top.Color = ConvertHexToColor(border.Color!);
               }
            }
         }
      }

      // Converts a hex color code (e.g., "FF0000") to a MigraDoc Color
      static MigraDoc.DocumentObjectModel.Color ConvertHexToColor(string hex) {
         if (hex.Length == 6) {
            return new MigraDoc.DocumentObjectModel.Color(
                Convert.ToByte(hex.Substring(0, 2), 16),
                Convert.ToByte(hex.Substring(2, 2), 16),
                Convert.ToByte(hex.Substring(4, 2), 16)
            );
         }
         return Colors.Black;
      }

      // ApplyParagraphFormatting and ApplyTextFormatting remain unchanged
   }

}
