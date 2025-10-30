using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Fonts;
using System.Collections.Generic;

namespace net.nick4name.MergeService.Docx {
   internal class MergeDocx<T> : IMergeDocType where T : class {
      string? _filename = null;
      string? _maskFileMerged = null;
      string? _fileMerged = null;

      /// <summary>
      /// byte[] che rappresenta il contenuto del documento di testo template.
      /// Proprietà non supportata. Referenziare il documento template tramite la proprietà Filename.
      /// </summary>
      /// <exception cref="NotImplementedException">Proprietà non supportata.</exception>
      public byte[] SourceContent { set => throw new NotImplementedException(); }

      /// <summary>
      /// Nome file del documento template.
      /// </summary>
      public string Filename {
         set {
            _filename = value;
         }
      }

      /// <summary>
      /// Nome che assumerà il file generato dopo il merge.
      /// Può essere un nome fisso oppure una maschera con parti variabili.
      /// Le parti variabili sono rappresentate dai placeholders {NomeProprietà},
      /// dove NomeProprietà è il nome di una proprietà dell'istanza generica T.
      /// </summary>
      /// <remarks>
      /// Supponento che l'istanza generica T abbia proprietà "Nome" con valore "Mario", e "Cognome" con valore "Rossi",
      /// e che MaskFileMerged sia impostato a "Lettera_{Nome}_{Cognome}.pdf", il nome file generato dopo il merge sarà 
      /// "Lettera_Mario_Rossi.pdf".
      /// </remarks>
      public string MaskFileMerged { set => _maskFileMerged = value; }

      /// <summary>
      /// Nome del file generato dopo il merge in accordo con la definizione in MaskFileMerged.
      /// </summary>
      public string FileMerged { get => _fileMerged!; }

      /// <summary>
      /// Esegue il merge di Filename con i dati dell'istanza generica T.
      /// Scriverà il file risultante in MaskFileMerged.
      /// </summary>
      /// <typeparam name="T">Classe generica che rappresenta un'istanza di tabella o vista di db.</typeparam>
      /// <param name="context">Istanza di tabella o vista di db.</param>
      /// <returns>byte[] del file prodotto in MaskFileMerged.</returns>
      public byte[] ExecuteMerge<T>(T context) {
         if (string.IsNullOrEmpty(_filename)) {
            throw new InvalidOperationException("Filename non impostato per l'operazione di merge.");
         }

         if (string.IsNullOrEmpty(_maskFileMerged)) {
            throw new InvalidOperationException("MaskFileMerged non impostato per l'operazione di merge.");
         }

         // lista dei placeholders del documento
         List<Placeholder> placeholders = ExtractPlaceholders();

         // esegue il merge
         MigraDoc.DocumentObjectModel.Document doc = ConvertWordToPdf(placeholders, context);

         return File.ReadAllBytes(_maskFileMerged!);
      }

      // Converts a Word document (.docx) to a MigraDoc Document and saves it as a PDF
      public MigraDoc.DocumentObjectModel.Document ConvertWordToPdf<T>(List<Placeholder> placeholders, T context) {
         GlobalFontSettings.FontResolver = new DejaVuFontResolver();

         var doc = new MigraDoc.DocumentObjectModel.Document();
         var section = doc.AddSection();
         section.PageSetup.PageFormat = PageFormat.A4; //A4 page
         var style = doc.Styles["Normal"]!;
         style.Font.Name = "DejaVuSans";
         style.Font.Bold = true;

         // Open the Word document in read-only mode
         using (var wordDoc = WordprocessingDocument.Open(_filename!, false)) {
            var body = wordDoc.MainDocumentPart!.Document.Body;
            ApplyPageSetup(wordDoc, section); // Apply page margins from Word to MigraDoc

            // Process each paragraph in the Word document
            foreach (var para in body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()) {
               // Add a paragraph
               var migraPara = section.AddParagraph();
               ApplyParagraphFormatting(para, migraPara, section); // Apply paragraph-level formatting

               // Process each run (text span) within the paragraph
               foreach (var run in para.Elements<Run>()) {

                  // Skip MERGEFIELD placeholders (Avoid duplication)
                  if (!run.InnerText.Contains("MERGEFIELD")) {
                     // Add formatted text
                     var text = migraPara.AddFormattedText(run.InnerText);
                     ApplyTextFormatting(run, text); // Apply text-level formatting
                  } else {
                     // Extract the placeholder name from the run
                     string[] plc = run.InnerText.Replace(" MERGEFIELD ", "").Split(' ', StringSplitOptions.RemoveEmptyEntries);
                     string ph = plc[0];
                     Placeholder? plh = null;
                     if (placeholders.Any(p => p.Name == ph)) {
                        plh = placeholders.Where(p => p.Name == ph).FirstOrDefault();
                     }

                     // costruisce il placeholder secondo la sintassi del documento
                     string placeholder = " MERGEFIELD " + plh!.Name;

                     // ottiene il valore corrispondente alla colonna avente per nome il placeholder
                     var prop = typeof(T).GetProperty(ph);
                     string value = prop?.GetValue(context)?.ToString() ?? "";

                     string txt = run.InnerText.Replace(placeholder, value);

                     var text = migraPara.AddFormattedText(txt);
                     ApplyTextFormatting(run, text); // Apply text-level formatting

                  }
               }
            }
         }

         // Render the PDF from the MigraDoc document
         PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
         pdfRenderer.Document = doc;
         pdfRenderer.RenderDocument();

         // Save the generated PDF
         pdfRenderer.PdfDocument.Save(_maskFileMerged!);
         // Return the file.
         return doc;
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
               migraPara.Format.SpaceBefore = Unit.FromPoint(props!.SpacingBetweenLines.Before! != null ? Convert.ToDouble(props!.SpacingBetweenLines.Before!) : 0);
               migraPara.Format.SpaceAfter = Unit.FromPoint(props!.SpacingBetweenLines.After! != null ? Convert.ToDouble(props!.SpacingBetweenLines.After!) : 0);
               migraPara.Format.LineSpacing = Unit.FromPoint(props!.SpacingBetweenLines.Line! != null ? Convert.ToDouble(props!.SpacingBetweenLines.Line!) : 0);
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

      /// <summary>
      /// Estrae tutti i placeholders presenti nel file di testo.
      /// </summary>
      /// <returns>Elenco dei placeholders presenti nei documenti.</returns>
      private List<Placeholder> ExtractPlaceholders() {
         var mergeFields = new List<Placeholder>();

         using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(_filename!, false)) {
            var body = wordDoc.MainDocumentPart!.Document.Body;
            var runs = body!.Descendants<Run>().ToList();

            for (int i = 0; i < runs.Count; i++) {
               var fldChar = runs[i].GetFirstChild<FieldChar>();
               if (fldChar != null && fldChar.FieldCharType! == FieldCharValues.Begin) {
                  string fieldCode = "";
                  bool foundEnd = false;
                  Placeholder? plc = null;

                  for (int j = i + 1; j < runs.Count; j++) {
                     var innerFldChar = runs[j].GetFirstChild<FieldChar>();
                     if (innerFldChar != null && innerFldChar.FieldCharType! == FieldCharValues.End) {
                        foundEnd = true;
                        break;
                     }

                     foreach (var element in runs[j].Elements<OpenXmlElement>()) {
                        if (element.LocalName == "instrText") {
                           fieldCode += element.InnerText;
                        }
                     }
                  }

                  plc = new Placeholder();
                  if (foundEnd && fieldCode.Contains("MERGEFIELD")) {
                     var parts = fieldCode.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                     int index = Array.IndexOf(parts, "MERGEFIELD");
                     if (index >= 0 && index + 1 < parts.Length) {
                        string ph = parts[index + 1];
                        if (!mergeFields.Any(p => p.Name == ph)) {
                           switch (parts.Count() - 1) {
                              case 1:
                                 plc.Name = parts[1];
                                 plc.RawDefinition = fieldCode.Replace(" MERGEFIELD ", "").Trim();
                                 break;
                              case 2:
                                 plc.Name = parts[1];
                                 plc.Type = parts[2];
                                 plc.RawDefinition = fieldCode.Replace(" MERGEFIELD ", "").Trim();
                                 break;
                              case 4:
                                 plc.Name = parts[1];
                                 plc.Type = parts[2];
                                 plc.Format = parts[4].Replace("\"", "");
                                 plc.RawDefinition = fieldCode.Replace(" MERGEFIELD ", "").Trim();
                                 break;
                           }
                           mergeFields.Add(plc);
                        }
                     }
                  }
               }
            }
         }

         return mergeFields;
      }
   }
}
