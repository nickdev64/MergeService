using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Fonts;
using MigraListType = MigraDoc.DocumentObjectModel.ListType;
using MigraListInfo = MigraDoc.DocumentObjectModel.ListInfo;
using MigraDocDoc = MigraDoc.DocumentObjectModel.Document;
using MigraDocParagraph = MigraDoc.DocumentObjectModel.Paragraph;
using OpenXmlParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using OpenXmlColor = DocumentFormat.OpenXml.Wordprocessing.Color;
using OpenXmlUnderline = DocumentFormat.OpenXml.Wordprocessing.Underline;

namespace net.nick4name.MergeService.Docx {
   internal class MergeDocx<T> : IMergeDocType, IMerge where T : class {
      string? _fileToMerge = null;
      string? _fileMerged = null;

      /// <summary>
      /// byte[] che rappresenta il contenuto del documento di testo template.
      /// Proprietà non supportata. Referenziare il documento template tramite la proprietà FileToMerge.
      /// </summary>
      /// <exception cref="NotImplementedException">Proprietà non supportata.</exception>
      public byte[] SourceContent { set => throw new NotImplementedException(); }

      /// <summary>
      /// Nome file del documento template.
      /// </summary>
      public string FileToMerge { set => _fileToMerge = value; }

      /// <summary>
      /// Imposta il nome del file generato dopo il merge.
      /// </summary>
      public string FileMerged { set => _fileMerged = value; }

      /// <summary>
      /// Esegue il merge di FileToMerge con i dati dell'istanza generica T.
      /// Scriverà il file risultante in MaskFileMerged.
      /// </summary>
      /// <typeparam name="T">Classe generica che rappresenta un'istanza di tabella o vista di db.</typeparam>
      /// <param name="context">Istanza di tabella o vista di db.</param>
      /// <returns>byte[] del file prodotto in MaskFileMerged.</returns>
      public byte[] ExecuteMerge<T>(T context) {
         if (string.IsNullOrEmpty(_fileToMerge)) {
            throw new InvalidOperationException("FileToMerge non impostato per l'operazione di merge.");
         }

         if (string.IsNullOrEmpty(_fileMerged)) {
            throw new InvalidOperationException("FileMerged non impostato per l'operazione di merge.");
         }

         // lista dei placeholders del documento
         List<Placeholder> placeholders = ExtractPlaceholders();

         // esegue il merge
         MigraDocDoc doc = ConvertWordToPdf(placeholders, context);

         return File.ReadAllBytes(_fileMerged!);
      }

      /// <summary>
      /// Metodo non supportato.
      /// Utlizzare ExecuteMerge<T>(T context).
      /// </summary>
      /// <returns></returns>
      /// <exception cref="NotImplementedException"></exception>
      public byte[] ExecuteMerge() {
         throw new NotImplementedException();
      }

      /// <summary>
      /// Estrae tutti i placeholders presenti nel file di testo.
      /// </summary>
      /// <returns>Elenco dei placeholders presenti nei documenti.</returns>
      private List<Placeholder> ExtractPlaceholders() {
         var mergeFields = new List<Placeholder>();

         using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(_fileToMerge!, false)) {
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

      #region Convert Word To PDF

      /// <summary>
      /// Converte il documento Word FileToMerge in PDF sostituendone i placeholders con i valori dell'istanza generica T.
      /// </summary>
      /// <typeparam name="T"></typeparam>
      /// <param name="placeholders"></param>
      /// <param name="context"></param>
      /// <returns></returns>
      private MigraDocDoc ConvertWordToPdf<T>(List<Placeholder> placeholders, T context) {
         GlobalFontSettings.FontResolver = new DejaVuFontResolver();

         var doc = new MigraDocDoc();
         var section = doc.AddSection();
         section.PageSetup.PageFormat = PageFormat.A4; //A4 page
         var style = doc.Styles["Normal"]!;
         style.Font.Name = "DejaVuSans";
         style.Font.Bold = true;

         // Open the Word document in read-only mode
         using (var wordDoc = WordprocessingDocument.Open(_fileToMerge!, false)) {
            var body = wordDoc.MainDocumentPart!.Document.Body;
            ApplyPageSetup(wordDoc, section); // Apply page margins from Word to MigraDoc

            foreach (var para in body!.Elements<OpenXmlParagraph>()) {
               // Prova a convertire come elenco, altrimenti crea paragrafo normale
               var migraPara = TryConvertListParagraph(wordDoc, para, section);

               ApplyParagraphFormatting(para, migraPara!, section);

               var runs = para.Elements<Run>().ToList();
               int i = 0;

               while (i < runs.Count) {
                  var run = runs[i];
                  var fldChar = run.GetFirstChild<FieldChar>();

                  if (fldChar != null && fldChar?.FieldCharType! == FieldCharValues.Begin) {
                     int startIndex = i;
                     int endIndex = -1;
                     string? fieldName = null;

                     // Cerca il nome del campo e la fine
                     for (int j = i + 1; j < runs.Count; j++) {
                        var r = runs[j];

                        var fieldCode = r.GetFirstChild<FieldCode>();
                        if (fieldCode != null && fieldCode.Text.Contains("MERGEFIELD")) {
                           var tokens = fieldCode.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                           if (tokens.Length >= 2)
                              fieldName = tokens[1];
                        }

                        var endChar = r.GetFirstChild<FieldChar>();
                        if (endChar != null && endChar?.FieldCharType! == FieldCharValues.End) {
                           endIndex = j;
                           break;
                        }
                     }

                     if (fieldName != null && endIndex > startIndex) {
                        var prop = typeof(T).GetProperty(fieldName);
                        string value = prop?.GetValue(context)?.ToString() ?? "";

                        // Rimuove tutti i run tra begin e end inclusi
                        for (int j = startIndex; j <= endIndex; j++) {
                           para.RemoveChild(runs[j]);
                        }

                        // Inserisce il valore sostitutivo
                        var newRun = new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(value));
                        para.InsertAt(newRun, startIndex);

                        var text = migraPara!.AddFormattedText(value);
                        ApplyTextFormatting(run, text, para);

                        // Ricostruisci la lista dei run dopo la modifica
                        runs = para.Elements<Run>().ToList();
                        i = startIndex + 1;
                        continue;
                     }
                  }

                  // Run normale
                  if (!run.InnerText.Contains("MERGEFIELD")) {
                     var text = migraPara!.AddFormattedText(run.InnerText);
                     ApplyTextFormatting(run, text, para);
                  }

                  i++;
               }
            }
         }

         // Render the PDF from the MigraDoc document
         PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
         pdfRenderer.Document = doc;
         pdfRenderer.RenderDocument();

         // Save the generated PDF
         pdfRenderer.PdfDocument.Save(_fileMerged!);
         // Return the file.
         return doc;
      }

      /// <summary>
      /// Converte i paragrafi Word che sono elenchi in paragrafi MigraDoc con stile elenco.
      /// </summary>
      /// <param name="wordDoc"></param>
      /// <param name="para"></param>
      /// <param name="section"></param>
      /// <param name="migraPara"></param>
      /// <returns>True se la conversione è avvenuta altrimenti false.</returns>
      bool TryConvertListParagraph_(
          WordprocessingDocument wordDoc,
          OpenXmlParagraph para,
          Section section,
          out MigraDocParagraph? migraPara) {

         migraPara = null;

         var numberingProps = para.GetFirstChild<NumberingProperties>();
         if (numberingProps?.NumberingId?.Val == null)
            return false;

         int numId = numberingProps.NumberingId.Val.Value;
         int ilvl = numberingProps.NumberingLevelReference?.Val?.Value ?? 0;

         var numberingPart = wordDoc.MainDocumentPart?.NumberingDefinitionsPart;
         if (numberingPart == null)
            return false;

         var numberingInstance = numberingPart.Numbering.Elements<NumberingInstance>()
             .FirstOrDefault(n => n.NumberID?.Value == numId);
         if (numberingInstance?.AbstractNumId?.Val == null)
            return false;

         int abstractNumId = numberingInstance.AbstractNumId.Val.Value;

         var abstractNum = numberingPart.Numbering.Elements<AbstractNum>()
             .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
         if (abstractNum == null)
            return false;

         var level = abstractNum.Elements<Level>()
             .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);
         if (level?.NumberingFormat?.Val == null)
            return false;

         var format = level.NumberingFormat.Val.Value;

         // Crea il paragrafo MigraDoc con stile elenco
         migraPara = section.AddParagraph();
         migraPara.Style = "List";

         MigraListType listType = format == NumberFormatValues.Bullet
             ? ilvl switch {
                0 => MigraListType.BulletList1,
                1 => MigraListType.BulletList2,
                _ => MigraListType.BulletList3
             }
             : ilvl switch {
                0 => MigraListType.NumberList1,
                1 => MigraListType.NumberList2,
                _ => MigraListType.NumberList3
             };

         migraPara.Format.ListInfo = new MigraListInfo {
            ListType = listType
         };

         return true;
      }

      private MigraDocParagraph TryConvertListParagraph(WordprocessingDocument wordDoc, OpenXmlParagraph para, Section section) {
         var numberingProps = para.ParagraphProperties?.NumberingProperties;
         //if (numberingProps?.NumberingId?.Val == null)
         //   return null!;
         
         if (numberingProps?.NumberingId?.Val != null) {
            int numId = numberingProps.NumberingId.Val.Value;
            int ilvl = numberingProps.NumberingLevelReference?.Val?.Value ?? 0;

            var numberingPart = wordDoc.MainDocumentPart?.NumberingDefinitionsPart;
            var numberingInstance = numberingPart?.Numbering.Elements<NumberingInstance>()
                .FirstOrDefault(n => n.NumberID?.Value == numId);
            var abstractNumId = numberingInstance?.AbstractNumId?.Val?.Value;

            var abstractNum = numberingPart?.Numbering.Elements<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
            var level = abstractNum?.Elements<Level>()
                .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);
            var format = level?.NumberingFormat?.Val?.Value;

            if (format != null) {
               var migraPara = section.AddParagraph();
               migraPara.Style = "List";
               migraPara.Format.ListInfo = new ListInfo {
                  ListType = format == NumberFormatValues.Bullet
                       ? ilvl switch {
                          0 => ListType.BulletList1,
                          1 => ListType.BulletList2,
                          _ => ListType.BulletList3
                       }
                       : ilvl switch {
                          0 => ListType.NumberList1,
                          1 => ListType.NumberList2,
                          _ => ListType.NumberList3
                       }
               };
               return migraPara;
            }
         }

         return section.AddParagraph();
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

      static void ApplyParagraphFormatting(OpenXmlParagraph para, MigraDocParagraph migraPara, Section section) {
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

      /// <summary>
      /// Converte la formattazione del testo da Word a MigraDoc.
      /// Considera sia la formattazione del run che quella del paragrafo.
      /// Gestisce grassetto, corsivo, sottolineato, dimensione, colore e font.
      /// </summary>
      /// <param name="run">Unità di base di testo nell'XML in OpenXML <w:r>.</param>
      /// <param name="text">Frammento di testo in MigraDoc con formattazione applicata all’interno di un paragrafo.</param>
      /// <param name="para">Unità strutturale base di un documento Word in OpenXML: ogni paragrafo può contenere testo, 
      /// formattazione, elenchi, stili, interruzioni e altro.<w:p></param>
      static void ApplyTextFormatting(Run run, FormattedText text, OpenXmlParagraph para) {
         var runProps = run.RunProperties;
         var paraProps = para.ParagraphProperties?.ParagraphMarkRunProperties;

         bool IsBold(OpenXmlElement? props) {
            var b = props?.GetFirstChild<Bold>();
            var bcs = props?.GetFirstChild<BoldComplexScript>();
            return (b != null && (b.Val == null || b.Val.Value != false)) ||
                   (bcs != null && (bcs.Val == null || bcs.Val.Value != false));
         }

         bool IsItalic(OpenXmlElement? props) {
            var i = props?.GetFirstChild<Italic>();
            var ics = props?.GetFirstChild<ItalicComplexScript>();
            return (i != null && (i.Val == null || i.Val.Value != false)) ||
                   (ics != null && (ics.Val == null || ics.Val.Value != false));
         }

         bool IsUnderline(OpenXmlElement? props) {
            var u = props?.GetFirstChild<OpenXmlUnderline>();
            return u != null && (u.Val == null || u.Val.Value != UnderlineValues.None);
         }

         text.Bold = IsBold(runProps) || IsBold(paraProps);
         text.Italic = IsItalic(runProps) || IsItalic(paraProps);
         text.Underline = IsUnderline(runProps) || IsUnderline(paraProps)
             ? MigraDoc.DocumentObjectModel.Underline.Single
             : MigraDoc.DocumentObjectModel.Underline.None;

         var fontSize = runProps?.FontSize?.Val ?? paraProps?.GetFirstChild<FontSize>()?.Val;
         text.Size = fontSize != null ? Convert.ToDouble(fontSize) / 2.0 : 12;

         var color = runProps?.Color?.Val ?? paraProps?.GetFirstChild<OpenXmlColor>()?.Val;
         if (!string.IsNullOrEmpty(color))
            text.Color = ConvertHexToColor(color!);

         var fontName = runProps?.RunFonts?.Ascii?.Value ?? paraProps?.GetFirstChild<RunFonts>()?.Ascii?.Value;
         text.Font.Name = !string.IsNullOrEmpty(fontName) ? fontName : "Times New Roman";
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
      #endregion Convert Word To PDF
   }
}
