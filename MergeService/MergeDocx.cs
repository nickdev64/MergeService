using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace net.nick4name.MergeService {
   internal class MergeDocx<T> : IMergeDocType where T : class {
      string? _filename = null;

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

      public byte[] ExecuteMerge<T>(T context) {
         // lista dei placeholders del documento
         List<Placeholder> placeholders = ExtractPlaceholders();

         // esegue il merge
         var properties = typeof(T).GetProperties();
         foreach (Placeholder ph in placeholders) {
            // costruisce il placeholder secondo la sintassi del documento
            string placeholder = "{" + ph + "}";

            // ottiene il valore corrispondente alla colonna avente per nome il placeholder
            var prop = typeof(T).GetProperty(ph);
            string value = prop?.GetValue(context)?.ToString() ?? "";

            // sostituisce il placeholder nel documento con il valore ottenuto
            _contTyped = _contTyped!.Replace(placeholder, value);
         }

         return null!;
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
                           //plc = new Placeholder(fieldCode.Trim());
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
                                 break;
                              case 2:
                                 plc.Name = parts[1];
                                 plc.Type = parts[2];
                                 break;
                              case 4:
                                 plc.Name = parts[1];
                                 plc.Type = parts[2];
                                 plc.Format = parts[4].Replace("\"", "");
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
