using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class MergeFieldExtractor {
   public static List<string> GetMergeFields(string filePath) {
      var mergeFields = new List<string>();

      using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false)) {
         var body = wordDoc.MainDocumentPart!.Document.Body;
         var runs = body!.Descendants<Run>().ToList();

         for (int i = 0; i < runs.Count; i++) {
            var fldChar = runs[i].GetFirstChild<FieldChar>();
            if (fldChar != null && fldChar.FieldCharType! == FieldCharValues.Begin) {
               string fieldCode = "";
               bool foundEnd = false;

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

               if (foundEnd && fieldCode.Contains("MERGEFIELD")) {
                  var parts = fieldCode.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                  int index = Array.IndexOf(parts, "MERGEFIELD");
                  if (index >= 0 && index + 1 < parts.Length) {
                     mergeFields.Add(parts[index + 1]);
                  }
               }
            }
         }
      }

      return mergeFields;
   }
}