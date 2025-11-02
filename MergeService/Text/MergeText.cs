using System.Text;
using System.Text.RegularExpressions;

namespace net.nick4name.MergeService.Text {
   /// <summary>
   /// Classe che implementa il merge di un template di file di testo UTF8 in formato byte[], SourceContent, 
   /// con il contesto dati generico di tipo T.
   /// </summary>
   /// <typeparam name="T">Istanza di tabella o vista di db.</typeparam>
   internal class MergeText<T> : IMergeDocType where T : class {
      byte[]? _content;
      string _contTyped = "";
      string? _filename = null;

      /// <summary>
      /// byte[] che rappresenta il contenuto del documento di testo template in formato UTF8.
      /// </summary>
      public byte[] SourceContent {
         set {
            _content = value;
            _contTyped = Encoding.UTF8.GetString(_content);
         }
      }

      /// <summary>
      /// Nome file del documento template.
      /// </summary>
      public string Filename {
         set {
            _filename = value;
         }
      }

      /// <summary>
      /// Proprietà non supportata.
      /// </summary>
      public string MaskFileMerged { set => throw new NotImplementedException(); }

      /// <summary>
      /// Proprietà non supportata.
      /// </summary>
      string IMergeDocType.FileMerged { set => throw new NotImplementedException(); }

      /// <summary>
      /// Esegue il merge del documento di testo template con i dati dell'istanza generica T
      /// i cui placeholders sono nel formato '{column_name}'.
      /// </summary>
      /// <typeparam name="T">Classe generica che definisce una tabella o vista di db.</typeparam>
      /// <param name="context">Istanza di tabella o vista di db.</param>
      /// <returns>byte[] del documento di testo UTF8 dopo il merge.</returns>
      /// <remarks>
      /// I placeholder nel documento devono essere nel formato {column_name} dove column_name è nome 
      /// della proprietà della classe T, ovvero della colonna della tabella o vista che T descrive.
      /// </remarks>
      /// <example>
      /// Esempio di template:
      /// Egregio {FirstName} {LastName},
      /// </example>
      public byte[] ExecuteMerge<T>(T context) {
         if (!string.IsNullOrEmpty(_filename)) {
            _content = File.ReadAllBytes(_filename);
         }

         // lista dei placeholders del documento
         List<string> placeholders = ExtractPlaceholders();

         // esegue il merge
         foreach (string ph in placeholders) {
            // costruisce il placeholder secondo la sintassi del documento
            string placeholder = "{" + ph + "}";

            // ottiene il valore corrispondente alla colonna avente per nome il placeholder
            var prop = typeof(T).GetProperty(ph);
            string value = prop?.GetValue(context)?.ToString() ?? "";

            // sostituisce il placeholder nel documento con il valore ottenuto
            _contTyped = _contTyped!.Replace(placeholder, value);
         }

         return Encoding.UTF8.GetBytes(_contTyped);
      }

      /// <summary>
      /// Estrae tutti i placeholders presenti nel file di testo.
      /// </summary>
      /// <returns>Elenco dei placeholders presenti nei documenti.</returns>
      private List<string> ExtractPlaceholders() {
         var matches = Regex.Matches(_contTyped, @"\{([^{}]+)\}");
         var result = new List<string>();

         foreach (Match match in matches) {
            result.Add(match.Groups[1].Value);
         }

         return result;
      }
   }
}
