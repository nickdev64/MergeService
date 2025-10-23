using Microsoft.AspNetCore.StaticFiles;
using System.Text;
using System.Text.RegularExpressions;

namespace net.nick4name.MergeService {

   /// <summary>
   /// Rappresenta l'implementazione concreta dell'interfaccia IMerge per l'istanza generica T.
   /// </summary>
   /// <typeparam name="T">Classe di tipo DBContext che rappresenta un'istanza di tabella o vista di db</typeparam>
   /// <remarks>Gestisce il merge del file di testo template con i dati dell'istanza generica T.</remarks>
   public class Merge<T> : IMerge where T : class {
      private byte[]? _file;
      private string? _filetext;
      private readonly IMyContext<T>? _context = null;

      /// <summary>
      /// Rappresenta il contesto dati per l'istanza generica T con cui realizzare il merge con il file template.
      /// </summary>
      public IMyContext<T>? SrcContext { get; } = null;

      /// <summary>
      /// Restituisce l'istanza di merge per l'istanza generica T.
      /// </summary>
      /// <param name="srcContext">Contesto dati per l'istanza generica T con cui realizzare il merge con il file template.</param>
      /// <remarks>
      /// Richiede che l'istanza di contesto dati sia valorizzata.
      /// Utilizzata per passare i dati con cui realizzare il merge.
      /// </remarks>
      public Merge(IMyContext<T> srcContext) {
         SrcContext = srcContext;
      }

      /// <summary>
      /// Filename del file di testo template di cui eseguire il merge con i dati dell'istanza generica T.
      /// MIME types supportati:
      /// - text/plain
      /// - application/vnd.openxmlformats-officedocument.wordprocessingml.document
      /// </summary>
      /// <exception cref="InvalidOperationException">Il MIME type del file non è supportato.</exception>
      public string FileToMerge {
         set {
            string filePath = Path.GetFullPath(value);
            string mime = getContentType(filePath);
            switch (mime) {
               case "text/plain":
                  using (StreamReader reader = new StreamReader(filePath, Encoding.UTF8)) {
                     _filetext = reader.ReadToEnd();
                  }
                  break;
               case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                  break;
               default:
                  throw new InvalidOperationException("File type not supported for merge: " + mime);
            }
         }
      }

      /// <summary>
      /// La funzione esegue il merge del file di testo d'istanza FileToMerge con i dati dell'istanza generica T.
      /// La logica di funzionamento è la seguente:
      /// 1) Estrae tutti i placeholder presenti nel file di testo, identificati dalla sintassi {NomeProprietà}.
      ///    I placeholders corrispondono alle colonne dell'istanza generica T.
      /// 2) Per ogni placeholder trovato, recupera il valore della proprietà corrispondente dall'istanza T.
      /// 3) Sostituisce ogni placeholder nel file di testo con il valore ottenuto dalla proprietà.
      /// </summary>
      /// <returns></returns>
      /// <exception cref="InvalidOperationException"></exception>
      public byte[] ExecuteMerge() {
         if (_filetext == null || _filetext.Length == 0)
            throw new InvalidOperationException("File to merge is not set or empty.");

         // lista dei placeholders del documento
         List<string> placeholders = ExtractPlaceholders();

         // accesso alla istanza generica T relativa alla fonte dati
         T inst = SrcContext?.GetInstance()!;

         // esegue il merge
         var properties = typeof(T).GetProperties();
         foreach (string ph in placeholders) {
            // costruisce il placeholder secondo la sintassi del documento
            string placeholder = "{" + ph + "}";

            // ottiene il valore corrispondente alla colonna avente per nome il placeholder
            var prop = typeof(T).GetProperty(ph);
            string value = prop?.GetValue(inst)?.ToString() ?? "";

            // sostituisce il placeholder nel documento con il valore ottenuto
            _filetext = _filetext!.Replace(placeholder, value);
         }

         _file = Encoding.UTF8.GetBytes(_filetext);
         return _file;
      }

      /// <summary>
      /// Estrae tutti i placeholders presenti nel file di testo.
      /// </summary>
      /// <returns>Elenco dei placeholders presenti nei documenti.</returns>
      private List<string> ExtractPlaceholders() {
         var matches = Regex.Matches(_filetext, @"\{([^{}]+)\}");
         var result = new List<string>();

         foreach (Match match in matches) {
            result.Add(match.Groups[1].Value);
         }

         return result;
      }

      /// <summary>
      /// Restituisce il MIME type di filePath in base alla sua estensione.
      /// </summary>
      /// <param name="filePath">Nome file inclusivo di path, se diverso dalla directory corrente, ed estensione.</param>
      /// <returns>MIME type. "application/octet-stream" se non determinato.</returns>
      /// <remarks>Created: 30/07/2025</remarks>
      private static string getContentType(string filePath) {
         var provider = new FileExtensionContentTypeProvider();
         if (!provider.TryGetContentType(filePath, out string contentType)) {
            contentType = "application/octet-stream"; // default generico
         }
         return contentType;
      }
   }
}
