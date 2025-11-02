using Microsoft.AspNetCore.StaticFiles;
using net.nick4name.MergeService.Docx;
using net.nick4name.MergeService.Text;
using System.Text;

namespace net.nick4name.MergeService {

   /// <summary>
   /// Rappresenta l'implementazione concreta dell'interfaccia IMerge per l'istanza generica T.
   /// </summary>
   /// <typeparam name="T">Classe generica che rappresenta un'istanza di tabella o vista di db</typeparam>
   /// <remarks>Gestisce il merge del file di testo template con i dati dell'istanza generica T.</remarks>
   public class Merge<T> : IMerge where T : class {
      private string? _filetext;
      private string? _filePath;
      private string? _mime = null;
      private string? _fileMerged = null;

      /// <summary>
      /// Rappresenta il contesto dati per l'istanza generica T con cui realizzare il merge con il file template.
      /// </summary>
      public IMyContext<T>? SrcContext { get; } = null;

      /// <summary>
      /// Restituisce l'istanza di merge per l'istanza generica T.
      /// </summary>
      /// <param name="srcContext">
      /// Contesto dati per l'istanza generica T con cui realizzare il merge con il file template.
      /// </param>
      /// <remarks>
      /// Richiede che l'istanza di contesto dati sia valorizzata.
      /// Utilizzata per passare i dati con cui realizzare il merge.
      /// </remarks>
      public Merge(IMyContext<T> srcContext) {
         SrcContext = srcContext;
      }

      /// <summary>
      /// FileToMerge del file di testo template di cui eseguire il merge con i dati dell'istanza generica T.
      /// MIME types supportati:
      /// - text/plain
      /// - application/vnd.openxmlformats-officedocument.wordprocessingml.document
      /// </summary>
      /// <exception cref="InvalidOperationException">Il MIME type del file non è supportato.</exception>
      public string FileToMerge {
         set {
            string filePath = Path.GetFullPath(value);
            _mime = getContentType(filePath);
            switch (_mime) {
               case "text/plain":
                  using (StreamReader reader = new StreamReader(filePath, Encoding.UTF8)) {
                     _filetext = reader.ReadToEnd();
                  }
                  break;
               case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                  _filePath = filePath;
                  break;
               default:
                  throw new InvalidOperationException("File type not supported for merge: " + _mime);
            }
         }
      }

      /// <summary>
      /// Imposta il nome del file generato dopo il merge.
      /// </summary>
      public string FileMerged { set => _fileMerged = value; }

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
         //public byte[] ExecuteMerge<T1>(T1 context) {
         byte[]? file = null;

         switch (_mime) {
            case "text/plain":
               if (_filetext == null || _filetext.Length == 0)
                  throw new InvalidOperationException("File template non impostato oppure vuoto.");

               IMergeDocType mergeTxt = new MergeText<T>();
               mergeTxt.SourceContent = Encoding.UTF8.GetBytes(_filetext);
               byte[] mergedText = mergeTxt.ExecuteMerge<T>(SrcContext!.GetInstance());

               file = mergedText;
               break;
            case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
               if (string.IsNullOrEmpty(_filePath))
                  throw new InvalidOperationException("File template non impostato.");

               IMergeDocType mergeDocx = new MergeDocx<T>();
               mergeDocx.FileToMerge = _filePath;
               mergeDocx.FileMerged = _fileMerged!;
               mergeDocx.ExecuteMerge<T>(SrcContext!.GetInstance());
               break;
            default:
               throw new InvalidOperationException("File type not supported for merge: " + _mime);
         }

         return file!;
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
