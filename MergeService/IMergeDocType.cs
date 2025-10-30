namespace net.nick4name.MergeService {
   /// <summary>
   /// Interfaccia che astrae il tipo di documento template, txt, docx, ..., su cui eseguire il merge.
   /// </summary>
   internal interface IMergeDocType {
      /// <summary>
      /// Contenuto del documento template
      /// </summary>
      byte[] SourceContent { set; }

      /// <summary>
      /// Esegue il merge del documento template con i dati dell'istanza generica T.
      /// </summary>
      /// <typeparam name="T">Classe che definisce una tabella o vista di db.</typeparam>
      /// <param name="context">Istanza di tabella o vista di db.</param>
      /// <returns>byte[] del documento dopo il merge.</returns>
      byte[] ExecuteMerge<T>(T context);

      /// <summary>
      /// Nome file del documento template.
      /// </summary>
      string Filename { set; }

      /// <summary>
      /// Descrizione del nome che il file generato dopo il merge dovrà avere.
      /// </summary>
      string MaskFileMerged { set; }

      /// <summary>
      /// Nome del file generato dopo il merge.
      /// </summary>
      string FileMerged { get; }
      }
}
