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
      /// <typeparam name="T">Classe di tipo DBContext che definisce una tabella o vista di db.</typeparam>
      /// <param name="context">Istanza di tabella o vista di db nella forma DBContext.</param>
      /// <returns>byte[] del documento dopo il merge.</returns>
      byte[] ExecuteMerge<T>(T context);
   }
}
