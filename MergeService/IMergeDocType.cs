namespace net.nick4name.MergeService {
   /// <summary>
   /// Interfaccia che astrae il tipo di documento template, txt, docx, ..., su cui eseguire il merge.
   /// </summary>
   internal interface IMergeDocType : IMerge {
      /// <summary>
      /// Passa il contenuto del documento template come byte[].
      /// </summary>
      byte[] SourceContent { set; }

      /// <summary>
      /// Esegue il merge del documento template con i dati dell'istanza generica T.
      /// </summary>
      /// <typeparam name="T">Classe che definisce una tabella o vista di db.</typeparam>
      /// <param name="context">Istanza di tabella o vista di db.</param>
      /// <returns>byte[] del documento dopo il merge.</returns>
      byte[] ExecuteMerge<T>(T context);
   }
}
