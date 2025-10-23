namespace net.nick4name.MergeService {

   /// <summary>
   /// Interfaccia che definisce le funzionalità di merge del file di testo template con i dati dell'istanza generica T.
   /// </summary>
   public interface IMerge {
      /// <summary>
      /// Esegue il merge.
      /// </summary>
      /// <returns>Array di byte del documento dopo il merge.</returns>
      byte[] ExecuteMerge();

      /// <summary>
      /// Proprietà che imposta il path del file template di cui eseguire il merge con i dati dell'istanza generica T.
      /// </summary>
      string FileToMerge { set; }
   }
}
