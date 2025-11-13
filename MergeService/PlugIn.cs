namespace net.nick4name.MergeService {
   /// <summary>
   /// Classe che descrive i plug-in di merge presenti nel file di configurazione.
   /// </summary>
   public class PlugIn {
      /// <summary>
      /// Nome del plug-in.
      /// </summary>
      public string? Name { get; set; }

      /// <summary>
      /// Nome dell'eseguibile del plug-in.
      /// </summary>
      public string? Assembly { get; set; }

      /// <summary>
      /// Classe che istanzia il plug-in.
      /// </summary>
      public string? Class { get; set; }

      /// <summary>
      /// Mime type supportato dal plug-in.
      /// </summary>
      public string? Mime { get; set; }
      }
}
