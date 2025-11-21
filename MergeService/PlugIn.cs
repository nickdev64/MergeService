namespace net.nick4name.MergeService {
   /// <summary>
   /// Classe che descrive i plug-in di merge presenti nel file di configurazione.
   /// </summary>
   public class PlugIn {
      private string? _mode;

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

      /// <summary>
      /// Modalità di utilizzo del servizio di merge: "sync" o "async".
      /// Defalt "sync".
      /// </summary>
      public string? Mode
      {
         get => _mode ?? "sync";
         set => _mode = value;
      }
   }
}
