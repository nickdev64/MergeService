namespace net.nick4name.MergeService.Exceptions {

   /// <summary>
   /// Eccezione sollevata per errori di configurazione del file .config dell'applicazione.
   /// </summary>
   public class ConfigErrorException : Exception {
      public ConfigErrorException(string msg)
          : base(msg) {

      }
   }
}
