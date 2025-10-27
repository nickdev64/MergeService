namespace net.nick4name.MergeService {

   /// <summary>
   /// Classe entry point del servizio per eseguire il merge di un file di testo template con i dati 
   /// rappresentati dalla classe generica T.
   /// </summary>
   /// <typeparam name="T">Classe che definisce una tabella o vista di db.</typeparam>
   public class MergeService<T> where T : class {

      // *** ATTENZIONE!!! Non rinominare _ctx. Vedi Stampe.xaml.cs -> GetGenericInstance
      private IMyContext<T>? _ctx;
      // ***
      private IMerge _merge;

      /// <summary>
      /// Restituisce il servizio di merge per l'istanza generica T.
      /// </summary>
      /// <param name="ctx">Istanza che rappresenta un'istanza di tabella o vista di db.</param>
      public MergeService(IMyContext<T> ctx) {
         _ctx = ctx;
         _merge = new Merge<T>(ctx);
         //...
      }

      /// <summary>
      /// Path del file template di cui eseguire il merge.
      /// </summary>
      public string FileToMerge {
         set { _merge.FileToMerge = value; }
      }

      /// <summary>
      /// Esegue il merge del file di testo template con i dati dell'istanza generica T.
      /// </summary>
      /// <returns>Array di byte del documento dopo il merge.</returns>
      /// <remarks>Richiede che la proprietà FileToMerge sia stata valorizzata.</remarks>
      public byte[] ExecuteMerge() {
         return _merge.ExecuteMerge();
      }

      /// <summary>
      /// Restituisce l'istanza del tipo interno alla classe generica T associata al contesto dati.
      /// </summary>
      /// <returns>Istanza di tabella o vista di db.</returns>
      public T GetInstance() {
         return _ctx!.GetInstance();
      }

   }
}
