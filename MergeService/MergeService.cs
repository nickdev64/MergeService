using DocumentFormat.OpenXml.ExtendedProperties;
using net.nick4name.MergeExtensions;
using System.Reflection;
using System.Xml.Linq;

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

      private static Dictionary<string, PlugIn>? _plugIns;

      /// <summary>
      /// Restituisce il servizio di merge per l'istanza generica T.
      /// </summary>
      /// <param name="ctx">Istanza che rappresenta un'istanza di tabella o vista di db.</param>
      public MergeService(IMyContext<T> ctx) {
         _ctx = ctx;
         _merge = new Merge<T>(ctx);

         if (!initPlugins()) {
            throw new InvalidOperationException("No plug-ins initialized for merge service.");
         }
         ((Merge<T>)_merge).PlugIns = _plugIns;
      }

      private static bool initPlugins() {
         string dllPath = Assembly.GetExecutingAssembly().Location;
         string nomeDll = Path.GetFileName(dllPath);
         string configPath = $"{nomeDll}.config";

         _plugIns = loadPluginList(configPath);

         return _plugIns != null;
      }

      /// <summary>
      /// Restituisce la lista dei plug-in di merge definiti nel file di configurazione.
      /// </summary>
      /// <param name="configPath">Path del file di configurazione.</param>
      /// <returns>Proprietà Mime è chiave della lista.</returns>
      private static Dictionary<string, PlugIn> loadPluginList(string configPath) {
         Dictionary<string, PlugIn> plugins = new Dictionary<string, PlugIn>();
         XDocument xConfig = XDocument.Load(configPath);
         var pluginElements = xConfig.Descendants("plug-in");
         foreach (var element in pluginElements) {
            PlugIn plugin = new PlugIn {
               Name = element.Attribute("name")?.Value,
               Assembly = element.Attribute("assembly")?.Value,
               Class = element.Attribute("class")?.Value,
               Mime = element.Attribute("mime")?.Value
            };
            if (plugin.Mime != null) {
               plugins[plugin.Mime] = plugin;
            }
         }

         return plugins;
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
         return _merge.ExecuteMerge(); //<T>(_ctx!.GetInstance());
      }

      /// <summary>
      /// Restituisce l'istanza del tipo interno alla classe generica T associata al contesto dati.
      /// </summary>
      /// <returns>Istanza di tabella o vista di db.</returns>
      public T GetInstance() {
         return _ctx!.GetInstance();
      }

      /// <summary>
      /// Imposta il nome del file generato dopo il merge.
      /// </summary>
      public string FileMerged { set => _merge.FileMerged = value; }

   }
}
