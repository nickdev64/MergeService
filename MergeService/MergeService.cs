using DocumentFormat.OpenXml.ExtendedProperties;
using net.nick4name.MergeExtensions;
using net.nick4name.MergeService.Exceptions;
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

      private static string? _pluginsPath = "";
      private static Dictionary<string, PlugIn>? _plugIns;

      private const string CONFIG_VERSION = "1.0";

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
         if (!File.Exists(configPath)) {
            // file .config non trovato. Viene creato uno inizializzato
            initFileConfig(configPath);
         }

         _plugIns = loadPluginList(configPath);

         return _plugIns != null;
      }

      private static void initFileConfig(string configFile) {
         XElement root = new XElement("configuration",
               new XAttribute("version", CONFIG_VERSION),
            new XElement("pluginsSection",
               new XAttribute("path", ""),
                     new XElement("plug-in",
                        new XAttribute("name", ""),
                        new XAttribute("assembly", ""),
                        new XAttribute("class", ""),
                        new XAttribute("mime", "")
                      )
             )
          );

         XDocument doc = new XDocument(root);
         doc.Save(configFile);
      }

      /// <summary>
      /// Restituisce la lista dei plug-in di merge definiti nel file di configurazione.
      /// </summary>
      /// <param name="configPath">Path del file di configurazione.</param>
      /// <returns>Proprietà Mime è chiave della lista.</returns>
      private static Dictionary<string, PlugIn> loadPluginList(string configPath) {
         Dictionary<string, PlugIn> plugins = new Dictionary<string, PlugIn>();
         XDocument xConfig = XDocument.Load(configPath);
         
         string ver = xConfig.Root!.Attribute("version")!.Value ??
            throw new ConfigErrorException($"Wrong configuration for {configPath}. Missing 'version' attribute for 'configuration' element.");
         if (!ver.Equals(CONFIG_VERSION)) {
            throw new ConfigWrongVersionException(configPath, ver, CONFIG_VERSION);
         }

         var pluginSect = xConfig.Root!.Elements().Where(x => x.Name.LocalName == "pluginsSection").FirstOrDefault();
         if (pluginSect != null) {
            _pluginsPath = pluginSect.Attribute("path")?.Value ?? null;
         } else {
            throw new ConfigErrorException($"Wrong configuration for {configPath}. Missing 'path' attribute for 'pluginsSection' element.");
         }

         if (_pluginsPath != null) {
            if (!Path.Exists(_pluginsPath)) {
               Directory.CreateDirectory(_pluginsPath);
            }
         } else {
            throw new Exception("Plugins path not defined in configuration file.");
         }

         var pluginElements = xConfig.Descendants("plug-in");
         foreach (var element in pluginElements) {
            PlugIn plugin = new PlugIn {
               Name = element.Attribute("name")?.Value,
               Assembly = Path.Combine(_pluginsPath, element.Attribute("assembly")?.Value!),
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
