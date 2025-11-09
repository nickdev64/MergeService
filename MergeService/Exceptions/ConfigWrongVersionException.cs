using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace net.nick4name.MergeService.Exceptions {
   /// <summary>
   /// Eccezione sollevata quando la versione del file di configurazione non è quella attesa.
   /// </summary>
   public class ConfigWrongVersionException : Exception {
      public ConfigWrongVersionException(string configFilename,string currentVer, string expectedVer)
          : base($"Wrong version for '{configFilename}'. Current version is '{currentVer}', expected version is '{expectedVer}'.") {

      }
   }
}
