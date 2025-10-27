using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace net.nick4name.MergeService {
   /// <summary>
   /// Interfaccia che definisce il contesto dati per l'istanza generica T.
   /// </summary>
   /// <typeparam name="T">Classe generica che rappresenta un'istanza di tabella o vista di db.</typeparam>
   public interface IMyContext<T> {
      T GetInstance();
   }
}
