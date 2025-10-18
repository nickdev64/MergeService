using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace net.nick4name.MergeService {

   public class MergeService<T> where T : class {

      private IMyContext<T>? _ctx;
      private IMerge _merge;

   }
}
