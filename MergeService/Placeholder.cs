using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace net.nick4name.MergeService {
   internal class Placeholder {
      public string? Name { get; set; }
      public string? Type { get; set; } = null;
      public string? Format { get; set; } = null;

      public Placeholder() {
      }

      public Placeholder(string name) {
         Name = name;
      }

      public Placeholder(string name, string type, string format) {
         Name = name;
         Type = type;
         Format = format;
      }
   }
}
