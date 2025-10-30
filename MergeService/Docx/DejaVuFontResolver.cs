using PdfSharp.Fonts;
using System.Reflection;

namespace net.nick4name.MergeService.Docx {
   public class DejaVuFontResolver : IFontResolver {
      public string DefaultFontName => "DejaVuSans";

      public byte[] GetFont(string faceName) {
         return faceName switch {
            "DejaVuSans" => LoadFont("net.nick4name.MergeService.Fonts.DejaVuSans.ttf"),
            "DejaVuSans#b" => LoadFont("net.nick4name.MergeService.Fonts.DejaVuSans-Bold.ttf"),
            "DejaVuSans#i" => LoadFont("net.nick4name.MergeService.Fonts.DejaVuSans-Oblique.ttf"),
            "DejaVuSans#bi" => LoadFont("net.nick4name.MergeService.Fonts.DejaVuSans-BoldOblique.ttf"),
            "Courier New" => LoadFont("net.nick4name.MergeService.Fonts.DejaVuSansMono.ttf"), // fallback per Courier New
            _ => throw new InvalidOperationException($"Font '{faceName}' non gestito.")
         };
      }

      public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic) {
         if (familyName.Equals("DejaVuSans", StringComparison.OrdinalIgnoreCase)) {
            var suffix = (isBold, isItalic) switch {
               (true, true) => "#bi",
               (true, false) => "#b",
               (false, true) => "#i",
               _ => ""
            };
            return new FontResolverInfo("DejaVuSans" + suffix);
         }

         return new FontResolverInfo("DejaVuSans");
      }

      private byte[] LoadFont(string resourceName) {
         using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
         if (stream == null)
            throw new FileNotFoundException($"Font embedded '{resourceName}' non trovato.");
         using var ms = new MemoryStream();
         stream.CopyTo(ms);
         return ms.ToArray();
      }
   }
}
