namespace net.nick4name.MergeService {
   /// <summary>
   /// Definisce un'istanza di placeholder presente nel documento template.
   /// </summary>
   internal class Placeholder {
      /// <summary>
      /// [Obbligatorio] Nome del placeholder, es. il nome della colonna di tabella.
      /// </summary>
      public string? Name { get; set; }

      /// <summary>
      /// [Facoltativo] Tipo del placeholder, es. DATE, ...
      /// </summary>
      public string? Type { get; set; } = null;

      /// <summary>
      /// [Facoltativo] Formato della rappresentazione del placeholder, es. dd/MM/yyyy per le date.
      /// </summary>
      public string? Format { get; set; } = null;

      /// <summary>
      /// [Facoltativo] Descrizione testuale del placeholder, es. 'DATE \@ "dd/MM/yyyy"«DataRilDoc»'.
      /// </summary>
      /// <remarks>
      /// Può essere utile laddove nel documento template fosse rappresentato in forma estesa e fosse 
      /// così più pratico sostituirlo con il valore di merge. Ad esempio per i campi data in Word. 
      /// Es.: "Nato il DATE \@"dd/MM/yyyy"«DataNascita» a..."
      /// sostituire 'DATE \@"dd/MM/yyyy"«DataNascita»' con '15/08/2023' 
      /// ottenendo quindi "Nato il 15/08/2023 a..."
      /// </remarks>
      public string? RawDefinition { get; set; } = null;

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
