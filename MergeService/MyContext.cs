using net.nick4name.MergeService;
using System;
using System.Collections.Generic;

/// <summary>
/// Implementa l'interfaccia IMyContext per l'istanza generica T e rappresenta il contesto dati con cui effettuare 
/// il merge con il file di testo template.
/// </summary>
/// <typeparam name="T">Classe di tipo DBContext che rappresenta un'istanza di tabella o vista di db</typeparam>
public class MyContext<T> : IMyContext<T> {
   private readonly T _instance;

   /// <summary>
   /// Restituisce il contesto dati per l'istanza generica T.
   /// </summary>
   /// <param name="instance">Istanza di tipo DBContext che rappresenta un'istanza di tabella o vista di db.</param>
   public MyContext(T instance) {
      _instance = instance;
   }

   /// <summary>
   /// Restituisce l'istanza del tipo interno alla classe generica T associata al contesto dati.
   /// </summary>
   /// <returns>Istanza di tabella o vista di db nella forma DBContext.</returns>
   public T GetInstance() {
      return _instance;
   }
}
