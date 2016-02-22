using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Infodinamica.Framework.Exportable.Engines
{
    public interface IImportEngine
    {
        void SetDocument(MemoryStream file);
        void SetDocument(string path);
        void AddContainer<T>(string key) where T : class;
        IList<T> GetList<T>(string key) where T : class;
        IDictionary<string, string> RunBusinessRules();
    }
}
