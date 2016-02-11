﻿using System.Collections.Generic;
using System.IO;
using Infodinamica.Framework.Exportable.Tools;

namespace Infodinamica.Framework.Exportable.Engines
{
    public interface IExportEngine
    {
        void AddData<T>(IList<T> data) where T : class;
        MemoryStream Export();
        void Export(string path);
        IDictionary<string, string> RunBusinessRules();
    }
}