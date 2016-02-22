using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infodinamica.Framework.Exportable.Engines.Excel
{
    public interface IExcelImportEngine: IImportEngine
    {
        void AddContainer<T>(string key, string sheetName) where T : class;
        void AddContainer<T>(string key, string sheetName, int firsRowWithData) where T : class;
        void SetFormat(ExcelVersion version);
    }
}
