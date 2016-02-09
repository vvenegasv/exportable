using System.Collections.Generic;

namespace Infodinamica.Framework.Exportable.Engines.Excel
{
    public interface IExcelExportEngine: IExportEngine
    {
        void AddData<T>(IList<T> data, string sheetName) where T : class;
        void SetFormat(ExcelVersion version);
    }
}
