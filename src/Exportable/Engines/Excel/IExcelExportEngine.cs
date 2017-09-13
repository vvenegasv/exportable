using System.Collections.Generic;

namespace Exportable.Engines.Excel
{
    /// <summary>
    /// Interface to export data to Excel
    /// </summary>
    public interface IExcelExportEngine: IExportEngine
    {
        /// <summary>
        /// Add dta to export
        /// </summary>
        /// <typeparam name="TModel">Model type of data to export</typeparam>
        /// <param name="data">Array with data to export</param>
        /// <param name="sheetName">name of the worksheet to export</param>
        /// <returns>Key of the dataset to export</returns>
        string AddData<TModel>(IEnumerable<TModel> data, string sheetName) where TModel : class;

        /// <summary>
        /// Set the excel format (XLS | XLSX)
        /// </summary>
        /// <param name="version">Excel format</param>
        void SetFormat(ExcelVersion version);
    }
}
