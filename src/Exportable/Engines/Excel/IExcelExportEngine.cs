using System.Collections.Generic;

namespace Exportable.Engines.Excel
{
    /// <summary>
    /// Interfaz para exportar a un documento Excel
    /// </summary>
    public interface IExcelExportEngine: IExportEngine
    {
        /// <summary>
        /// Establece el listado de datos
        /// </summary>
        /// <typeparam name="T">Tipo de dato a exportar</typeparam>
        /// <param name="data">Listado de datos a exportar</param>
        /// <param name="sheetName">Nombre de la hoja</param>
        void AddData<T>(IList<T> data, string sheetName) where T : class;

        /// <summary>
        /// Establece el formato (XLS | XLSX)
        /// </summary>
        /// <param name="version">Formato Excel</param>
        void SetFormat(ExcelVersion version);
    }
}
