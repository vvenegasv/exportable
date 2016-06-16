using System.Collections.Generic;
using System.IO;
using Infodinamica.Framework.Exportable.Engines.Excel;

namespace Infodinamica.Framework.Exportable.Engines
{
    /// <summary>
    /// Interfaz para exportar
    /// </summary>
    public interface IExportEngine
    {
        /// <summary>
        /// Establece los datos a exportar
        /// </summary>
        /// <typeparam name="T">Tipo de dato a exportar</typeparam>
        /// <param name="data">Listado con datos a exportar</param>
        void AddData<T>(IList<T> data) where T : class;

        /// <summary>
        /// Exporta el documento
        /// </summary>
        /// <returns>MemoryStream del documento exportado</returns>
        MemoryStream Export();

        /// <summary>
        /// Exporta el documento
        /// </summary>
        /// <param name="path">Ruta en la cual se debe guardar</param>
        void Export(string path);

        /// <summary>
        /// Se ejecutan validaciones
        /// </summary>
        /// <returns>Diccionario con los errores ocurridos</returns>
        IDictionary<string, string> RunBusinessRules();

        /// <summary>
        /// Obtiene la interfaz para expotar a Excel
        /// </summary>
        /// <returns></returns>
        IExcelExportEngine AsExcel();
    }
}
