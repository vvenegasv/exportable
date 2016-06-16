using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infodinamica.Framework.Exportable.Engines.Excel
{
    /// <summary>
    /// Interfaz para importar un documento Excel
    /// </summary>
    public interface IExcelImportEngine: IImportEngine
    {
        /// <summary>
        /// Establece el tipo de dato que se obtendrá al leer una hoja Excel
        /// </summary>
        /// <typeparam name="T">Tipo de dato a retornar</typeparam>
        /// <param name="key">Nombre único asociado al tipo de dato</param>
        /// <param name="sheetName">Nombre de la hoja</param>
        void AddContainer<T>(string key, string sheetName) where T : class;

        /// <summary>
        /// Establece el tipo de dato que se obtendrá al leer una hoja Excel
        /// </summary>
        /// <typeparam name="T">Tipo de dato a retornar</typeparam>
        /// <param name="key">Nombre único asociado al tipo de dato</param>
        /// <param name="sheetName">Nombre de la hoja</param>
        /// <param name="firsRowWithData">Primera fila con datos zero-based</param>
        void AddContainer<T>(string key, string sheetName, int firsRowWithData) where T : class;

        /// <summary>
        /// Establece el formato del documento a leer
        /// </summary>
        /// <param name="version">Versión de Excel</param>
        void SetFormat(ExcelVersion version);
    }
}
