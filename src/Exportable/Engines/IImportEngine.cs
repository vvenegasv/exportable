using System.Collections.Generic;
using System.IO;
using Exportable.Engines.Excel;

namespace Exportable.Engines
{
    /// <summary>
    /// Interfaz para importar documentos
    /// </summary>
    public interface IImportEngine
    {
        /// <summary>
        /// Se establece el documento a importar
        /// </summary>
        /// <param name="file">MemoryStream del documento a leer</param>
        void SetDocument(MemoryStream file);

        /// <summary>
        /// Se establece el documento a importar
        /// </summary>
        /// <param name="path">Ruta en disco del documento a leer</param>
        void SetDocument(string path);

        /// <summary>
        /// Establece el tipo de dato que contendrá el resultado
        /// </summary>
        /// <typeparam name="T">Tipo de dato retornado</typeparam>
        /// <param name="key">Nombre único asociado al tipo de dato retornado</param>
        void AddContainer<T>(string key) where T : class;

        /// <summary>
        /// Obtiene el listado de datos
        /// </summary>
        /// <typeparam name="T">Tipo de dato retornado</typeparam>
        /// <param name="key">Nombre único para leer el documento excel</param>
        /// <returns>Listado del tipo de dato especificado</returns>
        IList<T> GetList<T>(string key) where T : class;

        /// <summary>
        /// Ejecuta una serie de validaciones
        /// </summary>
        /// <returns>Diccionario con las pruebas no superadas</returns>
        IDictionary<string, string> RunBusinessRules();

        /// <summary>
        /// Obtiene la interfaz para importar un documento Excel
        /// </summary>
        /// <returns></returns>
        IExcelImportEngine AsExcel();
    }
}
