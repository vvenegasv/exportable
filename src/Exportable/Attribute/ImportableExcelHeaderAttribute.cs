using System;

namespace Exportable.Attribute
{
    /// <summary>
    /// Define información para la lectura de un documento Excel
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ImportableExcelHeaderAttribute : System.Attribute
    {
        private string _sheetName;
        private int _firstRowWithData;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="sheetName">Nombre de la hoja</param>
        public ImportableExcelHeaderAttribute(string sheetName)
        {
            _sheetName = sheetName;
            _firstRowWithData = 1;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="sheetName">Nombre de la hoja</param>
        /// <param name="firstRowWithData">Primera fila con datos zero-based</param>
        public ImportableExcelHeaderAttribute(string sheetName, int firstRowWithData)
        {
            _sheetName = sheetName;
            _firstRowWithData = firstRowWithData;
        }

        /// <summary>
        /// Obtiene el nombre de la hoja
        /// </summary>
        /// <returns></returns>
        public string GetSheetName()
        {
            return _sheetName;
        }

        /// <summary>
        /// Obtiene la primera fila con datos
        /// </summary>
        /// <returns></returns>
        public int GetFirstRowWithData()
        {
            return _firstRowWithData;
        }
    }
}
