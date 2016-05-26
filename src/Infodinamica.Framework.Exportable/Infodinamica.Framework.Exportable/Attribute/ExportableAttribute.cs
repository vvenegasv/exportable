using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Infodinamica.Framework.Exportable.Tools;

namespace Infodinamica.Framework.Exportable.Attribute
{
    /// <summary>
    /// Gestiona la forma en que se presentarán los datos
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ExportableAttribute: System.Attribute
    {
        private int _position;
        private string _headerName;
        private FieldValueType _typeValue;
        private string _format;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="position">Posición de la columna zero-based</param>
        /// <param name="headerName">Nombre de la columna</param>
        public ExportableAttribute(int position, string headerName)
        {
            _position = position;
            _headerName = headerName;
            _typeValue = FieldValueType.Any;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="position">Posición de la columna zero-based</param>
        /// <param name="headerName">Nombre de la columna</param>
        /// <param name="typeValue">Tipo de dato en la cual se representará el valor</param>
        public ExportableAttribute(int position, string headerName, FieldValueType typeValue)
        {
            _position = position;
            _headerName = headerName;
            _typeValue = typeValue;
            _format = string.Empty;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="position">Posición de la columna zero-based</param>
        /// <param name="headerName">Nombre de la columna</param>
        /// <param name="typeValue">Tipo de dato en la cual se representará el valor</param>
        /// <param name="format">Formato en la cual se representará el valor</param>
        public ExportableAttribute(int position, string headerName, FieldValueType typeValue, string format)
        {
            _position = position;
            _headerName = headerName;
            _typeValue = typeValue;
            _format = format;
        }

        /// <summary>
        /// Obtiene la posición
        /// </summary>
        /// <returns></returns>
        public int GetPosition()
        {
            return _position;
        }

        /// <summary>
        /// Obtiene el nombre de la columna
        /// </summary>
        /// <returns></returns>
        public string GetHeaderName()
        {
            return _headerName;
        }

        /// <summary>
        /// Obtiene el tipo de dato a representar
        /// </summary>
        /// <returns></returns>
        public FieldValueType GetTypeValue()
        {
            return _typeValue;
        }

        /// <summary>
        /// Obtiene el formato a representar
        /// </summary>
        /// <returns></returns>
        public string GetFormat()
        {
            return _format;
        }
    }
}
