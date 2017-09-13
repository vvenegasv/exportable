using System;
using Exportable.Tools;

namespace Exportable.Attribute
{
    /// <summary>
    /// Gestiona la forma en que se presentarán los datos
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ExportableAttribute: System.Attribute
    {
        private bool _isIgnored = false;

        /// <summary>
        /// Posición zero-based en la que aparece la columna
        /// </summary>
        public int Position { get; set; }

        /// <summary>
        /// Nombre de la columna
        /// </summary>
        public string HeaderName { get; set; }

        /// <summary>
        /// Tipo de valor que se usara para mostrar 
        /// </summary>
        public FieldValueType TypeValue { get; set; }

        /// <summary>
        /// Formato en el cual se representará el valor
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// Recurso de localización
        /// </summary>
        public Type ResourceType { get; set; }

        /// <summary>
        /// Indica si la columna debe ser omitida en el proceso de exportación. Por defecto es falso
        /// </summary>
        public bool IsIgnored
        {
            get { return _isIgnored; }
            set { _isIgnored = value; }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public ExportableAttribute()
        {
            
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="position">Posición de la columna zero-based</param>
        /// <param name="headerName">Nombre de la columna</param>
        public ExportableAttribute(int position, string headerName)
        {
            Position = position;
            HeaderName = headerName;
            TypeValue = FieldValueType.Any;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="position">Posición de la columna zero-based</param>
        /// <param name="headerName">Nombre de la columna</param>
        /// <param name="typeValue">Tipo de dato en la cual se representará el valor</param>
        public ExportableAttribute(int position, string headerName, FieldValueType typeValue)
        {
            Position = position;
            HeaderName = headerName;
            TypeValue = typeValue;
            Format = string.Empty;
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
            Position = position;
            HeaderName = headerName;
            TypeValue = typeValue;
            Format = format;
        }
        
    }
}
