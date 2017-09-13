using System;

namespace Exportable.Attribute
{
    /// <summary>
    /// Información para leer un documento
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ImportableAttribute : System.Attribute
    {
        /// <summary>
        /// Posición de la columna que debe ser leida
        /// </summary>
        public int Position { get; set; }

        /// <summary>
        /// Valor por defecto en caso que el dato contenido en la celda no exista o sea inválido
        /// </summary>
        public string DefaultForNullOrInvalidValues { get; set; }

        /// <summary>
        /// Constructor de la clase
        /// </summary>
        /// <param name="position"></param>
        public ImportableAttribute(int position)
        {
            this.Position = position;
        }
    }
}
