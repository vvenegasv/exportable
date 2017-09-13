using System;

namespace Exportable.Attribute
{
    /// <summary>
    /// Establece el formato de la primera fila
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ExportableExcelHeaderAttribute : System.Attribute
    {
        /// <summary>
        /// Color de la fuente en formato HTML hexadecimal
        /// </summary>
        public string FontColor { get; set; }

        /// <summary>
        /// Nombre de la fuente
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        /// Tamaño de la fuente
        /// </summary>
        public short FontSize { get; set; }

        /// <summary>
        /// Color del borde
        /// </summary>
        public string BorderColor { get; set; }

        /// <summary>
        /// Color de fondo
        /// </summary>
        public string BackColor { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public ExportableExcelHeaderAttribute()
        {

        }
    }
}