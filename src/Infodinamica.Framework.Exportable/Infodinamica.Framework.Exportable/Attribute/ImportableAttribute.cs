using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infodinamica.Framework.Exportable.Attribute
{
    /// <summary>
    /// Información para leer un documento
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ImportableAttribute : System.Attribute
    {
        private int _position;
        
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="position">Posición de la columna</param>
        public ImportableAttribute(int position)
        {
            _position = position;
        }
        
        /// <summary>
        /// Obtiene la posición de la columna
        /// </summary>
        /// <returns></returns>
        public int GetPosition()
        {
            return _position;
        }
    }
}
