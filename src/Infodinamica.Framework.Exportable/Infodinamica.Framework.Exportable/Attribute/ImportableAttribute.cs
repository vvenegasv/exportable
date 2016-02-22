using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infodinamica.Framework.Exportable.Attribute
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ImportableAttribute : System.Attribute
    {
        private int _position;
        
        public ImportableAttribute(int position)
        {
            _position = position;
        }
        
        public int GetPosition()
        {
            return _position;
        }
    }
}
