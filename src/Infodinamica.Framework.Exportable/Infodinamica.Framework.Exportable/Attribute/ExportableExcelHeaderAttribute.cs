using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Infodinamica.Framework.Exportable.Tools;

namespace Infodinamica.Framework.Exportable.Attribute
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ExportableExcelHeaderAttribute : System.Attribute
    {
        public string FontColor { get; set; }
        public string FontName { get; set; }
        public short FontSize { get; set; }
        public string BorderColor { get; set; }
        public string BackColor { get; set; }


        public ExportableExcelHeaderAttribute()
        {

        }
    }
}