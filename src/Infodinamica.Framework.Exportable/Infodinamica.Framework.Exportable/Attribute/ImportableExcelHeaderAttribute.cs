using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infodinamica.Framework.Exportable.Attribute
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ImportableExcelHeaderAttribute : System.Attribute
    {
        private string _sheetName;
        private int _firstRowWithData;

        public ImportableExcelHeaderAttribute(string sheetName)
        {
            _sheetName = sheetName;
            _firstRowWithData = 1;
        }

        public ImportableExcelHeaderAttribute(string sheetName, int firstRowWithData)
        {
            _sheetName = sheetName;
            _firstRowWithData = firstRowWithData;
        }

        public string GetSheetName()
        {
            return _sheetName;
        }

        public int GetFirstRowWithData()
        {
            return _firstRowWithData;
        }
    }
}
