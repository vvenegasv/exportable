using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Infodinamica.Framework.Exportable.Tools;

namespace Infodinamica.Framework.Exportable.Attribute
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ExportableAttribute: System.Attribute
    {
        private int _position;
        private string _headerName;
        private FieldValueType _typeValue;
        private string _format;

        public ExportableAttribute(int position, string headerName)
        {
            _position = position;
            _headerName = headerName;
            _typeValue = FieldValueType.Any;
        }

        public ExportableAttribute(int position, string headerName, FieldValueType typeValue)
        {
            _position = position;
            _headerName = headerName;
            _typeValue = typeValue;
            _format = string.Empty;
        }

        public ExportableAttribute(int position, string headerName, FieldValueType typeValue, string format)
        {
            _position = position;
            _headerName = headerName;
            _typeValue = typeValue;
            _format = format;
        }

        public int GetPosition()
        {
            return _position;
        }

        public string GetHeaderName()
        {
            return _headerName;
        }

        public FieldValueType GetTypeValue()
        {
            return _typeValue;
        }

        public string GetFormat()
        {
            return _format;
        }
    }
}
