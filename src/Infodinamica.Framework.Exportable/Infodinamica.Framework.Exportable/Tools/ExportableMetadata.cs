﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infodinamica.Framework.Exportable.Tools
{
    internal class ExportableMetadata
    {
        public int Position { get; set; }
        public string Format { get; set; }
        public FieldValueType FieldValueType { get; set; }
        public string Name { get; internal set; }

        public ExportableMetadata(string name, int position, string format, FieldValueType type)
        {
            Name = name;
            Position = position;
            Format = format;
            FieldValueType = type;
        }
    }
}