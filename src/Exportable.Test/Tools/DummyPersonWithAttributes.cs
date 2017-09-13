﻿using System;
using Exportable.Attribute;
using Exportable.Models;

namespace Exportable.Test.Tools
{
    [ImportableExcelHeader("Dummy People")]
    class DummyPersonWithAttributes
    {
        [Importable(0)]
        [Exportable(0, "Full Name", FieldValueType.Text)]
        public string Name { get; set; }

        [Importable(1)]
        [Exportable(1, "Birth Date", FieldValueType.Date, "dd/MM/yyyy")]
        public DateTime BirthDate { get; set; }

        [Importable(2)]
        [Exportable(2, "How Many Years", FieldValueType.Numeric, "#0")]
        public int Age { get; set; }

        [Importable(3)]
        [Exportable(3, "Is Adult", FieldValueType.Bool)]
        public bool IsAdult { get; set; }
    }
}
