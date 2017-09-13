using System;
using Exportable.Attribute;
using Exportable.Models;

namespace Exportable.Test.Tools
{
    [ImportableExcelHeader("Dummy People")]
    class DummyPersonWithAttributesAndDefaultValues
    {
        [Importable(0)]
        [Exportable(0, "Full Name", FieldValueType.Text)]
        public string Name { get; set; }

        [Importable(1, DefaultForNullOrInvalidValues = "05/05/1900")]
        [Exportable(1, "Birth Date", FieldValueType.Date, "dd/MM/yyyy")]
        public DateTime BirthDate { get; set; }

        [Importable(2, DefaultForNullOrInvalidValues = "0")]
        [Exportable(2, "How Many Years", FieldValueType.Numeric, "#0")]
        public int Age { get; set; }

        [Importable(3, DefaultForNullOrInvalidValues = "false")]
        [Exportable(3, "Is Adult", FieldValueType.Bool)]
        public bool IsAdult { get; set; }
    }
}
