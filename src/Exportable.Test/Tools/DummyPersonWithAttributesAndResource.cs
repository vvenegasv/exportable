using System;
using Exportable.Attribute;
using Exportable.Tools;

namespace Exportable.Test.Tools
{
    class DummyPersonWithAttributesAndResource
    {
        [Exportable(Position = 0, HeaderName = "Header1", ResourceType = typeof(res), TypeValue = FieldValueType.Text)]
        public string Name { get; set; }

        [Exportable(1, "Header2", FieldValueType.Date, "MM-yyyy", ResourceType = typeof(res))]
        public DateTime BirthDate { get; set; }

        [Exportable(IsIgnored = true)]
        public int Age { get; set; }

        [Exportable(3, "Is Adult", FieldValueType.Bool)]
        public bool IsAdult { get; set; }
    }
}
