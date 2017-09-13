namespace Exportable.Tools
{
    internal class Metadata
    {
        public int Position { get; set; }
        public string Format { get; set; }
        public FieldValueType FieldValueType { get; set; }
        public string Name { get; internal set; }
        public string DefaultForNullOrInvalidValues { get; set; }

        public Metadata(string name, int position, string format, FieldValueType type, string defaultForNullOrInvalidValues)
        {
            Name = name;
            Position = position;
            Format = format;
            FieldValueType = type;
            DefaultForNullOrInvalidValues = defaultForNullOrInvalidValues;
        }
    }
}
