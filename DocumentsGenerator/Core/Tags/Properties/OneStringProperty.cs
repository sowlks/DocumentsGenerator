namespace DocumentsGenerator.Core.Tags.Properties
{
    [Property(PropertyNames.OneStr)]
    internal class OneStringProperty : Property
    {
        public string Delimiter { get; private set; }

        public OneStringProperty(string name, ITag parent, string? value)
            : base(name, parent, value)
        {
            if (string.IsNullOrEmpty(value))
                Delimiter = ", ";
            else
                Delimiter = value;
        }
    }
}
