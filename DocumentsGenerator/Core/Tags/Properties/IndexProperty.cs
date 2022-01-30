using DocumentsGenerator.Core.Exceptions;

namespace DocumentsGenerator.Core.Tags.Properties
{
    [Property(PropertyNames.Index)]
    internal class IndexProperty : Property
    {
        public int Index { get; private set; }

        public IndexProperty(string name, ITag parent, string? value)
            : base(name, parent, value)
        {
            if (string.IsNullOrEmpty(value))
                throw new PropertyValueEmptyException();

            Index = int.Parse(value);
        }
    }
}
