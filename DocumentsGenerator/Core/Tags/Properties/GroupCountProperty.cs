using DocumentsGenerator.Core.Exceptions;

namespace DocumentsGenerator.Core.Tags.Properties
{
    [Property(PropertyNames.GroupCount)]
    internal class GroupCountProperty : Property
    {
        public int Count { get; private set; }

        public GroupCountProperty(string name, ITag parent, string? value)
            : base(name, parent, value)
        {
            if (string.IsNullOrEmpty(value))
                throw new PropertyValueEmptyException();

            Count = int.Parse(value);
        }
    }
}
