namespace DocumentsGenerator.Core.Tags.Properties
{
    internal abstract class Property : IProperty
    {
        public ITag Parent { get; private set; }

        public string Name { get; private set; }

        public string? Value { get; private set; }

        protected Property(string name, ITag parent, string? value)
        {
            Parent = parent;
            Name = name;
            Value = value;
        }
    }
}
