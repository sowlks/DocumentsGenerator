namespace DocumentsGenerator.Core.Tags.Properties
{
    [Property(PropertyNames.RemoveEmpty)]
    internal class RemoveEmptyProperty : Property
    {
        public RemoveEmptyProperty(string name, ITag parent, string? value)
            : base(name, parent, value)
        {
        }
    }
}
