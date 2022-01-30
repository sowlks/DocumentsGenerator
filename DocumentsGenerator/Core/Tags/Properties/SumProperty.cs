namespace DocumentsGenerator.Core.Tags.Properties
{
    [Property(PropertyNames.Sum)]
    internal class SumProperty : Property
    {
        public SumProperty(string name, ITag parent, string? value)
            : base(name, parent, value)
        {
        }
    }
}
