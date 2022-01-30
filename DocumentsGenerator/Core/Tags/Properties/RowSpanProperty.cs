namespace DocumentsGenerator.Core.Tags.Properties
{
    [Property(PropertyNames.RowSpan)]
    internal class RowSpanProperty : Property
    {
        public RowSpanProperty(string name, ITag parent, string? value)
            : base(name, parent, value)
        {
        }
    }
}
