namespace DocumentsGenerator.Core.Tags.Properties
{
    [Property(PropertyNames.Split)]
    internal class SplitProperty : Property
    {
        public SplitMode Mode { get; private set; }

        public SplitProperty(string name, ITag parent, string? value)
            : base(name, parent, value)
        {
            if (string.IsNullOrWhiteSpace(value))
                Mode = SplitMode.ClearValues;
            else
                Mode = value == "1" ? SplitMode.CopyValues : SplitMode.ClearValues;
        }
    }
}
