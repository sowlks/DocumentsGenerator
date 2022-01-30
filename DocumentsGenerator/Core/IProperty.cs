namespace DocumentsGenerator.Core
{
    internal interface IProperty
    {
        ITag Parent { get; }

        string Name { get; }

        string? Value { get; }
    }
}
