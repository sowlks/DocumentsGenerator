namespace DocumentsGenerator.Core
{
    internal interface ISourceDataStep
    {
        ITag Tag { get; }

        object Data { get; }
    }
}
