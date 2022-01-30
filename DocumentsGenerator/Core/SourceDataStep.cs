namespace DocumentsGenerator.Core
{
    internal class SourceDataStep : ISourceDataStep
    {
        public ITag Tag { get; private set; }

        public object Data { get; private set; }

        public SourceDataStep(ITag tag, object data)
        {
            Data = data;
            Tag = tag;
        }
    }
}
