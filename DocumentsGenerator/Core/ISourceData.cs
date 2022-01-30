namespace DocumentsGenerator.Core
{
    internal interface ISourceData
    {
        ISourceDataStep? Current { get; }

        void AddStep(ISourceDataStep step);

        void RemoveLastStep();

        ISourceDataStep? FindStepByTable(string tableName);

        object CommonData { get; }
    }
}
