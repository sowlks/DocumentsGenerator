using System.Data;

namespace DocumentsGenerator.Core
{
    public interface IGenerator
    {
        string Template { get; }

        DataSet Source { get; }

        void Generate();
    }
}
