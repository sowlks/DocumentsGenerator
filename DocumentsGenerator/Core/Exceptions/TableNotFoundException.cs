using System;

namespace DocumentsGenerator.Core.Exceptions
{
    public class TableNotFoundException : Exception
    {
        public TableNotFoundException(string tableName)
            : base($"Table \"{tableName}\" not found in source data set.")
        {
        }
    }
}
