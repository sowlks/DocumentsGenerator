using System;

namespace DocumentsGenerator.Core.Exceptions
{
    public class TableColumnNotFoundException : Exception
    {
        public TableColumnNotFoundException(string table, string column)
            : base($"Column \"{column}\" in table \"{table}\" not found.")
        {
        }
    }
}
