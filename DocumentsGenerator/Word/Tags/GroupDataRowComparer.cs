using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace DocumentsGenerator.Word.Tags
{
    internal class GroupDataRowComparer : IEqualityComparer<DataRow>
    {
        private readonly IEnumerable<string> columns;

        public GroupDataRowComparer(IEnumerable<string> columns)
        {
            this.columns = columns;
        }

        public bool Equals(DataRow x, DataRow y)
        {
            foreach (string column in columns)
            {
                string? value1 = null;
                string? value2 = null;

                if (x[column] != DBNull.Value)
                    value1 = x[column].ToString();

                if (x[column] != DBNull.Value)
                    value2 = x[column].ToString();

                if (!string.Equals(value1, value2))
                    return false;
            }

            return true;
        }

        public int GetHashCode(DataRow obj)
        {
            var text = new StringBuilder();
            foreach (string column in columns)
            {
                if (obj[column] != DBNull.Value)
                    text.Append(obj[column].ToString());
            }

            int hash = text.ToString().GetHashCode();
            return hash;
        }
    }
}
