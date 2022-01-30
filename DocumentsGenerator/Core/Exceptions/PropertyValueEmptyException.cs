using System;

namespace DocumentsGenerator.Core.Exceptions
{
    public class PropertyValueEmptyException : Exception
    {
        public PropertyValueEmptyException()
            : base("Value is null or empty.")
        {
        }
    }
}
