using System;

namespace DocumentsGenerator.Core.Exceptions
{
    public class ParsePropertyException : Exception
    {
        public ParsePropertyException(string propertyText, Exception innerException)
            : base($"Property \"{propertyText}\". See inner exception", innerException)
        {
        }
    }
}
