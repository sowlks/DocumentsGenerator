using System;

namespace DocumentsGenerator.Core.Exceptions
{
    public class ParseTagException : Exception
    {
        public ParseTagException(string tagName, Exception innerException)
            : base($"Tag \"{tagName}\". See inner exception", innerException)
        {
        }
    }
}
