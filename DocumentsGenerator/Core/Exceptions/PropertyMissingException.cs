using System;

namespace DocumentsGenerator.Core.Exceptions
{
    public class PropertyMissingException : Exception
    {
        public PropertyMissingException(string tagText, string propertyName)
            : base($"Tag \"{tagText}\". Property \"{propertyName}\" is missing.")
        {
        }
    }
}
