using System;

namespace DocumentsGenerator.Core.Exceptions
{
    public class TagsStructureException : Exception
    {
        public TagsStructureException(string tag, string message)
            : base($"Tag \"{tag}\". {message}")
        {
        }
    }
}
