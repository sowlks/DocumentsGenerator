using System;
using System.Collections.Generic;
using DocumentsGenerator.Core.Exceptions;

namespace DocumentsGenerator.Core.Tags.Properties
{
    [Property(PropertyNames.Group)]
    internal class GroupProperty : Property
    {
        public IEnumerable<string> Fields { get; private set; }

        public GroupProperty(string name, ITag parent, string? value)
            : base(name, parent, value)
        {
            if (string.IsNullOrEmpty(value))
                throw new PropertyValueEmptyException();

            Fields = value.Split(Consts.FieldsDelimiter);
        }
    }
}
