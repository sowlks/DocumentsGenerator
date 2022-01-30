using System;

namespace DocumentsGenerator.Core.Tags.Properties
{
    internal class PropertyAttribute : Attribute
    {
        public string PropertyName { get; private set; }

        public PropertyAttribute(string propertyName)
        {
            if (string.IsNullOrEmpty(propertyName))
                throw new ArgumentNullException(nameof(propertyName));

            PropertyName = propertyName.ToLower();
        }
    }
}
