using System;

namespace DocumentsGenerator.Core.Tags
{
    internal class TagAttribute : Attribute
    {
        public string TagName { get; private set; }
        public TagAttribute(string tagName)
        {
            if (string.IsNullOrEmpty(tagName))
                throw new ArgumentNullException(nameof(tagName));

            TagName = tagName.ToLower();
        }
    }
}
