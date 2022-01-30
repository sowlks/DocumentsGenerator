using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;

namespace DocumentsGenerator.Core
{
    internal interface ITag : IEnumerable<ITag>
    {
        Guid Id { get; set; }

        string Text { get; }

        string TableName { get; }

        ITag? Parent { get; set; }

        OpenXmlElement? ParentOpenXmlElement { get; set; }

        OpenXmlElement? ParentCloseXmlElement { get; set; }

        IEnumerable<IProperty>? Properties { get; }

        /// <summary>
        /// <c>True</c> if this tag can contains inner tags.
        /// </summary>
        bool IsContainer { get; }

        void Process(ISourceData sourceData);

        void AddInnerTag(ITag aTag);

        /// <summary>
        /// Gets property by name.
        /// </summary>
        /// <param name="name">Name of property.</param>
        /// <param name="required">If required and returned property is empty then throw exception.</param>
        /// <returns>Return property.</returns>
        IProperty? GetProperty(string name, bool required = true);
    }
}
