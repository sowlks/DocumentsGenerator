using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentsGenerator.Core;
using DocumentsGenerator.Core.Exceptions;
using DocumentsGenerator.Core.Tags;
using DocumentsGenerator.Core.Tags.Properties;

namespace DocumentsGenerator.Word.Tags
{
    [Tag(TagNames.ExistIf)]
    internal class ExistsIfTag : WordTag
    {
        public override bool IsContainer => true;

        public ExistsIfTag(string text)
            : base(text)
        {
        }

        protected override void SetSource(SourceData sourceData)
        {
            base.SetSource(sourceData);

            if (SourceData is DataSet data)
            {
                var table = data.Tables[TableName];
                SourceData = table.Rows.Cast<DataRow>();
            }
        }

        protected override void ChangeTemplate(SourceData sourceData)
        {
            base.ChangeTemplate(sourceData);

            ClearOpenTag();
            ClearCloseTag();

            bool isTrue = IsTrueCondition();
            if (isTrue)
            {
                foreach (ITag tag in this)
                {
                    tag.Process(sourceData);
                }
            }
            else
            {
                var firstElement = ParentOpenXmlElement;
                var lastElement = GetElementEqualLevelCloseTag();

                if (firstElement == null || lastElement == null)
                    throw new TagsStructureException(Text, "Open tags and/or close tag not found.");

                RemoveRangeElements(firstElement, lastElement);
            }

            RemoveEmptyOpenTag();
            RemoveEmptyCloseTag();
        }

        private bool IsTrueCondition()
        {
            var propCondition = GetProperty(PropertyNames.Condition, true) as ConditionProperty;

            var value = GetStringValueFromRow(propCondition.ParameterName);
            return value == propCondition.ParameterValue;
        }

        private string GetStringValueFromRow(string parameter)
        {
            if (SourceData is IEnumerable<DataRow> rows)
            {
                if (rows.Count() > 0 && rows.ElementAt(0).Table.Columns.Contains(parameter))
                {
                    if (rows.ElementAt(0)[parameter] != DBNull.Value)
                        return rows.ElementAt(0)[parameter].ToString();
                    else
                        return "";
                }
            }

            return "";
        }

        private OpenXmlElement? GetElementEqualLevelCloseTag()
        {
            if (ParentCloseXmlElement == null)
                return null;

            var current = ParentCloseXmlElement;
            while (true)
            {
                if (current.Parent == null)
                    return null;

                if (current.Parent == ParentOpenXmlElement?.Parent)
                    break;

                current = current.Parent;
            }

            return current;
        }

        private void RemoveRangeElements(OpenXmlElement firstElement, OpenXmlElement lastElement)
        {
            var parentElements = firstElement.Parent;
            bool begin = false;
            foreach (var element in parentElements.ToList())
            {
                if (!begin)
                {
                    if (element == firstElement)
                        begin = true;
                    else if (element == lastElement)
                        throw new TagsStructureException(Text, "Close tag locates above open tag.");
                }

                if (begin)
                {
                    parentElements?.RemoveChild(element);

                    if (element == lastElement)
                        break;
                }
            }
        }
    }
}
