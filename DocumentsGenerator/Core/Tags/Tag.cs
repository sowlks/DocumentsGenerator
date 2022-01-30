using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentsGenerator.Core.Exceptions;
using DocumentsGenerator.Core.Tags.Properties;

namespace DocumentsGenerator.Core.Tags
{
    internal abstract class Tag : ITag
    {
        public Guid Id { get; set; }

        public string Text { get; private set; } = "";

        public string TableName { get; private set; } = "";

        public ITag? Parent { get; set; }

        public OpenXmlElement? ParentOpenXmlElement { get; set; }

        public OpenXmlElement? ParentCloseXmlElement { get; set; }

        public IEnumerable<IProperty>? Properties { get; protected set; }

        public virtual bool IsContainer => true;

        public object? SourceData { get; protected set; }

        private readonly List<ITag> tags = new List<ITag>();

        protected Tag(string? text)
        {
            Id = Guid.NewGuid();
            Text = text ?? "";

            if (!string.IsNullOrEmpty(text))
            {
                ExtractTableName(text);
                ExtractProperties(text);
            }
        }

        public void AddInnerTag(ITag tag)
        {
            tags.Add(tag);
        }

        private void ExtractTableName(string text)
        {
            var index = text.IndexOf(Consts.NamesDelimiter);
            if (index < 0)
                throw new Exception("Table name not found.");

            TableName = text[..index];
        }

        protected virtual void ExtractProperties(string text)
        {
            var properties = new List<IProperty>();

            var index = text.IndexOf(Consts.MainAndPropertiesDelimiter);
            if (index > -1)
            {
                var strProps = text[(index + 1)..];
                if (strProps.Length > 0)
                {
                    var props = strProps.Split(new string[] { Consts.PropertiesDelimiter }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string prop in props)
                    {
                        var property = PropertiesFactory.Create(prop, this);
                        properties.Add(property);
                    }
                }
            }

            Properties = properties;
        }

        public void Process(ISourceData sourceData)
        {
            if (sourceData == null)
                throw new ArgumentNullException(nameof(sourceData));

            SetSource((SourceData)sourceData);

            ChangeTemplate((SourceData)sourceData);
        }

        protected virtual void SetSource(SourceData sourceData)
        {
            if (!IsSpecialTag() && !string.IsNullOrEmpty(TableName) && !sourceData.CommonData.Tables.Contains(TableName))
                throw new TableNotFoundException(TableName);

            if (!IsSpecialTag() && sourceData.Current?.Tag.TableName != TableName)
            {
                var step = sourceData.FindStepByTable(TableName);
                if (step == null)
                    SourceData = sourceData.CommonData;
                else
                    SourceData = step.Data;
            }
            else
            {
                SourceData = sourceData.Current?.Data;
            }
        }

        protected virtual bool IsSpecialTag()
        {
            return false;
        }

        protected virtual void ChangeTemplate(SourceData sourceData)
        {
        }

        protected void ClearOpenTag()
        {
            if (ParentOpenXmlElement == null)
                return;

            var tagOpen = $"{Consts.Open}{Text}{Consts.Close}";

            ReplaceText(ParentOpenXmlElement, tagOpen, "");
        }

        protected void ClearCloseTag()
        {
            if (ParentCloseXmlElement == null)
                return;

            var tagClose = $"{Consts.Open}{Consts.EndContainer}{Consts.Close}";

            ReplaceText(ParentCloseXmlElement, tagClose, "");
        }

        protected void ReplaceTagText(OpenXmlElement? element, object value)
        {
            if (element == null)
                return;

            var tagText = $"{Consts.Open}{Text}{Consts.Close}";

            ReplaceText(element, tagText, value);
        }

        protected virtual void ReplaceText(OpenXmlElement element, string searchText, object replaceValue)
        {
        }

        protected virtual void SetText(OpenXmlElement element, object value)
        {
        }

        public IProperty? GetProperty(string name, bool required = false)
        {
            var prop = Properties.FirstOrDefault(x => string.Compare(x.Name, name, true) == 0);
            if (prop == null && required)
                throw new PropertyMissingException(Text, name);

            return prop;
        }

        public override string ToString()
        {
            return $"{Text}, inner elements count: {this.Count()}";
        }

        public void Dispose()
        {
            tags.Clear();
        }

        public IEnumerator<ITag> GetEnumerator()
        {
            return new TagEnumerator(tags);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
