using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentsGenerator.Core;
using DocumentsGenerator.Core.Exceptions;
using DocumentsGenerator.Core.Tags;
using DocumentsGenerator.Core.Tags.Properties;
using DocumentsGenerator.Utils;

namespace DocumentsGenerator.Word.Tags
{
    [Tag(TagNames.Rows)]
    internal class RowsTag : WordTag
    {
        public RowsTag(string text)
            : base(text)
        {
        }

        public TableRow Row
        {
            get
            {
                var element = OpenXmlHelper.FindParent<TableRow>(ParentOpenXmlElement);
                if (element == null)
                    throw new TagsStructureException(Text, "Tag out of row of table.");

                return (TableRow)element;
            }
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

            var rowXml = Row;

            ClearOpenTag();
            ClearCloseTag();

            IEnumerable<object> data;
            var propertyGroup = GetProperty(PropertyNames.Group);
            if (propertyGroup == null)
                data = SourceData == null ? new DataRow[0] : (IEnumerable<DataRow>)SourceData;
            else
                data = GetGroups();

            foreach (var item in data)
            {
                TableRow rowCopy = (TableRow)rowXml.CloneNode(true);
                var tagDocument = new DocumentTag();
                DocumentHandler.FillTags(rowCopy, tagDocument, this);

                foreach (ITag tag in tagDocument)
                {
                    sourceData.AddStep(new SourceDataStep(this, item));

                    tag.Process(sourceData);

                    sourceData.RemoveLastStep();
                }

                if (rowXml.Parent != null)
                    rowXml.Parent.InsertBefore(rowCopy, rowXml);
            }

            ProcessPropertyOneStr(sourceData);
            ProcessPropertyRowSpan(sourceData);
            ProcessPropertySplit(sourceData);

            if (rowXml.Parent != null)
                rowXml.Parent.RemoveChild(rowXml);

            RemoveEmptyOpenTag();
            RemoveEmptyCloseTag();
        }

        protected void ProcessPropertyRowSpan(SourceData sourceData)
        {
            var oneStrProps = this.SelectMany(x => x.Properties).Where(y => y.Name == PropertyNames.RowSpan);

            foreach (var property in oneStrProps)
            {
                var tag = property.Parent;
                var cachedElements = sourceData.GetCachedChangeElements(property);
                if (cachedElements != null)
                {
                    var indexFirst = 0;
                    for (var i = 1; i < cachedElements.Count; i++)
                    {
                        var prevElement = cachedElements.ElementAt(i - 1);
                        var curElement = cachedElements.ElementAt(i);

                        if (prevElement.InnerText == curElement.InnerText)
                        {
                            var curCellFound = OpenXmlHelper.FindParent<TableCell>(curElement);

                            if (indexFirst == i - 1)
                            {
                                var prevCellFound = OpenXmlHelper.FindParent<TableCell>(prevElement);
                                if (prevCellFound is TableCell prevCell)
                                    prevCell.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
                            }

                            if (curCellFound is TableCell curCell)
                                curCell.Append(new VerticalMerge() { Val = MergedCellValues.Continue });
                        }
                        else
                        {
                            indexFirst = i;
                        }
                    }
                }

                sourceData.ClearCachedChangeElements(property);
            }
        }

        protected void ProcessPropertyOneStr(SourceData sourceData)
        {
            var oneStrProps = this.SelectMany(x => x.Properties).Where(y => y.Name == PropertyNames.OneStr);

            foreach (OneStringProperty property in oneStrProps)
            {
                var tag = property.Parent;
                var cachedElements = sourceData.GetCachedChangeElements(property);
                if (cachedElements != null)
                {
                    var texts = cachedElements.Select(x => x.InnerText);
                    var joinText = string.Join(property.Delimiter, texts.Distinct());

                    foreach (var element in cachedElements)
                    {
                        SetText(element, joinText);
                    }
                }

                sourceData.ClearCachedChangeElements(property);
            }
        }

        protected void ProcessPropertySplit(SourceData sourceData)
        {
            var splitProps = this.SelectMany(x => x.Properties).Where(y => y.Name == PropertyNames.Split);

            var splitCells = SplitCellsHandler.GetSplitCells(sourceData, splitProps);

            SplitCellsHandler.ProcessPropertySplit(splitCells);
        }

        private IEnumerable<IEnumerable<DataRow>> GetGroups()
        {
            var retVal = new List<IEnumerable<DataRow>>();

            if (SourceData is IEnumerable<DataRow> rows)
            {
                if (GetProperty(PropertyNames.Group) is GroupProperty property)
                {
                    foreach (var group in rows.GroupBy(x => x, new GroupDataRowComparer(property.Fields)))
                    {
                        retVal.Add(group.ToList());
                    }
                }
            }

            return retVal;
        }
    }
}
