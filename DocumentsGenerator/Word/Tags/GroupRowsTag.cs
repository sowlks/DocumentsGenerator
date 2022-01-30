using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentsGenerator.Core;
using DocumentsGenerator.Core.Exceptions;
using DocumentsGenerator.Core.Tags;
using DocumentsGenerator.Core.Tags.Properties;
using DocumentsGenerator.Utils;

namespace DocumentsGenerator.Word.Tags
{
    [Tag(TagNames.GroupRows)]
    internal class GroupRowsTag : WordTag
    {
        public GroupRowsTag(string text)
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

            var rows = SourceData as IEnumerable<DataRow>;

            ClearOpenTag();
            ClearCloseTag();

            var firstElement = GetElementEqualLevelOpenTag();

            if (firstElement == null)
                throw new TagsStructureException(Text, "Start tag of group rows not found.");

            var lastElement = GetElementEqualLevelCloseTag(firstElement);

            if (lastElement == null)
                throw new TagsStructureException(Text, "End tag of group rows not found.");

            var rangeElements = GetRangeElements(firstElement, lastElement);
            var groups = GetGroups();
            var splitCells = new List<SplitCell>();

            foreach (var group in groups)
            {
                var elementCopy = CopyRangeElements(rangeElements);

                var documentTag = new DocumentTag();
                DocumentHandler.FillTags(elementCopy, documentTag, this);

                foreach (ITag tag in documentTag)
                {
                    sourceData.AddStep(new SourceDataStep(this, group));

                    tag.Process(sourceData);

                    var splitProps = this.SelectMany(x => x.Properties).Where(y => y.Name == PropertyNames.Split);
                    splitCells.AddRange(SplitCellsHandler.GetSplitCells(sourceData, splitProps));

                    sourceData.RemoveLastStep();
                }

                SplitCellsHandler.ProcessPropertySplit(splitCells);

                foreach (var element in elementCopy)
                {
                    firstElement.Parent?.InsertBefore(element.CloneNode(true), firstElement);
                }
            }

            var parent = firstElement.Parent;
            foreach (var element in rangeElements)
                parent?.RemoveChild(element);

            RemoveEmptyOpenTag();
            RemoveEmptyCloseTag();
        }

        private IEnumerable<IEnumerable<DataRow>> GetGroups()
        {
            var retVal = new List<IEnumerable<DataRow>>();

            var rows = SourceData as IEnumerable<DataRow>;

            if (GetProperty(PropertyNames.Group) is GroupProperty property)
            {
                foreach (var group in rows.GroupBy(x => x, new GroupDataRowComparer(property.Fields)))
                {
                    retVal.Add(group.ToList().AsEnumerable());
                }
            }
            else
            {
                if (GetProperty(PropertyNames.GroupCount) is GroupCountProperty propertyCount)
                {
                    var tempList = new List<DataRow>();
                    for (int i = 1; i <= rows.Count(); i++)
                    {
                        tempList.Add(rows.ElementAt(i - 1));

                        if (i % propertyCount.Count == 0 || i == rows.Count())
                        {
                            retVal.Add(tempList.ToList());
                            tempList.Clear();
                        }
                    }
                }
                else
                {
                    if (rows != null)
                    {
                        foreach (var row in rows)
                            retVal.Add(new DataRow[] { row });
                    }
                }
            }

            return retVal;
        }

        private OpenXmlElement CopyRangeElements(IEnumerable<OpenXmlElement> elements)
        {
            var body = new Body(elements.Select(x => x.CloneNode(true)));

            return body;
        }

        private IEnumerable<OpenXmlElement> GetRangeElements(OpenXmlElement firstElement, OpenXmlElement lastElement)
        {
            var listElements = new List<OpenXmlElement>();

            if (firstElement.Parent != null)
            {
                bool begin = false;
                foreach (var element in firstElement.Parent)
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
                        listElements.Add(element);

                        if (element == lastElement)
                            break;
                    }
                }
            }

            if (!(firstElement is TableRow) && listElements.Count > 1)
            {
                if (listElements.First().InnerText.Length == 0)
                    listElements = listElements.Skip(1).ToList();

                if (listElements.Last().InnerText.Length == 0)
                    listElements = listElements.Take(listElements.Count - 1).ToList();
            }

            return listElements;
        }

        private OpenXmlElement? GetElementEqualLevelOpenTag()
        {
            if (ParentOpenXmlElement == null)
                return null;

            var cell = OpenXmlHelper.FindParent<TableCell>(ParentOpenXmlElement, 2);
            if (cell == null)
                return ParentOpenXmlElement;
            else
                return cell.Parent;
        }

        private OpenXmlElement? GetElementEqualLevelCloseTag(OpenXmlElement? firstElement)
        {
            if (ParentCloseXmlElement == null)
                return null;

            if (firstElement is TableRow)
            {
                var row = OpenXmlHelper.FindParent<TableRow>(ParentCloseXmlElement);
                if (row != null && row.Parent == firstElement.Parent)
                    return row;
                else
                    throw new TagsStructureException(Text, "Close tag must be in the same row as open tag.");
            }
            else
            {
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
        }
    }
}
