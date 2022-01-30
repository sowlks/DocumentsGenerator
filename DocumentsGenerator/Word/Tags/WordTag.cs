using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentsGenerator.Utils;

namespace DocumentsGenerator.Word.Tags
{
    internal abstract class WordTag : Core.Tags.Tag
    {
        protected WordTag(string? text)
            : base(text)
        {
        }

        protected void RemoveEmptyOpenTag()
        {
            if (ParentOpenXmlElement != null && ParentOpenXmlElement.Parent != null && ParentOpenXmlElement.InnerText == "")
                ParentOpenXmlElement.Remove();
        }

        protected void RemoveEmptyCloseTag()
        {
            if (ParentCloseXmlElement != null && ParentCloseXmlElement.Parent != null && ParentCloseXmlElement.InnerText == "")
                ParentCloseXmlElement.Remove();
        }

        protected IEnumerable<OpenXmlElement>? GetCellsColumnByElement(OpenXmlElement element)
        {
            var elementFound = OpenXmlHelper.FindParent<TableCell>(element);
            if (elementFound is TableCell cell)
            {
                elementFound = OpenXmlHelper.FindParent<TableRow>(cell);
                if (elementFound is TableRow row)
                {
                    elementFound = OpenXmlHelper.FindParent<Table>(row);
                    if (elementFound is Table table)
                    {
                        var index = 0;
                        foreach (var curCell in row)
                        {
                            if (curCell == cell)
                                break;

                            ++index;
                        }

                        return table.Select(x => x.ElementAt(index)).ToList().AsEnumerable();
                    }
                }
            }

            return null;
        }

        protected override void SetText(OpenXmlElement element, object value)
        {
            WordHelper.SetText(element, value);
        }

        protected override void ReplaceText(OpenXmlElement element, string searchText, object? replaceValue)
        {
            WordHelper.ReplaceTextWord((Paragraph)element, searchText, replaceValue?.ToString());
        }
    }
}
