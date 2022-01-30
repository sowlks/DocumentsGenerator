using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentsGenerator.Core;
using DocumentsGenerator.Core.Tags.Properties;
using DocumentsGenerator.Utils;

namespace DocumentsGenerator.Word.Tags
{
    internal class SplitCellsHandler
    {
        public static IEnumerable<SplitCell> GetSplitCells(SourceData sourceData, IEnumerable<IProperty> splitProperies)
        {
            var splitCells = new List<SplitCell>();

            foreach (SplitProperty property in splitProperies)
            {
                var cachedElements = sourceData.GetCachedChangeElements(property);
                if (cachedElements != null)
                {
                    foreach (var sourceElement in cachedElements)
                    {
                        var cellFound = OpenXmlHelper.FindParent<TableCell>(sourceElement);
                        if (cellFound is TableCell cell)
                        {
                            var splitCell = new SplitCell(sourceElement, cell, property.Mode);
                            splitCells.Add(splitCell);
                        }
                    }
                }

                sourceData.ClearCachedChangeElements(property);
            }

            return splitCells;
        }

        public static void ProcessPropertySplit(IEnumerable<SplitCell> splitCells, OpenXmlElement? parent = null)
        {
            var rows = splitCells.GroupBy(x => OpenXmlHelper.FindParent<TableRow>(x.Cell) as TableRow);
            foreach (var rowCells in rows)
            {
                if (rowCells == null || rowCells.Key == null)
                    continue;

                var countRows = rowCells.Max(x => x.Strings.Count());
                var curRow = rowCells.Key;
                for (var i = 0; i < countRows; i++)
                {
                    if (i > 0)
                    {
                        var rowCopy = (TableRow)rowCells.Key.CloneNode(true);
                        (parent ?? curRow.Parent)?.InsertAfter(rowCopy, curRow);
                        curRow = rowCopy;
                    }

                    var splitMode = rowCells.ElementAt(0).Mode;
                    var firstRowCells = rowCells.Key.OfType<TableCell>();
                    var curRowCells = curRow.OfType<TableCell>();
                    for (var j = 0; j < firstRowCells.Count(); j++)
                    {
                        var cellWithPropSplit = rowCells.FirstOrDefault(x => x.Cell == firstRowCells.ElementAt(j));
                        if (cellWithPropSplit == null)
                        {
                            if (splitMode == SplitMode.ClearValues && i > 0)
                            {
                                WordHelper.SetText(curRowCells.ElementAt(j), "");
                            }
                        }
                        else
                        {
                            var newText = i < cellWithPropSplit.Strings.Count() ? cellWithPropSplit.Strings.ElementAt(i) : "";

                            var firstRowCellParagraphs = cellWithPropSplit.Cell.Elements<Paragraph>();
                            var indexParagraph = 0;
                            for (indexParagraph = 0; indexParagraph < firstRowCellParagraphs.Count(); indexParagraph++)
                            {
                                if (firstRowCellParagraphs.ElementAt(indexParagraph) == cellWithPropSplit.ElementTag)
                                    break;
                            }

                            OpenXmlElement findParagraph = curRowCells.ElementAt(j).Elements<Paragraph>().ElementAt(indexParagraph);
                            WordHelper.SetText(findParagraph, newText);

                            var findCell = OpenXmlHelper.FindParent<TableCell>(findParagraph) as TableCell;
                            WordHelper.SetFitCell(findCell);
                        }
                    }
                }
            }
        }
    }
}
