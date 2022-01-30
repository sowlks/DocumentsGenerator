using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentsGenerator.Core.Tags.Properties;

namespace DocumentsGenerator.Word.Tags
{
    internal class SplitCell
    {
        public OpenXmlElement ElementTag { get; private set; }

        public TableCell Cell { get; private set; }

        public SplitMode Mode { get; private set; }

        public IEnumerable<string> Strings { get; private set; }

        public SplitCell(OpenXmlElement elementTag, TableCell cell, SplitMode mode)
        {
            ElementTag = elementTag;
            Cell = cell;
            Mode = mode;
            Strings = SplitTextByRows(Cell);
        }

        private IEnumerable<string> SplitTextByRows(TableCell cell)
        {
            var retVal = new List<string>();

            var splitTextByLines = cell.InnerText.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            foreach (var line in splitTextByLines)
            {
                var splitText = line.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                var cellWidth = WordHelper.GetCellActualWidth(cell);
                var text = new StringBuilder();
                var countWords = 0;

                for (var i = 0; i < splitText.Length; i++)
                {
                    var beforeText = new StringBuilder(text.ToString());

                    if (text.Length > 0)
                        text.Append(" ");

                    ++countWords;
                    text.Append(splitText[i]);
                    double textWidth = WordHelper.GetTextActualWidth(cell, text.ToString());

                    if (textWidth > cellWidth)
                    {
                        if (countWords > 1)
                        {
                            text = beforeText;
                            --i;
                        }

                        retVal.Add(text.ToString());
                        text.Clear();
                        countWords = 0;
                    }
                }

                retVal.Add(text.ToString());
            }

            return retVal;
        }
    }
}
