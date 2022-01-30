using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentsGenerator.Word
{
    internal class WordHelper
    {
        public static void SetText(OpenXmlElement element, object value)
        {
            if (element is Paragraph paragraph)
            {
                SetTextWord(paragraph, value?.ToString());
            }
            else
            {
                var paragraphs = element.Descendants<Paragraph>();
                if (paragraphs.Count() > 0)
                {
                    for (int i = 1; i < paragraphs.Count(); i++)
                        SetTextWord(paragraphs.ElementAt(i), "");

                    SetTextWord(paragraphs.ElementAt(0), value.ToString());
                }
            }
        }

        public static void SetTextWord(Paragraph element, string? text)
        {
            IEnumerable<Run> runs = element.Elements<Run>();
            if (runs.Count() > 0)
            {
                for (int i = 1; i < runs.Count(); i++)
                {
                    var elText1 = (Text)runs.ElementAt(i).Elements<Text>().FirstOrDefault();
                    if (elText1 != null)
                        elText1.Text = "";
                }

                var elText = (Text)runs.ElementAt(0).Elements<Text>().FirstOrDefault();
                if (elText != null)
                    elText.Text = text ?? "";
            }
        }

        public static void ReplaceTextWord(Paragraph element, string searchText, string? replaceText)
        {
            if (element == null)
                return;

            var index = -1;
            while (true)
            {
                if (element.InnerText.Length == 0 || index + 1 >= element.InnerText.Length)
                    break;

                index = element.InnerText.IndexOf(searchText, index + 1);

                if (index > -1)
                {
                    var runs = element.Elements<Run>();

                    if (runs.Count() == 1)
                    {
                        var text = (Text)runs.FirstOrDefault().Elements<Text>().FirstOrDefault();
                        text.Text = text.Text.Replace(searchText, replaceText);
                    }
                    else
                    {
                        var length = 0;
                        var indexInTextFirstRun = 0;
                        var indexFirstRun = -1;
                        var indexLastRun = -1;
                        for (var i = 0; i < runs.Count(); i++)
                        {
                            length += runs.ElementAt(i).InnerText.Length;

                            if (indexFirstRun == -1 && length > index)
                            {
                                indexFirstRun = i;
                                indexInTextFirstRun = index - (length - runs.ElementAt(i).InnerText.Length);
                            }

                            if (indexFirstRun > -1 && indexLastRun == -1 && index + searchText.Length <= length)
                            {
                                indexLastRun = i;
                                break;
                            }
                        }

                        if (indexFirstRun == indexLastRun)
                        {
                            var textRun = (Text)runs.ElementAt(indexFirstRun).Elements<Text>().FirstOrDefault();
                            if (textRun != null)
                                textRun.Text = textRun.Text.Replace(searchText, replaceText);
                        }
                        else
                        {
                            length = runs.ElementAt(indexFirstRun).InnerText.Length - indexInTextFirstRun;
                            for (int i = indexFirstRun + 1; i < indexLastRun; i++)
                            {
                                var textRun = (Text)runs.ElementAt(i).Elements<Text>().FirstOrDefault();
                                length += textRun.Text.Length;
                                textRun.Text = "";
                            }

                            var text = (Text)runs.ElementAt(indexLastRun).Elements<Text>().FirstOrDefault();
                            text.Text = text.Text[(searchText.Length - length)..];

                            text = (Text)runs.ElementAt(indexFirstRun).Elements<Text>().FirstOrDefault();
                            if (indexInTextFirstRun > 0)
                                text.Text = $"{text.Text[..indexInTextFirstRun]}{replaceText}";
                            else
                                text.Text = replaceText ?? "";
                        }
                    }
                }
                else
                {
                    break;
                }
            }
        }

        public static void SetFitCell(TableCell? cell)
        {
            if (cell?.TableCellProperties?.TableCellWidth?.Width != null)
            {
                var cellWidth = GetCellActualWidth(cell);
                var textWidth = GetTextActualWidth(cell);

                if (textWidth > cellWidth - 0.5)
                {
                    var fitText = new TableCellFitText();
                    cell.TableCellProperties.Append(fitText);
                }
            }
        }

        public static double GetTextActualWidth(OpenXmlElement element, string? text = null)
        {
            var runFonts = element.Descendants<RunFonts>().FirstOrDefault();
            var fontSize = element.Descendants<FontSize>().FirstOrDefault();
            var bold = element.Descendants<Bold>().FirstOrDefault();
            var italic = element.Descendants<Italic>().FirstOrDefault();

            var font = runFonts?.Ascii?.Value ?? "Times New Roman";
            var size = fontSize?.Val != null ? GetDoubleFromString(fontSize.Val.Value) / 2 : 12;

            FontStyle fontStyle = FontStyle.Regular;
            if (bold != null)
                fontStyle |= FontStyle.Bold;

            if (italic != null)
                fontStyle |= FontStyle.Italic;

            System.Drawing.Font fontDrawing = new System.Drawing.Font(font, (float)size, fontStyle);
            Image image = new Bitmap(1, 1);
            Graphics graphics = Graphics.FromImage(image);
            var drawText = string.IsNullOrEmpty(text) ? element.InnerText : text;

            var sizeText = graphics.MeasureString(drawText, fontDrawing, new SizeF(int.MaxValue, int.MaxValue), StringFormat.GenericDefault);

            double mm = sizeText.Width * (25.4 / image.VerticalResolution);

            return mm;
        }

        public static double GetCellActualWidth(TableCell cell)
        {
            double widthInDxa = GetDoubleFromString(cell?.TableCellProperties?.TableCellWidth?.Width?.Value);

            // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/;
            double widthInMm = widthInDxa / 56.69291338582677;

            var cellLeftMargin = cell?.Descendants<LeftMargin>().FirstOrDefault();
            var cellRightMargin = cell?.Descendants<RightMargin>().FirstOrDefault();

            // Default indent - 1.9 cm
            var leftPoints = 1.9;
            var rightPoints = 1.9;

            if (cellLeftMargin?.Width != null)
                leftPoints = GetDoubleFromString(cellLeftMargin.Width.Value) / 56.69291338582677;

            if (cellRightMargin?.Width != null)
                rightPoints = GetDoubleFromString(cellRightMargin.Width.Value) / 56.69291338582677;

            return widthInMm - leftPoints - rightPoints;
        }

        private static double GetDoubleFromString(string? inputString)
        {
            double.TryParse(inputString, out var dbl);

            return dbl;
        }
    }
}
