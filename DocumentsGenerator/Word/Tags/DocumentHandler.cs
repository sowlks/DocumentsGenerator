using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentsGenerator.Core;
using DocumentsGenerator.Core.Exceptions;
using DocumentsGenerator.Core.Tags;

namespace DocumentsGenerator.Word.Tags
{
    internal class DocumentHandler
    {
        private readonly WordprocessingDocument document;
        private readonly DataSet sourceData;
        private readonly DocumentTag tags;

        public DocumentHandler(WordprocessingDocument document, DataSet sourceData)
        {
            this.document = document;
            this.sourceData = sourceData;
            this.tags = CreateDocumentTag();
        }

        public void Process()
        {
            var source = new SourceData(sourceData);
            source.AddStep(new SourceDataStep(tags, sourceData));

            tags.Process(source);

            source.RemoveLastStep();
        }

        private DocumentTag CreateDocumentTag()
        {
            var tag = new DocumentTag();

            if (document.MainDocumentPart != null)
            {
                foreach (var header in document.MainDocumentPart.HeaderParts)
                {
                    FillTags(header.Header, tag);
                }

                foreach (var footer in document.MainDocumentPart.FooterParts)
                {
                    FillTags(footer.Footer, tag);
                }

                FillTags(document.MainDocumentPart.Document, tag);
            }

            var errors = tag.Correct();
            if (errors != null)
                throw new TagsStructureException(tag.Text, errors);

            return tag;
        }

        public static void FillTags(OpenXmlElement element, ITag tag, ITag? sourceTag = null)
        {
            ITag? tagRef = tag;
            FillTagsRecursive(element, ref tagRef);

            if (sourceTag != null)
                CopyTagsId(sourceTag, tag);
        }

        private static void CopyTagsId(ITag tagFrom, ITag tagTo)
        {
            tagTo.Id = tagFrom.Id;

            for (var i = 0; i < tagFrom.Count(); i++)
                CopyTagsId(tagFrom.ElementAt(i), tagTo.ElementAt(i));
        }

        private static void FillTagsRecursive(OpenXmlElement parentElement, ref ITag? tag)
        {
            foreach (OpenXmlElement element in parentElement)
            {
                if (element is Paragraph paragraph)
                {
                    var run = element.Elements<Run>().FirstOrDefault();
                    if (run != null)
                    {
                        var text = run.Elements<Text>().FirstOrDefault();
                        if (text != null)
                        {
                            var textTags = GetTextTags(paragraph);

                            foreach (var textTag in textTags)
                            {
                                if (tag != null)
                                {
                                    if (textTag == "/")
                                    {
                                        tag.ParentCloseXmlElement = element;
                                        tag = tag.Parent;
                                    }
                                    else
                                    {
                                        var newTag = TagsFactory.Create(textTag);

                                        if (newTag != null)
                                        {
                                            newTag.ParentOpenXmlElement = element;
                                            newTag.Parent = tag;

                                            tag.AddInnerTag(newTag);

                                            if (newTag.IsContainer)
                                                tag = newTag;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                FillTagsRecursive(element, ref tag);
            }
        }

        private static IEnumerable<string> GetTextTags(Paragraph paragraph)
        {
            var retVal = new List<string>();
            var text = paragraph.InnerText;

            int index = -1;
            int indexStart = -1;
            int indexEnd;

            while (++index < text.Length)
            {
                if (text[index] == '[')
                {
                    indexStart = index;
                }
                else if (text[index] == ']')
                {
                    indexEnd = index;
                    if (indexEnd - indexStart > 1)
                    {
                        var tagText = text.Substring(indexStart + 1, indexEnd - indexStart - 1);
                        retVal.Add(tagText);
                    }
                }
            }

            return retVal;
        }
    }
}
