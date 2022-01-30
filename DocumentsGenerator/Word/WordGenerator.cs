using System;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentsGenerator.Core;
using DocumentsGenerator.Core.Tags;
using DocumentsGenerator.Core.Tags.Properties;
using DocumentsGenerator.Word.Tags;

namespace DocumentsGenerator.Word
{
    public class WordGenerator : IGenerator
    {
        public string Template { get; private set; }

        public DataSet Source { get; private set; }

        public WordGenerator(string template, DataSet source)
        {
            Template = template ?? throw new ArgumentNullException(nameof(template));
            Source = source ?? throw new ArgumentNullException(nameof(source));

            TagsFactory.Init();
            PropertiesFactory.Init();
        }

        public void Generate()
        {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(Template, true);
            var handler = new DocumentHandler(wordDocument, Source);
            handler.Process();

            wordDocument.Save();
        }
    }
}
