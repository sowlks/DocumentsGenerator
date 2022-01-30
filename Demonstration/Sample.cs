using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using DocumentsGenerator.Word;

namespace Demonstration
{
    internal class Sample
    {
        public string Name { get; private set; }

        public string UseTags { get; private set; }

        public string TemplateName { get; private set; }

        public DelegateCommand GenerateCommand => new (GenerateDocument);

        public Sample(string name, string useTags, string templateName)
        {
            Name = name;
            UseTags = useTags;
            TemplateName = templateName;
        }

        private void GenerateDocument()
        {
            try
            {
                var templateDoc = $"{TemplateName}.docx";
                var tempFile = Path.Combine(Path.GetTempPath(), $"{Path.GetRandomFileName()}.docx");
                File.Copy(templateDoc, tempFile, true);

                var templateSet = $"{TemplateName}.xml";
                var dataSet = new DataSet();
                dataSet.ReadXml(templateSet);

                var generator = new WordGenerator(tempFile, dataSet);
                generator.Generate();

                var process = new Process
                {
                    StartInfo = new ProcessStartInfo(tempFile)
                    {
                        UseShellExecute = true
                    }
                };

                process.Start();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }
    }
}
