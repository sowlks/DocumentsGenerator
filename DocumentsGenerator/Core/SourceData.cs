using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml;

namespace DocumentsGenerator.Core
{
    internal class SourceData : ISourceData
    {
        public ISourceDataStep? Current => steps == null || steps.Count == 0 ? null : steps.Last();

        public DataSet CommonData { get; private set; }

        object ISourceData.CommonData => this.CommonData;

        private readonly List<ISourceDataStep> steps = new List<ISourceDataStep>();

        private readonly Dictionary<string, List<OpenXmlElement>> cacheChangeElements = new Dictionary<string, List<OpenXmlElement>>();

        public SourceData(DataSet source)
        {
            CommonData = source;
        }

        private string GetKey(IProperty property)
        {
            return $"{property.Parent.Id}{property.Name}";
        }

        public void AddCachedChangeElement(IProperty property, OpenXmlElement? element)
        {
            if (element == null)
                return;

            string key = GetKey(property);

            if (!cacheChangeElements.ContainsKey(key))
                cacheChangeElements.Add(key, new List<OpenXmlElement>());

            cacheChangeElements[key].Add(element);
        }

        public List<OpenXmlElement>? GetCachedChangeElements(IProperty property)
        {
            string key = GetKey(property);

            return cacheChangeElements.ContainsKey(key) ? cacheChangeElements[key] : null;
        }

        public void ClearCachedChangeElements(IProperty property)
        {
            string key = GetKey(property);

            if (cacheChangeElements.ContainsKey(key))
                cacheChangeElements.Remove(key);
        }

        public void AddStep(ISourceDataStep step)
        {
            steps.Add(step);
        }

        public void RemoveLastStep()
        {
            if (steps.Count > 0)
                steps.RemoveAt(steps.Count - 1);
        }

        public ISourceDataStep? FindStepByTable(string tableName)
        {
            return steps.Reverse<ISourceDataStep>().FirstOrDefault(x => x.Tag.TableName == tableName);
        }
    }
}
