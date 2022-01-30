using System;
using System.Collections.Generic;
using System.Reflection;
using DocumentsGenerator.Core.Exceptions;

namespace DocumentsGenerator.Core.Tags.Properties
{
    public static class PropertiesFactory
    {
        private static readonly Dictionary<string, Type> Properties = new Dictionary<string, Type>();
        private static bool isInit = false;

        internal static void Init()
        {
            if (isInit)
                return;

            var assembly = Assembly.GetExecutingAssembly();
            foreach (var assType in assembly.GetTypes())
            {
                var propAttr = assType.GetCustomAttribute<PropertyAttribute>();
                if (propAttr != null)
                {
                    Properties.Add(propAttr.PropertyName, assType);
                }
            }

            isInit = true;
        }

        internal static IProperty Create(string text, ITag parent)
        {
            try
            {
                ExtractData(text, out var name, out var value);

                if (name == null || !Properties.ContainsKey(name.ToLower()))
                    throw new Exception("Unknown property.");

                return (IProperty)Activator.CreateInstance(Properties[name], name, parent, value);
            }
            catch (Exception err)
            {
                throw new ParsePropertyException(text, err);
            }
        }

        private static void ExtractData(string propertyText, out string name, out string value)
        {
            var strs = propertyText.Split(Consts.NameAndPropertiesDelimiter, 2);
            name = strs[0].ToLower();
            value = strs.Length > 1 ? strs[1] : "";
        }
    }
}
