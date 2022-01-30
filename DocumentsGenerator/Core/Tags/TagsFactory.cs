using System;
using System.Collections.Generic;
using System.Reflection;
using DocumentsGenerator.Core.Exceptions;

namespace DocumentsGenerator.Core.Tags
{
    internal class TagsFactory
    {
        private static readonly Dictionary<string, Type> Tags = new Dictionary<string, Type>();
        private static bool isInit = false;

        internal static void Init()
        {
            if (isInit)
                return;

            var assembly = Assembly.GetExecutingAssembly();
            foreach (var assType in assembly.GetTypes())
            {
                var tagAttr = assType.GetCustomAttribute<TagAttribute>();
                if (tagAttr != null)
                {
                    Tags.Add(tagAttr.TagName, assType);
                }
            }

            isInit = true;
        }

        internal static ITag Create(string text)
        {
            try
            {
                var tagName = ExtractTagName(text);

                if (tagName == null)
                    throw new Exception("Unknown tag.");

                if (!Tags.ContainsKey(tagName))
                    tagName = TagNames.Value;

                return (ITag)Activator.CreateInstance(Tags[tagName], text);
            }
            catch (Exception err)
            {
                throw new ParseTagException(text, err);
            }
        }

        public static string? ExtractTagName(string textTag)
        {
            var index = textTag.IndexOf(Consts.NamesDelimiter);
            if (index > 0)
            {
                var index2 = textTag.IndexOf(Consts.MainAndPropertiesDelimiter);
                index2 = index2 < 0 ? textTag.Length - 1 : index2 - 1;
                if (index2 - index > 0)
                {
                    var tagName = textTag.Substring(index + 1, index2 - index);

                    return tagName.ToLower();
                }
            }

            return null;
        }
    }
}
