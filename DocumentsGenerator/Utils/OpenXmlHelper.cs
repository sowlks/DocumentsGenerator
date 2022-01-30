using DocumentFormat.OpenXml;

namespace DocumentsGenerator.Utils
{
    internal static class OpenXmlHelper
    {
        public static OpenXmlElement? FindParent<T>(OpenXmlElement? element, int levels = 100)
        {
            if (element?.Parent == null || levels == 0)
                return null;

            if (element.Parent is T)
                return element.Parent;

            return FindParent<T>(element.Parent, --levels);
        }

        public static OpenXmlElement? FindChild<T>(OpenXmlElement? parent)
        {
            if (parent == null)
                return null;

            foreach (var element in parent)
            {
                if (element is T)
                {
                    return element;
                }
                else
                {
                    var retVal = FindChild<T>(element);
                    if (retVal != null)
                        return retVal;
                }
            }

            return null;
        }
    }
}
