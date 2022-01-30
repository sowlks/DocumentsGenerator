using System.Text;
using DocumentsGenerator.Core;

namespace DocumentsGenerator.Word.Tags
{
    internal class DocumentTag : WordTag
    {
        public DocumentTag()
            : base(null)
        {
        }

        protected override void SetSource(SourceData sourceData)
        {
            base.SetSource(sourceData);
        }

        protected override void ChangeTemplate(SourceData sourceData)
        {
            base.ChangeTemplate(sourceData);

            foreach (ITag tag in this)
            {
                tag.Process(sourceData);
            }
        }

        public string? Correct()
        {
            var errors = Correct(this);

            return errors.Length > 0 ? errors.ToString() : null;
        }

        private StringBuilder Correct(ITag tag)
        {
            var retVal = new StringBuilder();

            if (tag.IsContainer && tag != this && tag.ParentCloseXmlElement == null)
            {
                retVal.AppendLine($"Close tag for tag \"{tag.Text}\" not found.");
            }

            foreach (ITag childTag in tag)
            {
                var sb = Correct(childTag);
                if (sb.Length > 0)
                    retVal.AppendLine(sb.ToString());
            }

            return retVal;
        }
    }
}
