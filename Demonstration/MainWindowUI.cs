using System.Collections.Generic;

namespace Demonstration
{
    internal class MainWindowUI
    {
        public IEnumerable<Sample> Samples { get; private set; }

        public MainWindowUI()
        {
            Samples = new[]
            {
                new Sample("Statement of materials", "Use tags: RowSpan, GroupRows, Sum, OneStr", "Template1")
            };
        }
    }
}
