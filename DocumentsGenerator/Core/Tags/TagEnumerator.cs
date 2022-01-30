using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace DocumentsGenerator.Core.Tags
{
    internal class TagEnumerator : IEnumerator<ITag>
    {
        private int currentIndex = -1;

        private IEnumerable<ITag>? tags = null;

        private bool disposed = false;

        public TagEnumerator(IEnumerable<ITag> tags)
        {
            this.tags = tags;
        }

        private ITag? current;
        public ITag Current
        {
            get
            {
                if (tags == null || current == null)
                {
                    throw new InvalidOperationException();
                }

                return current;
            }
        }

        object IEnumerator.Current => Current;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    tags = null;
                }

                current = null;
            }

            this.disposed = true;
        }

        public bool MoveNext()
        {
            currentIndex++;

            if (currentIndex == tags.Count())
                return false;

            current = tags.ElementAt(currentIndex);

            return true;
        }

        ~TagEnumerator()
        {
            Dispose(false);
        }

        public void Reset()
        {
            currentIndex = -1;
            current = null;
        }
    }
}
