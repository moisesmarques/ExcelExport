using System;
using System.Collections;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport
{
    public class ListSheet : IList<Sheet>
    {

        private List<ZipArchiveEntry> _sheets { get; set; }

        public ListSheet(ZipArchive _archive)
        {
            _sheets = new List<ZipArchiveEntry>();

            bool hasEntries = true;
            int sheetNumber = 1;

            while (hasEntries)
            {
                ZipArchiveEntry entry = _archive.GetEntry(string.Format("xl/worksheets/sheet{0}.xml", sheetNumber));
                if (entry != null)
                {
                    _sheets.Add(entry);
                    hasEntries = false;
                }

                sheetNumber++;

            }
        }

        public Sheet this[int index]
        {
            get
            {
                return new Sheet(_sheets[index], index + 1);
            }
            set => throw new NotImplementedException();
        }

        public int Count => throw new NotImplementedException();

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(Sheet item)
        {
            throw new NotImplementedException();
        }

        public void Clear()
        {
            throw new NotImplementedException();
        }

        public bool Contains(Sheet item)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(Sheet[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<Sheet> GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public int IndexOf(Sheet item)
        {
            throw new NotImplementedException();
        }

        public void Insert(int index, Sheet item)
        {
            throw new NotImplementedException();
        }

        public bool Remove(Sheet item)
        {
            throw new NotImplementedException();
        }

        public void RemoveAt(int index)
        {
            throw new NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }
}
