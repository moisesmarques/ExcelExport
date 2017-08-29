using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelExport
{
    public class ListRow : IList<Row>
    {
        private XElement _sheetData;

        public ListRow(XElement sheetData)
        {
            _sheetData = sheetData;
        }

        private Row[] _rows
        {
            get
            {
                return _sheetData.Descendants().Where(descentant => descentant.Name.LocalName == "row").Select(row => new Row(row)).ToArray();
            }
        }

        public Row this[int index]
        {
            get { return _rows[index]; }
            set => throw new NotImplementedException();
        }

        public int Count { get { return _rows.Length; } }

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(Row item)
        {
            XNamespace nameSpace = _sheetData.Name.Namespace;

            var row = new XElement(nameSpace + "row");
            row.SetAttributeValue("r", item.r.ToString());

            foreach (var c in item.Cs)
            {
                XElement xc = new XElement(nameSpace + "c");
                xc.SetAttributeValue("r", c.r);
                if (!string.IsNullOrWhiteSpace(c.t))
                    xc.SetAttributeValue("t", c.t);
                XElement xv = new XElement(nameSpace + "v");
                xv.Value = c.Valor;
                xc.Add(xv);
                row.Add(xc);
            }

            _sheetData.Add(row);
        }

        public void Clear()
        {
            throw new NotImplementedException();
        }

        public bool Contains(Row item)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(Row[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<Row> GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public int IndexOf(Row item)
        {
            throw new NotImplementedException();
        }

        public void Insert(int index, Row item)
        {
            throw new NotImplementedException();
        }

        public bool Remove(Row item)
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
