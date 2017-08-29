using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelExport
{
    public class ListSi : IList<Si>
    {
        private XElement _sst;
        public ListSi(XElement sst)
        {
            _sst = sst;            
        }

        private Si[] _si { get { return _sst.Descendants().Where(descendant => descendant.Name.LocalName == "si").Select(si => new Si(si)).ToArray(); } }

        public Si this[int index] {
            get { return _si[index]; }
            set { throw new NotImplementedException(); }
        }

        public int Count { get { return _si.Length; } }

        public bool IsReadOnly { get { return false; } }

        public void Add(Si item)
        {
            XNamespace nameSpace = _sst.Name.Namespace;
            var si = new XElement(nameSpace + "si");
            var t = new XElement(nameSpace + "t");
            t.Value = item.Valor;
            si.Add(t);
            _sst.Add(si);            
        }

        public void Clear()
        {
            throw new NotImplementedException();
        }

        public bool Contains(Si item)
        {
            return _si.Contains(item);
        }

        public void CopyTo(Si[] array, int arrayIndex)
        {
            _si.CopyTo(array, arrayIndex);
        }

        public IEnumerator<Si> GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public int IndexOf(Si item)
        {
            throw new NotImplementedException();
        }

        public void Insert(int index, Si item)
        {
            throw new NotImplementedException();
        }

        public bool Remove(Si item)
        {
            throw new NotImplementedException();
        }

        public void RemoveAt(int index)
        {
            throw new NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _si.GetEnumerator();
        }
    }
}
