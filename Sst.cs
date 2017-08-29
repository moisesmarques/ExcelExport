using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelExport
{
    public class Sst
    {
        private XElement sst;

        public Sst(XElement xElement)
        {
            sst = xElement;
            Sis = new ListSi(sst);
        }

        public XNamespace Namespace { get { return sst.Name.Namespace; } }

        public ListSi Sis { get; set; }

        public void RefreshIndex()
        {   
            sst.SetAttributeValue("count", Sis.Count);
            sst.SetAttributeValue("uniqueCount", Sis.Count);
        }
    }
}
