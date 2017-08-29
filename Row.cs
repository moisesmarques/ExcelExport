using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelExport
{
    public class Row
    {
        public Row()
        {
            Cs = new List<Cell>();
        }

        public Row(XElement element)
        {
            var t = element.Descendants().FirstOrDefault();
            Cs = new List<Cell>();

        }

        public int r { get; set; }

        public List<Cell> Cs { get; set; }
    }
}
