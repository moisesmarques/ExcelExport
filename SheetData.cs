using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelExport
{
    public class SheetData
    {
        public XNamespace Namespace { get; set; }

        private XElement _sheetData { get; set; }

        public SheetData(XElement sheetData)
        {
            _sheetData = sheetData;
            Namespace = sheetData.Name.Namespace;
            Rows = new ListRow(sheetData);
        }

        public ListRow Rows { get; set; }
    }
}
