using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelExport
{
    public class Si
    {
        public Si()
        { }

        public Si(XElement element)
        {     
            var t = element.Descendants().FirstOrDefault();
            Valor = t != null ? t.Value : string.Empty;
        }       

        public string Valor { get; set; }        
    }
}
