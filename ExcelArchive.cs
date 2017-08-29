using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelExport
{
    public class ExcelArchive
    {
        private readonly ZipArchive _archive;

        public ExcelArchive(ZipArchive archive)
        {
            this._archive = archive;
            this.SharedStrings =  new SharedStrings(_archive);
            this.Sheets = new ListSheet(_archive);
        }

        //{[Content_Types].xml}
        public ContentTypes ContentTypes { get; set; }

        //{_rels/.rels}
        private Object _rels { get; set; }

        //{xl/_rels/workbook.xml.rels}
        public WorkbookXmlRels WorkbookXmlRels { get; set; }

        //{xl/workbook.xml}
        public Object Workbook { get; set; }

        //{xl/styles.xml}
        public Object Styles { get; set; }

        //{xl/theme/theme1.xml}
        public List<Object> Themes { get; set; }

        //{xl/worksheets/sheet1.xml}
        public ListSheet Sheets { get; set; }

        //{xl/sharedStrings.xml}
        public SharedStrings SharedStrings { get; set; }

        //{docProps/core.xml}
        private Object Core { get; set; }

        //{docProps/app.xml}
        private Object App { get; set; }

    }
}
