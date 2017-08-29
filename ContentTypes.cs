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
    public class ContentTypes
    {
        public static ExcelElementWriter ObterContentTypes(ZipArchive archive)
        {
            ExcelElementWriter contentTypes = new ExcelElementWriter();

            contentTypes.ZipArchiveEntry = archive.GetEntry("[Content_Types].xml");

            contentTypes.StreamWriter = new StreamWriter(contentTypes.ZipArchiveEntry.Open());

            contentTypes.XDocument = XDocument.Load(contentTypes.StreamWriter.BaseStream);

            return contentTypes;
        }

        public static XElement ObterTypes(ExcelElementWriter contentTypes)
        {
            return contentTypes.XDocument.Descendants().FirstOrDefault(descendant => descendant.Name.LocalName == "Types");
        }
        
    }
}
