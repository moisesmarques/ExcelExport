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
    public class Sheet: ExcelElementWriter
    {
        //sheet.XDocument.Descendants().First(descendant => descendant.Name.LocalName == "sheetData")
        public SheetData SheetData { get; set; }
        
        public Sheet(ZipArchiveEntry entry, int sheetNumber)
        {
            
            ZipArchiveEntry = entry;
            StreamWriter = new StreamWriter(ZipArchiveEntry.Open());
            XDocument = XDocument.Load(StreamWriter.BaseStream);

            SheetData = new SheetData(XDocument.Descendants().First(descendant => descendant.Name.LocalName == "sheetData"));
        }

        public static ExcelElementWriter Obter(ZipArchive archive, int sheetNumber)
        {
            ExcelElementWriter sheet = new ExcelElementWriter();
            sheet.ZipArchiveEntry = archive.GetEntry(string.Format("xl/worksheets/sheet{0}.xml", sheetNumber));
            sheet.StreamWriter = new StreamWriter(sheet.ZipArchiveEntry.Open());
            sheet.XDocument = XDocument.Load(sheet.StreamWriter.BaseStream);
            return sheet;
        }
    }
}
