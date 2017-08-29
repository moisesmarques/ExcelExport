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
    public class SharedStrings : ExcelElementWriter
    {
        public static readonly string XmlDeclaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

        public SharedStrings(ZipArchive archive)
        {
            this.ZipArchiveEntry = archive.GetEntry("xl/sharedStrings.xml");

            if (this.ZipArchiveEntry != null)
            {                
                this.StreamWriter = new StreamWriter(this.ZipArchiveEntry.Open());
                this.XDocument = XDocument.Load(this.StreamWriter.BaseStream);
            }
            else
            {
                this.ZipArchiveEntry = archive.CreateEntry("xl/sharedStrings.xml");
                this.StreamWriter = new StreamWriter(this.ZipArchiveEntry.Open());
                this.XDocument = CriarSst();
                EditarWorkbookXmlRels(archive);
                EditarContentTypes(archive);
            }
            
        }

        public Sst Sst {
            get {
                return new Sst(this.XDocument.Descendants().FirstOrDefault(descendant => descendant.Name.LocalName == "sst"));
            } set { } }


        public static ExcelElementWriter Obter(ZipArchive archive)
        {
            var sharedStrings = new ExcelElementWriter();

            sharedStrings.ZipArchiveEntry = archive.GetEntry("xl/sharedStrings.xml");

            if (sharedStrings.ZipArchiveEntry == null)
            {
                sharedStrings.StreamWriter = new StreamWriter(archive.CreateEntry("xl/sharedStrings.xml").Open());
                sharedStrings.XDocument = CriarSst();
                EditarWorkbookXmlRels(archive);
                EditarContentTypes(archive);
            }
            else
            {
                sharedStrings.StreamWriter = new StreamWriter(sharedStrings.ZipArchiveEntry.Open());
                sharedStrings.XDocument = XDocument.Load(sharedStrings.StreamWriter.BaseStream);
            }

            return sharedStrings;
        }
        
        public static XElement ObterSst(ExcelElementWriter sharedStrings)
        {
            return sharedStrings.XDocument.Descendants().FirstOrDefault(descendant => descendant.Name.LocalName == "sst");
        }

        public static XDocument CriarSst()
        {
            XNamespace sstNameSpace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            var sstDoc = new XDocument();
            var sst = new XElement(sstNameSpace + "sst");

            sstDoc.Add(sst);
            sst.SetAttributeValue("count", "0");
            sst.SetAttributeValue("uniqueCount", "0");

            return sstDoc;
        }

        public static int ObterIndice(XElement sst)
        {
            return sst.Descendants().Count(descendant => descendant.Name.LocalName == "si");
        }

        public static XElement ObterSharedStringsType(XNamespace nameSpace)
        {
            XElement over = new XElement(nameSpace + "Override");
            over.SetAttributeValue("PartName", "/xl/sharedStrings.xml");
            over.SetAttributeValue("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
            return over;
        }

        public static XElement Inserir(string valor, XElement sst, XNamespace nameSpace)
        {
            var si = new XElement(nameSpace + "si");
            var t = new XElement(nameSpace + "t");
            t.Value = valor;
            si.Add(t);
            sst.Add(si);

            return sst;
        }

        private static void EditarWorkbookXmlRels(ZipArchive archive)
        {
            ExcelElementWriter workbookXmlRels = WorkbookXmlRels.Obter(archive);

            XElement relationships = WorkbookXmlRels.ObterRelationships(workbookXmlRels);

            WorkbookXmlRels.AdicionarRelacao(relationships, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml");

            workbookXmlRels.ConcluirEdicao(XmlDeclaration, false);
        }

        private static void EditarContentTypes(ZipArchive archive)
        {
            ExcelElementWriter contentTypes = ContentTypes.ObterContentTypes(archive);

            XElement types = ContentTypes.ObterTypes(contentTypes);

            AdicionarTypeSharedStrings(types);

            contentTypes.ConcluirEdicao(XmlDeclaration, false);
        }

        private static void AdicionarTypeSharedStrings(XElement types)
        {
            XNamespace nameSpace = types.Name.Namespace;
            XElement over = SharedStrings.ObterSharedStringsType(nameSpace);
            types.Add(over);
        }
    }
}
