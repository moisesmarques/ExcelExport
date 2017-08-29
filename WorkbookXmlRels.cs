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
    public class WorkbookXmlRels
    {
        public static ExcelElementWriter Obter(ZipArchive archive)
        {
            ExcelElementWriter workbookXmlRels = new ExcelElementWriter();

            workbookXmlRels.ZipArchiveEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");

            workbookXmlRels.StreamWriter = new StreamWriter(workbookXmlRels.ZipArchiveEntry.Open());

            workbookXmlRels.XDocument = XDocument.Load(workbookXmlRels.StreamWriter.BaseStream);

            return workbookXmlRels;
        }

        public static XElement ObterRelationships(ExcelElementWriter workbookXmlRels)
        {
            return workbookXmlRels.XDocument.Descendants().FirstOrDefault(descendant => descendant.Name.LocalName == "Relationships");
        }

        public static void AdicionarRelacao(XElement relationships, string type, string target)
        {
            XNamespace nameSpace = relationships.Name.Namespace;

            int relacoes = relationships.DescendantNodes().Count();

            XElement relationship = new XElement(nameSpace + "Relationship");

            relationship.SetAttributeValue("Id", string.Concat("rId", relacoes + 1));

            relationship.SetAttributeValue("Type", type);

            relationship.SetAttributeValue("Target", target);

            relationships.Add(relationship);
        }

    }
}
