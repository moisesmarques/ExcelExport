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
    public class ExcelElementWriter
    {
        public ZipArchiveEntry ZipArchiveEntry { get; set; }
        public StreamWriter StreamWriter { get; set; }
        public XDocument XDocument { get; set; }

        public void ConcluirEdicao()
        {
            ConcluirEdicao(string.Empty, true);
        }

        public void ConcluirEdicao(string declaracaoXml, bool formatarXml)
        {
            StreamWriter.BaseStream.Position = 0;

            string edicao = formatarXml ? XDocument.ToString() : XDocument.ToString(SaveOptions.DisableFormatting);
            edicao = edicao.Replace(" />", "/>");

            StreamWriter.Write(declaracaoXml != string.Empty ? string.Concat(declaracaoXml, "\r\n", edicao) : edicao);

            StreamWriter.Close();
        }

    }
}
