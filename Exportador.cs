using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelExport
{
    public static class Exportador
    {
        public static string[] TiposParaInserirEmSst = new string[] { "String", "DateTime" };

        public static string[] Alfabeto = new string[] {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
            "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ" };

        public static byte[] Exportar(DataTable dados,
            string caminhoNomeDoArquivoTemplate,
            string[] colunasParaImprimir,
            string[] colunasCabecalho)
        {
            MemoryStream template = ObterCopiaTemplate(caminhoNomeDoArquivoTemplate);

            return Exportar(template, dados, colunasParaImprimir, colunasCabecalho, 1);
        }

        public static byte[] Exportar(DataTable dados,
           string caminhoNomeDoArquivoTemplate,
           string[] colunasParaImprimir,
           string[] colunasCabecalho,
            int numeroPlanilha)
        {
            MemoryStream template = ObterCopiaTemplate(caminhoNomeDoArquivoTemplate);

            return Exportar(template, dados, colunasParaImprimir, colunasCabecalho, numeroPlanilha);
        }

        public static byte[] Exportar(byte[] template,
            DataTable dados,
            string[] colunasParaImprimir,
            string[] colunasCabecalho,
            int numeroPlanilha)
        {
            MemoryStream templateStream = ObterCopiaTemplate(template);

            return Exportar(templateStream, dados, colunasParaImprimir, colunasCabecalho, numeroPlanilha);
        }

        /// <summary> 
        /// Exporta um DataTable para planilha excel.
        /// </summary>
        /// <param name="template"></param>
        /// <param name="dados"></param>
        /// <param name="colunasParaImprimir"></param>
        /// <param name="colunasCabecalho"></param>
        /// <param name="numeroPlanilha"></param>        
        /// <returns></returns>
        public static byte[] Exportar(MemoryStream template,
            DataTable dados,
            string[] colunasParaImprimir,
            string[] colunasCabecalho,
            int numeroPlanilha)
        {
            ZipArchive archive = new ZipArchive(template, ZipArchiveMode.Update, true);

            ExcelArchive excel = new ExcelArchive(archive);

            string[] colunasQueUsamSst = dados.Columns.Cast<DataColumn>()
                .Where(dataColumn =>
                    TiposParaInserirEmSst.Contains(dataColumn.DataType.Name.ToString())
                    && colunasParaImprimir.Contains(dataColumn.ColumnName))
                    .Select(dataColumn => dataColumn.ColumnName).ToArray();

            EditarSharedStrings(excel.SharedStrings, dados, colunasCabecalho, colunasQueUsamSst);

            Sheet sheet = excel.Sheets[0];

            int sharedStringsIndice = 0;

            EditarSheet(sheet, dados, colunasCabecalho, colunasParaImprimir, colunasQueUsamSst, sharedStringsIndice);

            archive.Dispose();

            return template.ToArray();
        }

        private static int EditarSheet(Sheet sheet,
            DataTable dados,
            string[] colunasCabecalho,
            string[] colunas,
            string[] colunasQueUsamSst,
            int sharedStringsIndice)
        {

            SheetData sheetData = sheet.SheetData;

            sharedStringsIndice = InserirCabecalho(colunasCabecalho, sharedStringsIndice, sheetData);

            foreach (var dataRow in dados.Rows.Cast<DataRow>())
                sharedStringsIndice = InserirLinha(colunas, colunasQueUsamSst, sharedStringsIndice, sheetData, dataRow);

            sheet.ConcluirEdicao();

            return sharedStringsIndice;
        }

        private static int InserirLinha(string[] colunas, string[] colunasQueUsamSst, int sharedStringsIndice, SheetData sheetData, DataRow dataRow)
        {
            var novaRow = new Row();
            novaRow.r = sheetData.Rows.Count + 1;

            int letterIndex = 0;

            if (colunas != null)
                foreach (var coluna in colunas)
                    if (colunasQueUsamSst.Contains(coluna))
                        sharedStringsIndice = InserirCelulaValorSst(sharedStringsIndice, novaRow, ref letterIndex);
                    else
                        letterIndex = InserirCelula(dataRow, novaRow, letterIndex, coluna);

            sheetData.Rows.Add(novaRow);

            return sharedStringsIndice;
        }

        private static int InserirCelula(DataRow dataRow, Row novaRow, int letterIndex, string coluna)
        {
            novaRow.Cs.Add(new Cell
            {
                r = Alfabeto[letterIndex++].ToString() + novaRow.r,
                Valor = dataRow[coluna].ToString().Replace(",", ".")
            });

            return letterIndex;
        }

        private static int InserirCelulaValorSst(int sharedStringsIndice, Row novaRow, ref int letterIndex)
        {
            novaRow.Cs.Add(new Cell
            {
                r = Alfabeto[letterIndex++].ToString() + novaRow.r,
                t = "s",
                Valor = (sharedStringsIndice++).ToString()
            });

            return sharedStringsIndice;
        }

        private static int InserirCabecalho(string[] colunasCabecalho, int sharedStringsIndice, SheetData sheetData)
        {
            var newRow = new Row();
            newRow.r = sheetData.Rows.Count + 1;

            int letterIndex = 0;

            if (colunasCabecalho != null)
                foreach (var coluna in colunasCabecalho)
                    newRow.Cs.Add(new Cell
                    {
                        r = Alfabeto[letterIndex++].ToString() + newRow.r,
                        t = "s",
                        Valor = (sharedStringsIndice++).ToString()
                    });

            sheetData.Rows.Add(newRow);
            return sharedStringsIndice;
        }

        private static MemoryStream ObterCopiaTemplate(string caminhoNomeDoArquivo)
        {
            byte[] bytes = File.ReadAllBytes(caminhoNomeDoArquivo);

            return ObterCopiaTemplate(bytes);
        }

        private static MemoryStream ObterCopiaTemplate(byte[] template)
        {
            var memoryStreamInput = new MemoryStream(template, true);

            var memoryStream = new MemoryStream();

            memoryStreamInput.CopyTo(memoryStream);

            memoryStreamInput.Dispose();

            return memoryStream;
        }

        private static void EditarSharedStrings(SharedStrings sharedStrings, DataTable dados, string[] colunasCabecalho, string[] colunasQueUsamSst)
        {
            Sst sst = sharedStrings.Sst;

            if (colunasCabecalho != null)
                foreach (var coluna in colunasCabecalho)
                    sst.Sis.Add(new Si { Valor = coluna });

            dados.Rows.Cast<DataRow>().Select(dataRow =>
                    InserirValoresSst(sst, colunasQueUsamSst, dataRow)).ToArray();

            sst.RefreshIndex();

            sharedStrings.ConcluirEdicao();
        }

        private static Sst InserirValoresSst(Sst sst, string[] colunas, DataRow dataRow)
        {
            foreach (var coluna in colunas)
                sst.Sis.Add(new Si() { Valor = dataRow[coluna].ToString() });

            return sst;
        }


    }
}
