using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace BotGestaoDefeitos
{
    public class GestaoDefeitos
    {
        private readonly string _path;
        private List<Tuple<int, string, string, Dictionary<string, int>>> _itensFiles;
        public GestaoDefeitos()
        {
            _path = ConfigurationManager.AppSettings["path"];
        }
        public void ExecutarGestaoDefeitos()
        {
            ListaDocumentos();
        }
        public void ListaDocumentos()
        {
            try
            {

                var pathFileSource = Directory.GetFiles(_path, "*.*", SearchOption.AllDirectories).ToList();
                _itensFiles = new List<Tuple<int, string, string, Dictionary<string, int>>>();
                foreach (string path in pathFileSource)
                {
                    string[] pathFilePart = path.Split('\\');
                    string fileName = pathFilePart[pathFilePart.Length - 1];
                    string type = fileName.Split('_').Last().Substring(0, fileName.Split('_').Last().IndexOf("."));
                    if (fileName.StartsWith("Histórico"))
                    {
                        if (fileName.Contains("Geral"))
                            _itensFiles.Add(new Tuple<int, string, string, Dictionary<string, int>>(3, path, type, LayoutExcel(type)));
                        else
                            _itensFiles.Add(new Tuple<int, string, string, Dictionary<string, int>>(1, path, type, LayoutExcel(type)));
                    }
                    if (fileName.StartsWith("Defeitos"))
                    {
                        if (fileName.Contains("Geral"))
                            _itensFiles.Add(new Tuple<int, string, string, Dictionary<string, int>>(4, path, type, LayoutExcel(type)));
                        else
                            _itensFiles.Add(new Tuple<int, string, string, Dictionary<string, int>>(2, path, type, LayoutExcel(type)));
                    }
                }
                foreach (var item in _itensFiles.Where(x => x.Item1 == 1))
                {
                    LeArquivo(item.Item2, item.Item4, item.Item3);
                }

            }
            catch (Exception ex)
            {
                log4net.LogManager.GetLogger("Processamento.Geral.Erro").Error($"Falha ao realizar gestão de defeitos.", ex);
            }
        }

        private void LeArquivo(string path, Dictionary<string, int> layout, string type)
        {

            var listBueiros = new List<Bueiro>();
            List<IGrouping<string, Bueiro>> itensEmail = new List<IGrouping<string, Bueiro>>();
            List<int> itensRemover = new List<int>();


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var pacote = new ExcelPackage(new FileInfo(path)))
            {
                var planilha = pacote.Workbook.Worksheets[0]; // Obtém a primeira planilha

                int totalLinhas = planilha.Dimension.Rows;
                int totalColunas = planilha.Dimension.Columns;

                switch (type)
                {
                    case "Bueiros":
                        listBueiros = LeArquivoBueiro(totalLinhas, planilha, layout);
                        VerificaRepetidosBueiros(listBueiros, ref itensEmail, ref itensRemover);
                        RemoveItens(itensRemover.OrderByDescending(x=> x).ToList(), planilha, pacote);
                        var itensFinal = listBueiros.Where(x => !itensRemover.Contains(x.linha)).ToList();
                        GravaItensBueiros(itensFinal);
                        break;
                }
            }
            return null;
        }


        private void RemoveItens(List<int> itensRemover, ExcelWorksheet planilha, ExcelPackage pacote)
        {
            foreach (var rep in itensRemover)
            {

                if (rep > 0 && rep <= planilha.Dimension.Rows)
                {
                    planilha.DeleteRow(rep);
                }


            }
            pacote.Save(); // Salva as alterações no arquivo original
        }

        #region Bueiros
        private void GravaItensBueiros(List<Bueiro> itensFinal)
        {
            //TODO:

            string caminhoArquivo = "caminho/do/seu/arquivo.xlsx";

            // Configura a licença do EPPlus (obrigatório desde a versão 5)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var pacote = new ExcelPackage(new FileInfo(caminhoArquivo)))
            {
                var planilha = pacote.Workbook.Worksheets[0]; // Obtém a primeira planilha

                int linhaParaInserir = 3; // A posição onde a nova linha será inserida

                // Insere uma linha na posição desejada
                planilha.InsertRow(linhaParaInserir, 1);

                // Preenche a nova linha com valores
                planilha.Cells[linhaParaInserir, 1].Value = "Novo Dado 1";
                planilha.Cells[linhaParaInserir, 2].Value = "Novo Dado 2";
                planilha.Cells[linhaParaInserir, 3].Value = "Novo Dado 3";

                // Salva as alterações no arquivo
                pacote.Save();
            }

            Console.WriteLine("Linha inserida com sucesso!");
        }
        public List<Bueiro> LeArquivoBueiro(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            var listBueiros = new List<Bueiro>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                listBueiros.Add(new Bueiro
                {
                    linha = linha,
                    ID_REGISTRO = planilha.Cells[linha, layout[ELayoutExcelBueiro.ID_REGISTRO]].Text,
                    ID_DEFEITO = planilha.Cells[linha, layout[ELayoutExcelBueiro.ID_DEFEITO]].Text,
                    ID_RONDA = planilha.Cells[linha, layout[ELayoutExcelBueiro.ID_RONDA]].Text,
                    ATUALIZACAO = planilha.Cells[linha, layout[ELayoutExcelBueiro.ATUALIZACAO]].Text,
                    TIPO_INSPECAO = planilha.Cells[linha, layout[ELayoutExcelBueiro.TIPO_INSPECAO]].Text,
                    DATA = planilha.Cells[linha, layout[ELayoutExcelBueiro.DATA]].Text,
                    RESPONSAVEL = planilha.Cells[linha, layout[ELayoutExcelBueiro.RESPONSAVEL]].Text,
                    STATUS = planilha.Cells[linha, layout[ELayoutExcelBueiro.STATUS]].Text,
                    SUB = planilha.Cells[linha, layout[ELayoutExcelBueiro.SUB]].Text,
                    KM = planilha.Cells[linha, layout[ELayoutExcelBueiro.KM]].Text,
                    EQUIP_SUPER = planilha.Cells[linha, layout[ELayoutExcelBueiro.EQUIP_SUPER]].Text,
                    EQUIP = planilha.Cells[linha, layout[ELayoutExcelBueiro.EQUIP]].Text,
                    LOCAL = planilha.Cells[linha, layout[ELayoutExcelBueiro.LOCAL]].Text,
                    DEFEITO = planilha.Cells[linha, layout[ELayoutExcelBueiro.DEFEITO]].Text,
                    PRIORIDADE = planilha.Cells[linha, layout[ELayoutExcelBueiro.PRIORIDADE]].Text,
                    OBSERVACAO = planilha.Cells[linha, layout[ELayoutExcelBueiro.OBSERVACAO]].Text,
                    FOTOS = planilha.Cells[linha, layout[ELayoutExcelBueiro.FOTOS]].Text,
                    OS = planilha.Cells[linha, layout[ELayoutExcelBueiro.OS]].Text,
                    SUB_TRECHO = planilha.Cells[linha, layout[ELayoutExcelBueiro.SUB_TRECHO]].Text,
                    POWERAPPSID = planilha.Cells[linha, layout[ELayoutExcelBueiro.POWERAPPSID]].Text,
                    ENG = planilha.Cells[linha, layout[ELayoutExcelBueiro.ENG]].Text,
                });

            }
            return listBueiros;
        }
        private void VerificaRepetidosBueiros(List<Bueiro> listBueiros, ref List<IGrouping<string, Bueiro>> itensEmail, ref List<int> itensRemover)
        {
            var itensagrupados = listBueiros.GroupBy(x => x.ID_REGISTRO).ToList();

            var repetidos = itensagrupados.Where(x => x.Count() > 1);

            foreach (var rep in repetidos)
            {
                Bueiro bueiroInicial = null;
                foreach (var item in rep)
                {
                    if (bueiroInicial == null)
                        bueiroInicial = item;
                    else
                    {
                        if (bueiroInicial.ID_DEFEITO == item.ID_DEFEITO
                          && bueiroInicial.ID_RONDA == item.ID_RONDA
                          && bueiroInicial.TIPO_INSPECAO == item.TIPO_INSPECAO
                          && bueiroInicial.DATA == item.DATA
                          && bueiroInicial.RESPONSAVEL == item.RESPONSAVEL
                          && bueiroInicial.STATUS == item.STATUS)
                            itensRemover.Add(item.linha);
                        else
                        {
                            if (!itensEmail.Contains(rep))
                                itensEmail.Add(rep);
                        }
                    }
                }
            }

        }
        #endregion

        private Dictionary<string, int> LayoutExcel(string type)
        {
            switch (type)
            {
                case "Bueiros":
                    return new Dictionary<string, int>
                                {
                                    {ELayoutExcelBueiro.ID_REGISTRO,1 },
                                    {ELayoutExcelBueiro.ID_DEFEITO, 2},
                                    {ELayoutExcelBueiro.ID_RONDA, 3 },
                                    {ELayoutExcelBueiro.ATUALIZACAO,4 },
                                    {ELayoutExcelBueiro.TIPO_INSPECAO, 5 },
                                    {ELayoutExcelBueiro.DATA, 6 },
                                    {ELayoutExcelBueiro.RESPONSAVEL, 7 },
                                    {ELayoutExcelBueiro.STATUS,8},
                                    {ELayoutExcelBueiro.SUB,9 },
                                    {ELayoutExcelBueiro.KM, 10 },
                                    {ELayoutExcelBueiro.EQUIP_SUPER,11 },
                                    {ELayoutExcelBueiro.EQUIP, 12 },
                                    {ELayoutExcelBueiro.LOCAL, 13 },
                                    {ELayoutExcelBueiro.DEFEITO,14 },
                                    {ELayoutExcelBueiro.PRIORIDADE,15 },
                                    {ELayoutExcelBueiro.OBSERVACAO, 16 },
                                    {ELayoutExcelBueiro.FOTOS, 17 },
                                    {ELayoutExcelBueiro.OS, 18 },
                                    {ELayoutExcelBueiro.SUB_TRECHO, 19 },
                                    {ELayoutExcelBueiro.POWERAPPSID,20 },
                                    {ELayoutExcelBueiro.ENG, 21 },
                                };
            }
            return null;
        }
    }
}
