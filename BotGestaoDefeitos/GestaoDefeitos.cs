using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Windows.Media.Media3D;
using System.Windows.Media;

namespace BotGestaoDefeitos
{
    public class GestaoDefeitos
    {
        private readonly string _path;
        private readonly string _pathaux;
        private readonly string _user;
        private readonly string _password;
        private readonly string _host;
        private readonly int _port;
        private readonly string _remetente;
        private readonly string _destinatario;
        private List<Tuple<int, string, string, Dictionary<string, int>>> _itensFiles;
        public GestaoDefeitos()
        {
            _path = ConfigurationManager.AppSettings["path"];
            _pathaux = ConfigurationManager.AppSettings["pathaux"];
            _user = ConfigurationManager.AppSettings["user"];
            _password = ConfigurationManager.AppSettings["password"];
            _host = ConfigurationManager.AppSettings["host"];
            _port = Convert.ToInt32(ConfigurationManager.AppSettings["port"]);
            _destinatario = ConfigurationManager.AppSettings["destinatario"];
            _remetente = ConfigurationManager.AppSettings["remetente"];
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
                var email = "";
                foreach (var item in _itensFiles.Where(x => x.Item1 == 1))
                {
                    email += $"<p>{item.Item3}</p>";
                    email += LeArquivo(item.Item2, item.Item4, item.Item3);
                }
                EnviarEmail("Itens processados RUMO", email, new string[] { _pathaux });

            }
            catch (Exception ex)
            {
                log4net.LogManager.GetLogger("Processamento.Geral.Erro").Error($"Falha ao realizar gestão de defeitos.", ex);
            }
        }

        private string LeArquivo(string path, Dictionary<string, int> layout, string type)
        {

            var listBueiros = new List<Bueiro>();
            var itensRemoverBueiro = new List<Bueiro>();
            var itensCopiadosBueiro = new List<Bueiro>();
            var itensEmailBueiros = new List<IGrouping<string, Bueiro>>();

            var listContencoes = new List<Contencao>();
            var itensRemoverContencao = new List<Contencao>();
            var itensCopiadosContencao = new List<Contencao>();
            var itensEmailContencoes = new List<IGrouping<string, Contencao>>();

            var listInfraestrutura = new List<Infraestrutura>();
            var itensRemoverInfraestrutura = new List<Infraestrutura>();
            var itensCopiadosInfraestrutura = new List<Infraestrutura>();
            var itensEmailInfraestrutura = new List<IGrouping<string, Infraestrutura>>();

            var listPN = new List<PN>();
            var itensRemoverPN = new List<PN>();
            var itensCopiadosPN = new List<PN>();
            var itensEmailPN = new List<IGrouping<string, PN>>();

            var listTunel = new List<Tunel>();
            var itensRemoverTunel = new List<Tunel>();
            var itensCopiadosTunel = new List<Tunel>();
            var itensEmailTunel = new List<IGrouping<string, Tunel>>();

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
                        VerificaRepetidosBueiros(listBueiros, ref itensEmailBueiros, ref itensRemoverBueiro);
                        RemoveItens(itensRemoverBueiro.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);
                        var itensBueirosFinal = listBueiros.Where(x => !itensRemoverBueiro.Select(y => y.linha).Contains(x.linha)).ToList();
                        GravaItensBueiros(itensBueirosFinal);
                        GravaArquivoBueiros(itensEmailBueiros, itensRemoverBueiro, layout);
                        return MontaLayoutEmail(itensEmailBueiros, itensRemoverBueiro, itensCopiadosBueiro);
                    case "Contenções":
                        listContencoes = LeArquivoContencao(totalLinhas, planilha, layout);
                        VerificaRepetidosContencoes(listContencoes, ref itensEmailContencoes, ref itensRemoverContencao);
                        RemoveItens(itensRemoverContencao.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);
                        var itensContacoesFinal = listContencoes.Where(x => !itensRemoverContencao.Select(y => y.linha).Contains(x.linha)).ToList();
                        GravaItensContencoes(itensContacoesFinal);
                        return MontaLayoutEmail(itensEmailContencoes, itensRemoverContencao, itensCopiadosContencao);
                    case "Infraestrutura":
                        listInfraestrutura = LeArquivoInfraestrutura(totalLinhas, planilha, layout);
                        VerificaRepetidosInfraestrutura(listInfraestrutura, ref itensEmailInfraestrutura, ref itensRemoverInfraestrutura);
                        RemoveItens(itensRemoverInfraestrutura.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);
                        var itensInfraestruturaFinal = listInfraestrutura.Where(x => !itensRemoverInfraestrutura.Select(y => y.linha).Contains(x.linha)).ToList();
                        GravaItensInfraestrutura(itensInfraestruturaFinal);
                        return MontaLayoutEmail(itensEmailInfraestrutura, itensRemoverInfraestrutura, itensCopiadosInfraestrutura);
                    case "PN":
                        listPN = LeArquivoPN(totalLinhas, planilha, layout);
                        VerificaRepetidosPN(listPN, ref itensEmailPN, ref itensRemoverPN);
                        RemoveItens(itensRemoverPN.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);
                        var itensPNFinal = listPN.Where(x => !itensRemoverPN.Select(y => y.linha).Contains(x.linha)).ToList();
                        GravaItensPN(itensPNFinal);
                        return MontaLayoutEmail(itensEmailPN, itensRemoverPN, itensCopiadosPN);
                    case "Túneis":
                        listTunel = LeArquivoTunel(totalLinhas, planilha, layout);
                        VerificaRepetidosTunel(listTunel, ref itensEmailTunel, ref itensRemoverTunel);
                        RemoveItens(itensRemoverTunel.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);
                        var itensTunelFinal = listTunel.Where(x => !itensRemoverTunel.Select(y => y.linha).Contains(x.linha)).ToList();
                        GravaItensTunel(itensTunelFinal);
                        return MontaLayoutEmail(itensEmailTunel, itensRemoverTunel, itensCopiadosTunel);
                }

            }
            return "";
        }


        #region Tunel
        private void GravaItensTunel(List<Tunel> itensFinal)
        {
            //TODO:

        }
        public List<Tunel> LeArquivoTunel(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            var listTunel = new List<Tunel>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                listTunel.Add(new Tunel
                {
                    linha = linha,
                    ID_REGISTRO = planilha.Cells[linha, layout[ELayoutExcelTunel.ID_REGISTRO]].Text,
                    ID_DEFEITO = planilha.Cells[linha, layout[ELayoutExcelTunel.ID_DEFEITO]].Text,
                    ID_RONDA = planilha.Cells[linha, layout[ELayoutExcelTunel.ID_RONDA]].Text,
                    ATUALIZACAO = planilha.Cells[linha, layout[ELayoutExcelTunel.ATUALIZACAO]].Text,
                    TIPO_INSPECAO = planilha.Cells[linha, layout[ELayoutExcelTunel.TIPO_INSPECAO]].Text,
                    DATA = planilha.Cells[linha, layout[ELayoutExcelTunel.DATA]].Text,
                    RESPONSAVEL = planilha.Cells[linha, layout[ELayoutExcelTunel.RESPONSAVEL]].Text,
                    STATUS = planilha.Cells[linha, layout[ELayoutExcelTunel.STATUS]].Text,
                    SUB = planilha.Cells[linha, layout[ELayoutExcelTunel.SUB]].Text,
                    KM = planilha.Cells[linha, layout[ELayoutExcelTunel.KM]].Text,
                    EQUIP_SUPER = planilha.Cells[linha, layout[ELayoutExcelTunel.EQUIP_SUPER]].Text,
                    EQUIP = planilha.Cells[linha, layout[ELayoutExcelTunel.EQUIP]].Text,
                    KM_INICIO = planilha.Cells[linha, layout[ELayoutExcelTunel.KM_INICIO]].Text,
                    KM_FIM = planilha.Cells[linha, layout[ELayoutExcelTunel.KM_FIM]].Text,
                    EXTENSAO = planilha.Cells[linha, layout[ELayoutExcelTunel.EXTENSAO]].Text,
                    LATITUDE_INICIO = planilha.Cells[linha, layout[ELayoutExcelTunel.LATITUDE_INICIO]].Text,
                    LATITUDE_FIM = planilha.Cells[linha, layout[ELayoutExcelTunel.LATITUDE_FIM]].Text,
                    LONGITUDE_INICIO = planilha.Cells[linha, layout[ELayoutExcelTunel.LONGITUDE_INICIO]].Text,
                    LONGITUDE_FIM = planilha.Cells[linha, layout[ELayoutExcelTunel.LONGITUDE_FIM]].Text,
                    LAUDO = planilha.Cells[linha, layout[ELayoutExcelTunel.LAUDO]].Text,
                    DEFEITO = planilha.Cells[linha, layout[ELayoutExcelTunel.DEFEITO]].Text,
                    PRIORIDADE = planilha.Cells[linha, layout[ELayoutExcelTunel.PRIORIDADE]].Text,
                    OBSERVACAO = planilha.Cells[linha, layout[ELayoutExcelTunel.OBSERVACAO]].Text,
                    FOTOS = planilha.Cells[linha, layout[ELayoutExcelTunel.FOTOS]].Text,
                    OS = planilha.Cells[linha, layout[ELayoutExcelTunel.OS]].Text,
                    SUB_TRECHO = planilha.Cells[linha, layout[ELayoutExcelTunel.SUB_TRECHO]].Text,
                    POWERAPPSID = planilha.Cells[linha, layout[ELayoutExcelTunel.POWERAPPSID]].Text,
                    ENG = planilha.Cells[linha, layout[ELayoutExcelTunel.ENG]].Text,
                });

            }
            return listTunel;
        }
        private void VerificaRepetidosTunel(List<Tunel> listTunel, ref List<IGrouping<string, Tunel>> itensEmail, ref List<Tunel> itensRemover)
        {
            var itensagrupados = listTunel.Where(x => !string.IsNullOrEmpty(x.ID_REGISTRO)).GroupBy(x => x.ID_REGISTRO).ToList();

            var repetidos = itensagrupados.Where(x => x.Count() > 1);

            foreach (var rep in repetidos)
            {
                Tunel tunelInicial = null;
                foreach (var item in rep)
                {
                    if (tunelInicial == null)
                        tunelInicial = item;
                    else
                    {
                        if (tunelInicial.ID_DEFEITO == item.ID_DEFEITO
                          && tunelInicial.ID_RONDA == item.ID_RONDA
                          && tunelInicial.TIPO_INSPECAO == item.TIPO_INSPECAO
                          && tunelInicial.DATA == item.DATA
                          && tunelInicial.RESPONSAVEL == item.RESPONSAVEL
                          && tunelInicial.STATUS == item.STATUS)
                            itensRemover.Add(item);
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

        #region PN
        private void GravaItensPN(List<PN> itensFinal)
        {
            //TODO:

        }
        public List<PN> LeArquivoPN(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            var listPN = new List<PN>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                listPN.Add(new PN
                {
                    linha = linha,
                    ID_REGISTRO = planilha.Cells[linha, layout[ELayoutExcelPN.ID_REGISTRO]].Text,
                    ID_DEFEITO = planilha.Cells[linha, layout[ELayoutExcelPN.ID_DEFEITO]].Text,
                    ID_RONDA = planilha.Cells[linha, layout[ELayoutExcelPN.ID_RONDA]].Text,
                    ATUALIZACAO = planilha.Cells[linha, layout[ELayoutExcelPN.ATUALIZACAO]].Text,
                    TIPO_INSPECAO = planilha.Cells[linha, layout[ELayoutExcelPN.TIPO_INSPECAO]].Text,
                    DATA = planilha.Cells[linha, layout[ELayoutExcelPN.DATA]].Text,
                    RESPONSAVEL = planilha.Cells[linha, layout[ELayoutExcelPN.RESPONSAVEL]].Text,
                    STATUS = planilha.Cells[linha, layout[ELayoutExcelPN.STATUS]].Text,
                    SUB = planilha.Cells[linha, layout[ELayoutExcelPN.SUB]].Text,
                    KM = planilha.Cells[linha, layout[ELayoutExcelPN.KM]].Text,
                    EQUIP_SUPER = planilha.Cells[linha, layout[ELayoutExcelPN.EQUIP_SUPER]].Text,
                    EQUIP = planilha.Cells[linha, layout[ELayoutExcelPN.EQUIP]].Text,
                    KM_INICIO = planilha.Cells[linha, layout[ELayoutExcelPN.KM_INICIO]].Text,
                    KM_FIM = planilha.Cells[linha, layout[ELayoutExcelPN.KM_FIM]].Text,
                    EXTENSAO = planilha.Cells[linha, layout[ELayoutExcelPN.EXTENSAO]].Text,
                    LATITUDE_INICIO = planilha.Cells[linha, layout[ELayoutExcelPN.LATITUDE_INICIO]].Text,
                    LATITUDE_FIM = planilha.Cells[linha, layout[ELayoutExcelPN.LATITUDE_FIM]].Text,
                    LAUDO = planilha.Cells[linha, layout[ELayoutExcelPN.LAUDO]].Text,
                    DEFEITO = planilha.Cells[linha, layout[ELayoutExcelPN.DEFEITO]].Text,
                    PRIORIDADE = planilha.Cells[linha, layout[ELayoutExcelPN.PRIORIDADE]].Text,
                    OBSERVACAO = planilha.Cells[linha, layout[ELayoutExcelPN.OBSERVACAO]].Text,
                    FOTOS = planilha.Cells[linha, layout[ELayoutExcelPN.FOTOS]].Text,
                    OS = planilha.Cells[linha, layout[ELayoutExcelPN.OS]].Text,
                    SUB_TRECHO = planilha.Cells[linha, layout[ELayoutExcelPN.SUB_TRECHO]].Text,
                    POWERAPPSID = planilha.Cells[linha, layout[ELayoutExcelPN.POWERAPPSID]].Text,
                    ENG = planilha.Cells[linha, layout[ELayoutExcelPN.ENG]].Text,
                    GRADE = planilha.Cells[linha, layout[ELayoutExcelPN.GRADE]].Text,
                    TIPO_TERRENO = planilha.Cells[linha, layout[ELayoutExcelPN.TIPO_TERRENO]].Text,
                });

            }
            return listPN;
        }
        private void VerificaRepetidosPN(List<PN> listPN, ref List<IGrouping<string, PN>> itensEmail, ref List<PN> itensRemover)
        {
            var itensagrupados = listPN.Where(x => !string.IsNullOrEmpty(x.ID_REGISTRO)).GroupBy(x => x.ID_REGISTRO).ToList();

            var repetidos = itensagrupados.Where(x => x.Count() > 1);

            foreach (var rep in repetidos)
            {
                PN pNInicial = null;
                foreach (var item in rep)
                {
                    if (pNInicial == null)
                        pNInicial = item;
                    else
                    {
                        if (pNInicial.ID_DEFEITO == item.ID_DEFEITO
                          && pNInicial.ID_RONDA == item.ID_RONDA
                          && pNInicial.TIPO_INSPECAO == item.TIPO_INSPECAO
                          && pNInicial.DATA == item.DATA
                          && pNInicial.RESPONSAVEL == item.RESPONSAVEL
                          && pNInicial.STATUS == item.STATUS)
                            itensRemover.Add(item);
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

        #region Infraestrutura
        private void GravaItensInfraestrutura(List<Infraestrutura> itensFinal)
        {
            //TODO:

        }
        public List<Infraestrutura> LeArquivoInfraestrutura(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            var listInfraestrutura = new List<Infraestrutura>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                listInfraestrutura.Add(new Infraestrutura
                {
                    linha = linha,
                    ID_REGISTRO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_REGISTRO]].Text,
                    ID_DEFEITO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_DEFEITO]].Text,
                    ID_RONDA = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_RONDA]].Text,
                    ATUALIZACAO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.ATUALIZACAO]].Text,
                    TIPO_INSPECAO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.TIPO_INSPECAO]].Text,
                    DATA = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.DATA]].Text,
                    RESPONSAVEL = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.RESPONSAVEL]].Text,
                    STATUS = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.STATUS]].Text,
                    SUB = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.SUB]].Text,
                    KM = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.KM]].Text,
                    EQUIP_SUPER = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.EQUIP_SUPER]].Text,
                    EQUIP = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.EQUIP]].Text,
                    KM_INICIO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.KM_INICIO]].Text,
                    KM_FIM = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.KM_FIM]].Text,
                    EXTENSAO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.EXTENSAO]].Text,
                    LATITUDE_INICIO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.LATITUDE_INICIO]].Text,
                    LATITUDE_FIM = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.LATITUDE_FIM]].Text,
                    LONGITUDE_INICIO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.LONGITUDE_INICIO]].Text,
                    LONGITUDE_FIM = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.LONGITUDE_FIM]].Text,
                    LAUDO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.LAUDO]].Text,
                    DEFEITO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.DEFEITO]].Text,
                    PRIORIDADE = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.PRIORIDADE]].Text,
                    OBSERVACAO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.OBSERVACAO]].Text,
                    FOTOS = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.FOTOS]].Text,
                    OS = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.OS]].Text,
                    SUB_TRECHO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.SUB_TRECHO]].Text,
                    IDENT = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.IDENT]].Text,
                    POWERAPPSID = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.POWERAPPSID]].Text,
                    ENG = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.ENG]].Text,
                    TIPO_TERRENO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.TIPO_TERRENO]].Text,
                });

            }
            return listInfraestrutura;
        }
        private void VerificaRepetidosInfraestrutura(List<Infraestrutura> listInfraestrutura, ref List<IGrouping<string, Infraestrutura>> itensEmail, ref List<Infraestrutura> itensRemover)
        {
            var itensagrupados = listInfraestrutura.Where(x => !string.IsNullOrEmpty(x.ID_REGISTRO)).GroupBy(x => x.ID_REGISTRO).ToList();

            var repetidos = itensagrupados.Where(x => x.Count() > 1);

            foreach (var rep in repetidos)
            {
                Infraestrutura InfraestruturaInicial = null;
                foreach (var item in rep)
                {
                    if (InfraestruturaInicial == null)
                        InfraestruturaInicial = item;
                    else
                    {
                        if (InfraestruturaInicial.ID_DEFEITO == item.ID_DEFEITO
                          && InfraestruturaInicial.ID_RONDA == item.ID_RONDA
                          && InfraestruturaInicial.TIPO_INSPECAO == item.TIPO_INSPECAO
                          && InfraestruturaInicial.DATA == item.DATA
                          && InfraestruturaInicial.RESPONSAVEL == item.RESPONSAVEL
                          && InfraestruturaInicial.STATUS == item.STATUS)
                            itensRemover.Add(item);
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

        #region Contencoes
        private void GravaItensContencoes(List<Contencao> itensFinal)
        {
            //TODO:

        }
        public List<Contencao> LeArquivoContencao(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            var listContencoes = new List<Contencao>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                listContencoes.Add(new Contencao
                {
                    linha = linha,
                    ID_REGISTRO = planilha.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Text,
                    ID_DEFEITO = planilha.Cells[linha, layout[ELayoutExcelContencao.ID_DEFEITO]].Text,
                    ID_RONDA = planilha.Cells[linha, layout[ELayoutExcelContencao.ID_RONDA]].Text,
                    ATUALIZACAO = planilha.Cells[linha, layout[ELayoutExcelContencao.ATUALIZACAO]].Text,
                    TIPO_INSPECAO = planilha.Cells[linha, layout[ELayoutExcelContencao.TIPO_INSPECAO]].Text,
                    DATA = planilha.Cells[linha, layout[ELayoutExcelContencao.DATA]].Text,
                    RESPONSAVEL = planilha.Cells[linha, layout[ELayoutExcelContencao.RESPONSAVEL]].Text,
                    STATUS = planilha.Cells[linha, layout[ELayoutExcelContencao.STATUS]].Text,
                    SUB = planilha.Cells[linha, layout[ELayoutExcelContencao.SUB]].Text,
                    KM = planilha.Cells[linha, layout[ELayoutExcelContencao.KM]].Text,
                    EQUIP_SUPER = planilha.Cells[linha, layout[ELayoutExcelContencao.EQUIP_SUPER]].Text,
                    EQUIP = planilha.Cells[linha, layout[ELayoutExcelContencao.EQUIP]].Text,
                    KM_INICIO = planilha.Cells[linha, layout[ELayoutExcelContencao.KM_INICIO]].Text,
                    KM_FIM = planilha.Cells[linha, layout[ELayoutExcelContencao.KM_FIM]].Text,
                    EXTENSAO = planilha.Cells[linha, layout[ELayoutExcelContencao.EXTENSAO]].Text,
                    LATITUDE_INICIO = planilha.Cells[linha, layout[ELayoutExcelContencao.LATITUDE_INICIO]].Text,
                    LATITUDE_FIM = planilha.Cells[linha, layout[ELayoutExcelContencao.LATITUDE_FIM]].Text,
                    LONGITUDE_INICIO = planilha.Cells[linha, layout[ELayoutExcelContencao.LONGITUDE_INICIO]].Text,
                    LONGITUDE_FIM = planilha.Cells[linha, layout[ELayoutExcelContencao.LONGITUDE_FIM]].Text,
                    LAUDO = planilha.Cells[linha, layout[ELayoutExcelContencao.LAUDO]].Text,
                    DEFEITO = planilha.Cells[linha, layout[ELayoutExcelContencao.DEFEITO]].Text,
                    IDENT = planilha.Cells[linha, layout[ELayoutExcelContencao.IDENT]].Text,
                    PRIORIDADE = planilha.Cells[linha, layout[ELayoutExcelContencao.PRIORIDADE]].Text,
                    OBSERVACAO = planilha.Cells[linha, layout[ELayoutExcelContencao.OBSERVACAO]].Text,
                    FOTOS = planilha.Cells[linha, layout[ELayoutExcelContencao.FOTOS]].Text,
                    OS = planilha.Cells[linha, layout[ELayoutExcelContencao.OS]].Text,
                    SUB_TRECHO = planilha.Cells[linha, layout[ELayoutExcelContencao.SUB_TRECHO]].Text,
                    POWERAPPSID = planilha.Cells[linha, layout[ELayoutExcelContencao.POWERAPPSID]].Text,
                    DATA2 = planilha.Cells[linha, layout[ELayoutExcelContencao.DATA2]].Text,
                    ENG = planilha.Cells[linha, layout[ELayoutExcelContencao.ENG]].Text,
                });

            }
            return listContencoes;
        }
        private void VerificaRepetidosContencoes(List<Contencao> listContencoes, ref List<IGrouping<string, Contencao>> itensEmail, ref List<Contencao> itensRemover)
        {
            var itensagrupados = listContencoes.Where(x => !string.IsNullOrEmpty(x.ID_REGISTRO)).GroupBy(x => x.ID_REGISTRO).ToList();

            var repetidos = itensagrupados.Where(x => x.Count() > 1);

            foreach (var rep in repetidos)
            {
                Contencao contencaoInicial = null;
                foreach (var item in rep)
                {
                    if (contencaoInicial == null)
                        contencaoInicial = item;
                    else
                    {
                        if (contencaoInicial.ID_DEFEITO == item.ID_DEFEITO
                          && contencaoInicial.ID_RONDA == item.ID_RONDA
                          && contencaoInicial.TIPO_INSPECAO == item.TIPO_INSPECAO
                          && contencaoInicial.DATA == item.DATA
                          && contencaoInicial.RESPONSAVEL == item.RESPONSAVEL
                          && contencaoInicial.STATUS == item.STATUS)
                            itensRemover.Add(item);
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

        #region Bueiros
        private void GravaItensBueiros(List<Bueiro> itensFinal)
        {
            //TODO:

            //string caminhoArquivo = "caminho/do/seu/arquivo.xlsx";

            //// Configura a licença do EPPlus (obrigatório desde a versão 5)
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //using (var pacote = new ExcelPackage(new FileInfo(caminhoArquivo)))
            //{
            //    var planilha = pacote.Workbook.Worksheets[0]; // Obtém a primeira planilha

            //    int linhaParaInserir = 3; // A posição onde a nova linha será inserida

            //    // Insere uma linha na posição desejada
            //    planilha.InsertRow(linhaParaInserir, 1);

            //    // Preenche a nova linha com valores
            //    planilha.Cells[linhaParaInserir, 1].Value = "Novo Dado 1";
            //    planilha.Cells[linhaParaInserir, 2].Value = "Novo Dado 2";
            //    planilha.Cells[linhaParaInserir, 3].Value = "Novo Dado 3";

            //    // Salva as alterações no arquivo
            //    pacote.Save();
            //}

            //Console.WriteLine("Linha inserida com sucesso!");
        }
        private List<Bueiro> LeArquivoBueiro(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
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
        private void VerificaRepetidosBueiros(List<Bueiro> listBueiros, ref List<IGrouping<string, Bueiro>> itensEmail, ref List<Bueiro> itensRemover)
        {
            var itensagrupados = listBueiros.Where(x => !string.IsNullOrEmpty(x.ID_REGISTRO)).GroupBy(x => x.ID_REGISTRO).ToList();

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
                            itensRemover.Add(item);
                        else
                        {
                            if (!itensEmail.Contains(rep))
                                itensEmail.Add(rep);
                        }
                    }
                }
            }

        }
        private void GravaArquivoBueiros(List<IGrouping<string, Bueiro>> itensEmailBueiros, List<Bueiro> itensRemoverBueiro, Dictionary<string, int> layout)
        {
            // Configura a licença do EPPlus (obrigatório desde a versão 5)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var pacote = new ExcelPackage())
            {
                var planilha = pacote.Workbook.Worksheets["Bueiros_analise"];
                if (planilha != null)
                    pacote.Workbook.Worksheets.Delete("Bueiros_analise"); 
                var planilha2 = pacote.Workbook.Worksheets["Bueiros_excluidos"];
                if (planilha2 != null)
                    pacote.Workbook.Worksheets.Delete("Bueiros_excluidos");

                var planilhaAnalise = pacote.Workbook.Worksheets.Add("Bueiros_analise");

                // Preenche os cabeçalhos
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.ID_REGISTRO]].Value = "ID_Registro";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.ID_DEFEITO]].Value = "ID_Defeito";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.ID_RONDA]].Value = "ID_Ronda";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.ATUALIZACAO]].Value = "Atualização";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.TIPO_INSPECAO]].Value = "Tipo de Inspeção";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.DATA]].Value = "DATA ";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.RESPONSAVEL]].Value = "Responsável";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.STATUS]].Value = "Status Defeito";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.SUB]].Value = "SUB_Defeito_Bueiro ";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.KM]].Value = "Km_Nominal";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.EQUIP_SUPER]].Value = "Equip_Super ";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.EQUIP]].Value = "Equip_Bueiro";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.LOCAL]].Value = "Local_Defeito";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.DEFEITO]].Value = "Defeito_Bueiro";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.PRIORIDADE]].Value = "Prioridade_defeito ";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.OBSERVACAO]].Value = "Observação";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.FOTOS]].Value = "Fotos";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.OS]].Value = "OS ";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.SUB_TRECHO]].Value = "Sub_Trecho";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.POWERAPPSID]].Value = "__PowerAppsId__";
                planilhaAnalise.Cells[1, layout[ELayoutExcelBueiro.ENG]].Value = "Eng";
                int linha = 2;

                foreach (var item in itensEmailBueiros.SelectMany(x => x))
                {
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.ID_REGISTRO]].Value = item.ID_REGISTRO;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.ID_DEFEITO]].Value = item.ID_DEFEITO;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.ID_RONDA]].Value = item.ID_RONDA;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.ATUALIZACAO]].Value = item.ATUALIZACAO;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.DATA]].Value = item.DATA;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.RESPONSAVEL]].Value = item.RESPONSAVEL;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.STATUS]].Value = item.STATUS;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.SUB]].Value = item.SUB;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.KM]].Value = item.KM;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.EQUIP]].Value = item.EQUIP;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.LOCAL]].Value = item.LOCAL;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.DEFEITO]].Value = item.DEFEITO;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.PRIORIDADE]].Value = item.PRIORIDADE;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.OBSERVACAO]].Value = item.OBSERVACAO;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.FOTOS]].Value = item.FOTOS;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.OS]].Value = item.OS;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.SUB_TRECHO]].Value = item.SUB_TRECHO;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.POWERAPPSID]].Value = item.POWERAPPSID;
                    planilhaAnalise.Cells[linha, layout[ELayoutExcelBueiro.ENG]].Value = item.ENG;
                    linha++;
                }


                var planilhaExcluidos = pacote.Workbook.Worksheets.Add("Bueiros_excluidos");

                // Preenche os cabeçalhos
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.ID_REGISTRO]].Value = "ID_Registro";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.ID_DEFEITO]].Value = "ID_Defeito";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.ID_RONDA]].Value = "ID_Ronda";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.ATUALIZACAO]].Value = "Atualização";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.TIPO_INSPECAO]].Value = "Tipo de Inspeção";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.DATA]].Value = "DATA ";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.RESPONSAVEL]].Value = "Responsável";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.STATUS]].Value = "Status Defeito";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.SUB]].Value = "SUB_Defeito_Bueiro ";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.KM]].Value = "Km_Nominal";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.EQUIP_SUPER]].Value = "Equip_Super ";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.EQUIP]].Value = "Equip_Bueiro";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.LOCAL]].Value = "Local_Defeito";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.DEFEITO]].Value = "Defeito_Bueiro";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.PRIORIDADE]].Value = "Prioridade_defeito ";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.OBSERVACAO]].Value = "Observação";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.FOTOS]].Value = "Fotos";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.OS]].Value = "OS ";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.SUB_TRECHO]].Value = "Sub_Trecho";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.POWERAPPSID]].Value = "__PowerAppsId__";
                planilhaExcluidos.Cells[1, layout[ELayoutExcelBueiro.ENG]].Value = "Eng";
                linha = 2;

                foreach (var item in itensRemoverBueiro)
                {
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.ID_REGISTRO]].Value = item.ID_REGISTRO;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.ID_DEFEITO]].Value = item.ID_DEFEITO;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.ID_RONDA]].Value = item.ID_RONDA;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.ATUALIZACAO]].Value = item.ATUALIZACAO;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.DATA]].Value = item.DATA;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.RESPONSAVEL]].Value = item.RESPONSAVEL;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.STATUS]].Value = item.STATUS;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.SUB]].Value = item.SUB;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.KM]].Value = item.KM;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.EQUIP]].Value = item.EQUIP;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.LOCAL]].Value = item.LOCAL;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.DEFEITO]].Value = item.DEFEITO;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.PRIORIDADE]].Value = item.PRIORIDADE;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.OBSERVACAO]].Value = item.OBSERVACAO;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.FOTOS]].Value = item.FOTOS;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.OS]].Value = item.OS;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.SUB_TRECHO]].Value = item.SUB_TRECHO;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.POWERAPPSID]].Value = item.POWERAPPSID;
                    planilhaExcluidos.Cells[linha, layout[ELayoutExcelBueiro.ENG]].Value = item.ENG;
                    linha++;
                }

                // Salva o arquivo no disco
                File.WriteAllBytes(_pathaux, pacote.GetAsByteArray());
            }
        }
        #endregion

        private void RemoveItens(List<int> itensRemover, ExcelWorksheet planilha, ExcelPackage pacote)
        {
            if (itensRemover.Any())
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
        }

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
                case "Contenções":
                    return new Dictionary<string, int>
                                {
                                    {ELayoutExcelContencao.ID_REGISTRO,1 },
                                    {ELayoutExcelContencao.ID_DEFEITO, 2},
                                    {ELayoutExcelContencao.ID_RONDA, 3 },
                                    {ELayoutExcelContencao.ATUALIZACAO,4 },
                                    {ELayoutExcelContencao.TIPO_INSPECAO, 5 },
                                    {ELayoutExcelContencao.DATA, 6 },
                                    {ELayoutExcelContencao.RESPONSAVEL, 7 },
                                    {ELayoutExcelContencao.STATUS,8},
                                    {ELayoutExcelContencao.SUB,9 },
                                    {ELayoutExcelContencao.KM, 10 },
                                    {ELayoutExcelContencao.EQUIP_SUPER,11 },
                                    {ELayoutExcelContencao.EQUIP, 12 },
                                    {ELayoutExcelContencao.KM_INICIO, 13 },
                                    {ELayoutExcelContencao.KM_FIM, 14 },
                                    {ELayoutExcelContencao.EXTENSAO, 15 },
                                    {ELayoutExcelContencao.LATITUDE_INICIO, 16 },
                                    {ELayoutExcelContencao.LATITUDE_FIM, 17 },
                                    {ELayoutExcelContencao.LONGITUDE_INICIO, 18 },
                                    {ELayoutExcelContencao.LONGITUDE_FIM, 19 },
                                    {ELayoutExcelContencao.LAUDO, 20 },
                                    {ELayoutExcelContencao.DEFEITO,21 },
                                    {ELayoutExcelContencao.PRIORIDADE,22 },
                                    {ELayoutExcelContencao.IDENT, 23 },
                                    {ELayoutExcelContencao.OBSERVACAO, 24 },
                                    {ELayoutExcelContencao.FOTOS, 25 },
                                    {ELayoutExcelContencao.OS, 26 },
                                    {ELayoutExcelContencao.SUB_TRECHO, 27 },
                                    {ELayoutExcelContencao.POWERAPPSID,28 },
                                    {ELayoutExcelContencao.DATA2, 29 },
                                    {ELayoutExcelContencao.ENG, 30 },
                                };
                case "Infraestrutura":
                    return new Dictionary<string, int>
                                {
                                    {ELayoutExcelInfraestrutura.ID_REGISTRO,1 },
                                    {ELayoutExcelInfraestrutura.ID_DEFEITO, 2},
                                    {ELayoutExcelInfraestrutura.ID_RONDA, 3 },
                                    {ELayoutExcelInfraestrutura.ATUALIZACAO,4 },
                                    {ELayoutExcelInfraestrutura.TIPO_INSPECAO, 5 },
                                    {ELayoutExcelInfraestrutura.DATA, 6 },
                                    {ELayoutExcelInfraestrutura.RESPONSAVEL, 7 },
                                    {ELayoutExcelInfraestrutura.STATUS,8},
                                    {ELayoutExcelInfraestrutura.SUB,9 },
                                    {ELayoutExcelInfraestrutura.KM, 10 },
                                    {ELayoutExcelInfraestrutura.EQUIP_SUPER,11 },
                                    {ELayoutExcelInfraestrutura.EQUIP, 12 },
                                    {ELayoutExcelInfraestrutura.KM_INICIO, 13 },
                                    {ELayoutExcelInfraestrutura.KM_FIM, 14 },
                                    {ELayoutExcelInfraestrutura.EXTENSAO, 15 },
                                    {ELayoutExcelInfraestrutura.LATITUDE_INICIO, 16 },
                                    {ELayoutExcelInfraestrutura.LATITUDE_FIM, 17 },
                                    {ELayoutExcelInfraestrutura.LONGITUDE_INICIO, 18 },
                                    {ELayoutExcelInfraestrutura.LONGITUDE_FIM, 19 },
                                    {ELayoutExcelInfraestrutura.TIPO_TERRENO, 20 },
                                    {ELayoutExcelInfraestrutura.LAUDO, 21 },
                                    {ELayoutExcelInfraestrutura.DEFEITO,22 },
                                    {ELayoutExcelInfraestrutura.PRIORIDADE,23 },
                                    {ELayoutExcelInfraestrutura.OBSERVACAO, 24 },
                                    {ELayoutExcelInfraestrutura.FOTOS, 25 },
                                    {ELayoutExcelInfraestrutura.OS, 26 },
                                    {ELayoutExcelInfraestrutura.SUB_TRECHO, 27 },
                                    {ELayoutExcelInfraestrutura.IDENT, 28 },
                                    {ELayoutExcelInfraestrutura.POWERAPPSID,29 },
                                    {ELayoutExcelInfraestrutura.ENG, 30 },
                                };
                case "PN":
                    return new Dictionary<string, int>
                                {
                                    {ELayoutExcelPN.ID_REGISTRO,1 },
                                    {ELayoutExcelPN.ID_DEFEITO, 2},
                                    {ELayoutExcelPN.ID_RONDA, 3 },
                                    {ELayoutExcelPN.ATUALIZACAO,4 },
                                    {ELayoutExcelPN.TIPO_INSPECAO, 5 },
                                    {ELayoutExcelPN.DATA, 6 },
                                    {ELayoutExcelPN.RESPONSAVEL, 7 },
                                    {ELayoutExcelPN.STATUS,8},
                                    {ELayoutExcelPN.SUB,9 },
                                    {ELayoutExcelPN.KM, 10 },
                                    {ELayoutExcelPN.EQUIP_SUPER,11 },
                                    {ELayoutExcelPN.EQUIP, 12 },
                                    {ELayoutExcelPN.KM_INICIO, 13 },
                                    {ELayoutExcelPN.KM_FIM, 14 },
                                    {ELayoutExcelPN.EXTENSAO, 15 },
                                    {ELayoutExcelPN.LATITUDE_INICIO, 16 },
                                    {ELayoutExcelPN.LATITUDE_FIM, 17 },
                                    {ELayoutExcelPN.TIPO_TERRENO, 18 },
                                    {ELayoutExcelPN.LAUDO, 19 },
                                    {ELayoutExcelPN.DEFEITO,20 },
                                    {ELayoutExcelPN.PRIORIDADE,21 },
                                    {ELayoutExcelPN.OBSERVACAO, 22 },
                                    {ELayoutExcelPN.FOTOS, 23 },
                                    {ELayoutExcelPN.OS, 24 },
                                    {ELayoutExcelPN.SUB_TRECHO, 25 },
                                    {ELayoutExcelPN.POWERAPPSID,26 },
                                    {ELayoutExcelPN.ENG, 27 },
                                    {ELayoutExcelPN.GRADE, 28 },
                                };
                case "Túneis":
                    return new Dictionary<string, int>
                                {
                                    {ELayoutExcelTunel.ID_REGISTRO,1 },
                                    {ELayoutExcelTunel.ID_DEFEITO, 2},
                                    {ELayoutExcelTunel.ID_RONDA, 3 },
                                    {ELayoutExcelTunel.ATUALIZACAO,4 },
                                    {ELayoutExcelTunel.TIPO_INSPECAO, 5 },
                                    {ELayoutExcelTunel.DATA, 6 },
                                    {ELayoutExcelTunel.RESPONSAVEL, 7 },
                                    {ELayoutExcelTunel.STATUS,8},
                                    {ELayoutExcelTunel.SUB,9 },
                                    {ELayoutExcelTunel.KM, 10 },
                                    {ELayoutExcelTunel.EQUIP_SUPER,11 },
                                    {ELayoutExcelTunel.EQUIP, 12 },
                                    {ELayoutExcelTunel.KM_INICIO, 13 },
                                    {ELayoutExcelTunel.KM_FIM, 14 },
                                    {ELayoutExcelTunel.EXTENSAO, 15 },
                                    {ELayoutExcelTunel.LATITUDE_INICIO, 16 },
                                    {ELayoutExcelTunel.LATITUDE_FIM, 17 },
                                    {ELayoutExcelTunel.LONGITUDE_INICIO, 18 },
                                    {ELayoutExcelTunel.LONGITUDE_FIM, 19 },
                                    {ELayoutExcelTunel.LAUDO, 20 },
                                    {ELayoutExcelTunel.DEFEITO,21 },
                                    {ELayoutExcelTunel.PRIORIDADE,22 },
                                    {ELayoutExcelTunel.OBSERVACAO, 23 },
                                    {ELayoutExcelTunel.FOTOS, 24 },
                                    {ELayoutExcelTunel.OS, 25 },
                                    {ELayoutExcelTunel.SUB_TRECHO, 26 },
                                    {ELayoutExcelTunel.POWERAPPSID,27 },
                                    {ELayoutExcelTunel.ENG, 28 },
                                };
            }
            return null;
        }

        private string MontaLayoutEmail<T>(List<IGrouping<string, T>> itensEmail, List<T> itensRemover, List<T> itensCopiados) where T : Disciplina
        {
            return $@"<p>Itens que precisam ser analisados: {itensEmail.Count()}</p>
                        <p>Itens que foram removidos: {itensRemover.Count()}</p>
                        <p>Itens que foram copiados: {itensCopiados.Count()}</p>";
        }

        public void EnviarEmail(string assunto, string corpo, string[] anexos = null)
        {
            try
            {
                MailMessage mensagem = new MailMessage();
                mensagem.From = new MailAddress(_user);
                mensagem.To.Add(_destinatario);
                mensagem.Subject = assunto;
                mensagem.Body = corpo;
                mensagem.IsBodyHtml = true;

                if (anexos != null)
                {
                    foreach (var anexo in anexos)
                    {
                        mensagem.Attachments.Add(new Attachment(anexo));
                    }
                }

                using (SmtpClient smtp = new SmtpClient(_host, _port))
                {
                    smtp.Credentials = new NetworkCredential(_user, _password);
                    smtp.EnableSsl = true;
                    smtp.Send(mensagem);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao enviar e-mail: {ex.Message}");
            }
        }
    }
}
