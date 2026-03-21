using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace BotGestaoDefeitos.Service
{
    public class ContencaoService : BaseService
    {
        private readonly string _pathaux;

        public ContencaoService()
        {
            _pathaux = ConfigurationManager.AppSettings["pathaux"];
        }

        public string LeArquivo(string path, string pathDefeito)
        {
            try
            {
                List<Contencao> listContencoes = new List<Contencao>();
                List<Contencao> itensRemover = new List<Contencao>();
                List<IGrouping<long, Contencao>> itensAnalise = new List<IGrouping<long, Contencao>>();
                Dictionary<string, int> layout = LayoutExcel();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage pacote = new ExcelPackage(new FileInfo(path)))
                {
                    ExcelWorksheet planilha = pacote.Workbook.Worksheets[0]; // Obtém a primeira planilha

                    int totalLinhas = planilha.Dimension.Rows;

                    listContencoes = LeArquivoContencao(totalLinhas, planilha, layout);
                    VerificaRepetidosContencoes(listContencoes, ref itensAnalise, ref itensRemover);
                    RemoveItens(itensRemover.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);

                    pacote.Dispose();
                }

                AtualizarPowerQuery(pathDefeito);

                GravaArquivoContencoes(itensAnalise, itensRemover, layout);
                return MontaLayoutEmail(itensAnalise, itensRemover);
            }
            catch (Exception e)
            {
                logErro.Error($"Erro ao ler arquivo - LeArquivo {path}", e);
                throw e;
            }
        }

        private List<Contencao> LeArquivoContencao(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            List<Contencao> listContencoes = new List<Contencao>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                try
                {
                    if (string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Text))
                        break;
                    listContencoes.Add(new Contencao
                    {
                        linha = linha,
                        ID_REGISTRO = string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Text) ? (long?)null : Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Text.Replace(",00", "")),
                        ID_DEFEITO = string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelContencao.ID_DEFEITO]].Text) ? (long?)null : Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelContencao.ID_DEFEITO]].Text.Replace(",00", "")),
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
                        LADO = planilha.Cells[linha, layout[ELayoutExcelContencao.LADO]].Text,
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
                catch (Exception e)
                {
                    logErro.Error($"Erro ao ler linha {linha} - LeArquivoContencao", e);
                    throw e;
                }
            }

            return listContencoes;
        }

        private void VerificaRepetidosContencoes(List<Contencao> listContencoes, ref List<IGrouping<long, Contencao>> itensEmail, ref List<Contencao> itensRemover)
        {
            List<IGrouping<long, Contencao>> itensagrupados = listContencoes.Where(x => x.ID_REGISTRO.HasValue).GroupBy(x => x.ID_REGISTRO.Value).ToList();

            IEnumerable<IGrouping<long, Contencao>> repetidos = itensagrupados.Where(x => x.Count() > 1);

            foreach (IGrouping<long, Contencao> rep in repetidos)
            {
                Contencao contencaoInicial = null;
                foreach (Contencao item in rep)
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

        private void GravaArquivoContencoes(List<IGrouping<long, Contencao>> itensAnalise, List<Contencao> itensRemover, Dictionary<string, int> layout)
        {
            logInfo.Info("Gravando arquivo de Contenções");
            // Configura a licença do EPPlus (obrigatório desde a versão 5)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage pacote = new ExcelPackage(new FileInfo(_pathaux)))
            {
                ExcelWorksheet planilha = pacote.Workbook.Worksheets["Contencoes_analise"];
                if (planilha != null)
                    pacote.Workbook.Worksheets.Delete("Contencoes_analise");
                ExcelWorksheet planilha2 = pacote.Workbook.Worksheets["Contencoes_excluidos"];
                if (planilha2 != null)
                    pacote.Workbook.Worksheets.Delete("Contencoes_excluidos");

                int linha = 2;

                if (itensAnalise.Any())
                {
                    ExcelWorksheet planilhaAnalise = pacote.Workbook.Worksheets.Add("Contencoes_analise");

                    // Preenche os cabeçalhos
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.ID_REGISTRO]].Value = "ID_Registro";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.ID_DEFEITO]].Value = "ID_Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.ID_RONDA]].Value = "ID_Ronda";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.ATUALIZACAO]].Value = "Atualização";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.TIPO_INSPECAO]].Value = "Tipo de Inspeção";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.DATA]].Value = "DATA ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.RESPONSAVEL]].Value = "Responsável";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.STATUS]].Value = "Status Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.SUB]].Value = "SUB_Defeito_Co";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.KM]].Value = "Km_Nominal";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.EQUIP_SUPER]].Value = "Equip_Super ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.EQUIP]].Value = "Equip_Contencao";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.KM_INICIO]].Value = "km Início Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.KM_FIM]].Value = "km Fim Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.EXTENSAO]].Value = "Extensão Defeito(m)";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.LATITUDE_INICIO]].Value = "Latitude Inicio";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.LATITUDE_FIM]].Value = "Latitude Fim";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.LONGITUDE_INICIO]].Value = "Longitude Inicio";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.LONGITUDE_FIM]].Value = "Longitude Fim";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.LADO]].Value = "Lado";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.DEFEITO]].Value = "DEFEITO";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.PRIORIDADE]].Value = "Prioridade";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.IDENT]].Value = "Ident_Elemen";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.OBSERVACAO]].Value = "Observação";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.FOTOS]].Value = "Fotos";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.OS]].Value = "OS ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.SUB_TRECHO]].Value = "Sub_Trecho";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.POWERAPPSID]].Value = "__PowerAppsId__";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.DATA2]].Value = "Data2";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelContencao.ENG]].Value = "Eng";

                    foreach (Contencao item in itensAnalise.SelectMany(x => x))
                    {
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Value = item.ID_REGISTRO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.ID_DEFEITO]].Value = item.ID_DEFEITO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.ID_RONDA]].Value = item.ID_RONDA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.ATUALIZACAO]].Value = item.ATUALIZACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.DATA]].Value = item.DATA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.RESPONSAVEL]].Value = item.RESPONSAVEL;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.STATUS]].Value = item.STATUS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.SUB]].Value = item.SUB;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.KM]].Value = item.KM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.EQUIP]].Value = item.EQUIP;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.KM_INICIO]].Value = item.KM_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.KM_FIM]].Value = item.KM_FIM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.EXTENSAO]].Value = item.EXTENSAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.LATITUDE_INICIO]].Value = item.LATITUDE_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.LATITUDE_FIM]].Value = item.LATITUDE_FIM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.LONGITUDE_INICIO]].Value = item.LONGITUDE_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.LONGITUDE_FIM]].Value = item.LONGITUDE_FIM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.LADO]].Value = item.LADO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.DEFEITO]].Value = item.DEFEITO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.PRIORIDADE]].Value = item.PRIORIDADE;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.IDENT]].Value = item.IDENT;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.OBSERVACAO]].Value = item.OBSERVACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.FOTOS]].Value = item.FOTOS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.OS]].Value = item.OS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.SUB_TRECHO]].Value = item.SUB_TRECHO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.POWERAPPSID]].Value = item.POWERAPPSID;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.DATA2]].Value = item.DATA2;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelContencao.ENG]].Value = item.ENG;
                        linha++;
                    }
                }

                if (itensRemover.Any())
                {
                    ExcelWorksheet planilhaExcluidos = pacote.Workbook.Worksheets.Add("Contencoes_excluidos");

                    // Preenche os cabeçalhos
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.ID_REGISTRO]].Value = "ID_Registro";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.ID_DEFEITO]].Value = "ID_Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.ID_RONDA]].Value = "ID_Ronda";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.ATUALIZACAO]].Value = "Atualização";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.TIPO_INSPECAO]].Value = "Tipo de Inspeção";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.DATA]].Value = "DATA ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.RESPONSAVEL]].Value = "Responsável";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.STATUS]].Value = "Status Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.SUB]].Value = "SUB_Defeito_Co";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.KM]].Value = "Km_Nominal";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.EQUIP_SUPER]].Value = "Equip_Super ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.EQUIP]].Value = "Equip_Contencao";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.KM_INICIO]].Value = "km Início Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.KM_FIM]].Value = "km Fim Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.EXTENSAO]].Value = "Extensão Defeito(m)";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.LATITUDE_INICIO]].Value = "Latitude Inicio";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.LATITUDE_FIM]].Value = "Latitude Fim";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.LONGITUDE_INICIO]].Value = "Longitude Inicio";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.LONGITUDE_FIM]].Value = "Longitude Fim";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.LADO]].Value = "Lado";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.DEFEITO]].Value = "DEFEITO";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.PRIORIDADE]].Value = "Prioridade";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.IDENT]].Value = "Ident_Elemen";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.OBSERVACAO]].Value = "Observação";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.FOTOS]].Value = "Fotos";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.OS]].Value = "OS ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.SUB_TRECHO]].Value = "Sub_Trecho";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.POWERAPPSID]].Value = "__PowerAppsId__";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.DATA2]].Value = "Data2";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelContencao.ENG]].Value = "Eng";
                    linha = 2;

                    foreach (Contencao item in itensRemover)
                    {
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Value = item.ID_REGISTRO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.ID_DEFEITO]].Value = item.ID_DEFEITO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.ID_RONDA]].Value = item.ID_RONDA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.ATUALIZACAO]].Value = item.ATUALIZACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.DATA]].Value = item.DATA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.RESPONSAVEL]].Value = item.RESPONSAVEL;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.STATUS]].Value = item.STATUS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.SUB]].Value = item.SUB;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.KM]].Value = item.KM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.EQUIP]].Value = item.EQUIP;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.KM_INICIO]].Value = item.KM_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.KM_FIM]].Value = item.KM_FIM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.EXTENSAO]].Value = item.EXTENSAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.LATITUDE_INICIO]].Value = item.LATITUDE_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.LATITUDE_FIM]].Value = item.LATITUDE_FIM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.LONGITUDE_INICIO]].Value = item.LONGITUDE_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.LONGITUDE_FIM]].Value = item.LONGITUDE_FIM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.LADO]].Value = item.LADO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.DEFEITO]].Value = item.DEFEITO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.PRIORIDADE]].Value = item.PRIORIDADE;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.IDENT]].Value = item.IDENT;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.OBSERVACAO]].Value = item.OBSERVACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.FOTOS]].Value = item.FOTOS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.OS]].Value = item.OS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.SUB_TRECHO]].Value = item.SUB_TRECHO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.POWERAPPSID]].Value = item.POWERAPPSID;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.DATA2]].Value = item.DATA2;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelContencao.ENG]].Value = item.ENG;
                        linha++;
                    }
                }

                // Salva o arquivo no disco

                if (itensRemover.Any() || itensAnalise.Any())
                    try
                    {
                        pacote.Save();
                    }
                    catch (Exception e)
                    {
                        logErro.Error("Erro ao salvar arquivo - GravaArquivoContencoes", e);
                        throw e;
                    }
            }
        }

        private Dictionary<string, int> LayoutExcel()
        {
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
                                    {ELayoutExcelContencao.LADO, 20 },
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
        }
    }
}
