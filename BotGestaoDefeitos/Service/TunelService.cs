using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BotGestaoDefeitos.Service
{
    public class TunelService : BaseService
    {
        public string LeArquivo(string path, string pathDefeito)
        {
            try
            {
                var listTunel = new List<Tunel>();
                var itensRemover = new List<Tunel>();
                var itensAnalise = new List<IGrouping<long, Tunel>>();
                var layout = LayoutExcel();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var pacote = new ExcelPackage(new FileInfo(path)))
                {
                    var planilha = pacote.Workbook.Worksheets[0]; // Obtém a primeira planilha

                    int totalLinhas = planilha.Dimension.Rows;

                    listTunel = LeArquivoTunel(totalLinhas, planilha, layout);
                    VerificaRepetidosTunel(listTunel, ref itensAnalise, ref itensRemover);
                    RemoveItens(itensRemover.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);
                }

                AtualizarPowerQuery(pathDefeito);

                GravaArquivoTunel(itensAnalise, itensRemover, layout);
                return MontaLayoutEmail(itensAnalise, itensRemover);
            }
            catch (Exception e)
            {
                logErro.Error($"Erro ao ler arquivo - LeArquivo {path}", e);
                throw e;
            }
        }

        private List<Tunel> LeArquivoTunel(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            var listTunel = new List<Tunel>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                try
                {
                    if (string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Text))
                        break;
                    listTunel.Add(new Tunel
                    {
                        linha = linha,
                        ID_REGISTRO = string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelTunel.ID_REGISTRO]].Text) ? (long?)null : Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelTunel.ID_REGISTRO]].Text.Replace(",00", "")),
                        ID_DEFEITO = string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelPN.ID_DEFEITO]].Text) ? (long?)null : Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelTunel.ID_DEFEITO]].Text.Replace(",00", "")),
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
                        LADO = planilha.Cells[linha, layout[ELayoutExcelTunel.LADO]].Text,
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
                catch (Exception e)
                {
                    logErro.Error($"Erro ao ler linha {linha} - LeArquivoTunel", e);
                    throw e;
                }
            }
            return listTunel;
        }

        private void VerificaRepetidosTunel(List<Tunel> listTunel, ref List<IGrouping<long, Tunel>> itensAnalise, ref List<Tunel> itensRemover)
        {
            var itensagrupados = listTunel.Where(x => x.ID_REGISTRO.HasValue).GroupBy(x => x.ID_REGISTRO.Value).ToList();

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
                            if (!itensAnalise.Contains(rep))
                                itensAnalise.Add(rep);
                        }
                    }
                }
            }

        }

        private void GravaArquivoTunel(List<IGrouping<long, Tunel>> itensAnalise, List<Tunel> itensRemover, Dictionary<string, int> layout)
        {
            logInfo.Info("Gravando arquivo Tunel");
            // Configura a licença do EPPlus (obrigatório desde a versão 5)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var pacote = new ExcelPackage(new FileInfo(_pathaux)))
            {
                var planilha = pacote.Workbook.Worksheets["Tunel_analise"];
                if (planilha != null)
                    pacote.Workbook.Worksheets.Delete("Tunel_analise");
                var planilha2 = pacote.Workbook.Worksheets["Tunel_excluidos"];
                if (planilha2 != null)
                    pacote.Workbook.Worksheets.Delete("Tunel_excluidos");

                int linha = 2;

                if (itensAnalise.Any())
                {
                    var planilhaAnalise = pacote.Workbook.Worksheets.Add("Tunel_analise");

                    // Preenche os cabeçalhos
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.ID_REGISTRO]].Value = "ID_Registro";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.ID_DEFEITO]].Value = "ID_Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.ID_RONDA]].Value = "ID_Ronda";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.ATUALIZACAO]].Value = "Atualização";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.TIPO_INSPECAO]].Value = "Tipo de Inspeção";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.DATA]].Value = "DATA ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.RESPONSAVEL]].Value = "Responsável";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.STATUS]].Value = "Status Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.SUB]].Value = "SUB_Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.KM]].Value = "Km_Nominal";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.EQUIP_SUPER]].Value = "Equip_Super";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.EQUIP]].Value = "Equip_Tunel";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.KM_INICIO]].Value = "km Início Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.KM_FIM]].Value = "km Fim Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.EXTENSAO]].Value = "Extensão Defeito(m)";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.LATITUDE_INICIO]].Value = "Latitude Inicio";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.LATITUDE_FIM]].Value = "Latitude Fim";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.LONGITUDE_INICIO]].Value = "Longitude Inicio";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.LONGITUDE_FIM]].Value = "Longitude Fim";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.LADO]].Value = "Lado";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.DEFEITO]].Value = "DEFEITO";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.PRIORIDADE]].Value = "Prioridade";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.OBSERVACAO]].Value = "Observação";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.FOTOS]].Value = "Fotos";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.OS]].Value = "OS ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.SUB_TRECHO]].Value = "Sub_Trecho";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.POWERAPPSID]].Value = "__PowerAppsId__";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelTunel.ENG]].Value = "Eng";


                    foreach (var item in itensAnalise.SelectMany(x => x))
                    {
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.ID_REGISTRO]].Value = item.ID_REGISTRO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.ID_DEFEITO]].Value = item.ID_DEFEITO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.ID_RONDA]].Value = item.ID_RONDA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.ATUALIZACAO]].Value = item.ATUALIZACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.DATA]].Value = item.DATA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.RESPONSAVEL]].Value = item.RESPONSAVEL;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.STATUS]].Value = item.STATUS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.SUB]].Value = item.SUB;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.KM]].Value = item.KM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.EQUIP]].Value = item.EQUIP;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.KM_INICIO]].Value = item.KM_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.KM_FIM]].Value = item.KM_FIM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.EXTENSAO]].Value = item.EXTENSAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.LATITUDE_INICIO]].Value = item.LATITUDE_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.LATITUDE_FIM]].Value = item.LATITUDE_FIM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.LONGITUDE_INICIO]].Value = item.LONGITUDE_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.LONGITUDE_FIM]].Value = item.LONGITUDE_FIM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.LADO]].Value = item.LADO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.DEFEITO]].Value = item.DEFEITO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.PRIORIDADE]].Value = item.PRIORIDADE;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.OBSERVACAO]].Value = item.OBSERVACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.FOTOS]].Value = item.FOTOS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.OS]].Value = item.OS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.SUB_TRECHO]].Value = item.SUB_TRECHO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.POWERAPPSID]].Value = item.POWERAPPSID;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelTunel.ENG]].Value = item.ENG;
                        linha++;
                    }
                }

                if (itensRemover.Any())
                {
                    var planilhaExcluidos = pacote.Workbook.Worksheets.Add("Tunel_excluidos");

                    // Preenche os cabeçalhos

                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.ID_REGISTRO]].Value = "ID_Registro";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.ID_DEFEITO]].Value = "ID_Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.ID_RONDA]].Value = "ID_Ronda";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.ATUALIZACAO]].Value = "Atualização";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.TIPO_INSPECAO]].Value = "Tipo de Inspeção";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.DATA]].Value = "DATA ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.RESPONSAVEL]].Value = "Responsável";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.STATUS]].Value = "Status Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.SUB]].Value = "SUB_Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.KM]].Value = "Km_Nominal";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.EQUIP_SUPER]].Value = "Equip_Super";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.EQUIP]].Value = "Equip_Tunel";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.KM_INICIO]].Value = "km Início Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.KM_FIM]].Value = "km Fim Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.EXTENSAO]].Value = "Extensão Defeito(m)";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.LATITUDE_INICIO]].Value = "Latitude Inicio";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.LATITUDE_FIM]].Value = "Latitude Fim";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.LONGITUDE_INICIO]].Value = "Longitude Inicio";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.LONGITUDE_FIM]].Value = "Longitude Fim";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.LADO]].Value = "Lado";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.DEFEITO]].Value = "DEFEITO";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.PRIORIDADE]].Value = "Prioridade";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.OBSERVACAO]].Value = "Observação";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.FOTOS]].Value = "Fotos";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.OS]].Value = "OS ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.SUB_TRECHO]].Value = "Sub_Trecho";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.POWERAPPSID]].Value = "__PowerAppsId__";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelTunel.ENG]].Value = "Eng";
                    linha = 2;

                    foreach (var item in itensRemover)
                    {
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.ID_REGISTRO]].Value = item.ID_REGISTRO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.ID_DEFEITO]].Value = item.ID_DEFEITO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.ID_RONDA]].Value = item.ID_RONDA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.ATUALIZACAO]].Value = item.ATUALIZACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.DATA]].Value = item.DATA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.RESPONSAVEL]].Value = item.RESPONSAVEL;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.STATUS]].Value = item.STATUS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.SUB]].Value = item.SUB;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.KM]].Value = item.KM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.EQUIP]].Value = item.EQUIP;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.KM_INICIO]].Value = item.KM_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.KM_FIM]].Value = item.KM_FIM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.EXTENSAO]].Value = item.EXTENSAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.LATITUDE_INICIO]].Value = item.LATITUDE_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.LATITUDE_FIM]].Value = item.LATITUDE_FIM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.LONGITUDE_INICIO]].Value = item.LONGITUDE_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.LONGITUDE_FIM]].Value = item.LONGITUDE_FIM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.LADO]].Value = item.LADO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.DEFEITO]].Value = item.DEFEITO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.PRIORIDADE]].Value = item.PRIORIDADE;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.OBSERVACAO]].Value = item.OBSERVACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.FOTOS]].Value = item.FOTOS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.OS]].Value = item.OS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.SUB_TRECHO]].Value = item.SUB_TRECHO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.POWERAPPSID]].Value = item.POWERAPPSID;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelTunel.ENG]].Value = item.ENG;
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
                        logErro.Error("Erro ao salvar arquivo - GravaArquivoTunel", e);
                        throw e;
                    }
            }
        }

        private Dictionary<string, int> LayoutExcel()
        {
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
                                    {ELayoutExcelTunel.LADO, 20 },
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
    }
}
