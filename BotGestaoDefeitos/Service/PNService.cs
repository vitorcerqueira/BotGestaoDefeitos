using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BotGestaoDefeitos.Service
{
    public class PNService : BaseService
    {
        public string LeArquivo(string path, string pathDefeito)
        {
            try
            {
                var listPN = new List<PN>();
                var itensRemover = new List<PN>();
                var itensAnalise = new List<IGrouping<long, PN>>();
                var layout = LayoutExcel();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var pacote = new ExcelPackage(new FileInfo(path)))
                {
                    var planilha = pacote.Workbook.Worksheets[0]; // Obtém a primeira planilha

                    int totalLinhas = planilha.Dimension.Rows;

                    listPN = LeArquivoPN(totalLinhas, planilha, layout);
                    VerificaRepetidosPN(listPN, ref itensAnalise, ref itensRemover);
                    RemoveItens(itensRemover.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);

                    pacote.Dispose();
                }

                AtualizarPowerQuery(pathDefeito);

                GravaArquivoPN(itensAnalise, itensRemover, layout);
                return MontaLayoutEmail(itensAnalise, itensRemover);
            }
            catch (Exception e)
            {
                logErro.Error($"Erro ao ler arquivo - LeArquivo {path}", e);
                throw e;
            }
        }

        private List<PN> LeArquivoPN(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            var listPN = new List<PN>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                try
                {
                    if (string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Text))
                        break;
                    listPN.Add(new PN
                    {
                        linha = linha,
                        ID_REGISTRO = string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelPN.ID_REGISTRO]].Text) ? (long?)null : Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelPN.ID_REGISTRO]].Text.Replace(",00", "")),
                        ID_DEFEITO = string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelPN.ID_DEFEITO]].Text) ? (long?)null : Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelPN.ID_DEFEITO]].Text.Replace(",00", "")),
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
                        LONGITUDE_INICIO = planilha.Cells[linha, layout[ELayoutExcelPN.LONGITUDE_INICIO]].Text,
                        LADO = planilha.Cells[linha, layout[ELayoutExcelPN.LADO]].Text,
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
                catch (Exception e)
                {
                    logErro.Error($"Erro ao ler linha {linha} - LeArquivoPN", e);
                    throw e;
                }
            }
            return listPN;
        }

        private void VerificaRepetidosPN(List<PN> listPN, ref List<IGrouping<long, PN>> itensAnalise, ref List<PN> itensRemover)
        {
            var itensagrupados = listPN.Where(x => x.ID_REGISTRO.HasValue).GroupBy(x => x.ID_REGISTRO.Value).ToList();

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
                            if (!itensAnalise.Contains(rep))
                                itensAnalise.Add(rep);
                        }
                    }
                }
            }

        }

        private void GravaArquivoPN(List<IGrouping<long, PN>> itensAnalise, List<PN> itensRemover, Dictionary<string, int> layout)
        {
            logInfo.Info("Gravando arquivo PN");
            // Configura a licença do EPPlus (obrigatório desde a versão 5)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var pacote = new ExcelPackage(new FileInfo(_pathaux)))
            {
                var planilha = pacote.Workbook.Worksheets["PN_analise"];
                if (planilha != null)
                    pacote.Workbook.Worksheets.Delete("PN_analise");
                var planilha2 = pacote.Workbook.Worksheets["PN_excluidos"];
                if (planilha2 != null)
                    pacote.Workbook.Worksheets.Delete("PN_excluidos");

                int linha = 2;

                if (itensAnalise.Any())
                {
                    var planilhaAnalise = pacote.Workbook.Worksheets.Add("PN_analise");

                    // Preenche os cabeçalhos
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.ID_REGISTRO]].Value = "ID_Registro";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.ID_DEFEITO]].Value = "ID_Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.ID_RONDA]].Value = "ID_Ronda";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.ATUALIZACAO]].Value = "Atualização";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.TIPO_INSPECAO]].Value = "Tipo de Inspeção";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.DATA]].Value = "DATA ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.RESPONSAVEL]].Value = "Responsável";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.STATUS]].Value = "Status Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.SUB]].Value = "SUB_Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.KM]].Value = "Km_Nominal";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.EQUIP_SUPER]].Value = "Equip_Super ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.EQUIP]].Value = "Equip_Infra";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.KM_INICIO]].Value = "km Início Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.KM_FIM]].Value = "km Fim Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.EXTENSAO]].Value = "Extensão Defeito(m)";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.LATITUDE_INICIO]].Value = "Latitude Inicio";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.LONGITUDE_INICIO]].Value = "Longitude Inicio";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.TIPO_TERRENO]].Value = "TIPO TERRENO";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.LADO]].Value = "Lado";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.DEFEITO]].Value = "DEFEITO";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.PRIORIDADE]].Value = "Prioridade defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.OBSERVACAO]].Value = "Observação";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.FOTOS]].Value = "Fotos";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.OS]].Value = "OS ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.SUB_TRECHO]].Value = "Sub_Trecho";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.POWERAPPSID]].Value = "__PowerAppsId__";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.ENG]].Value = "Eng";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPN.GRADE]].Value = "Grade";


                    foreach (var item in itensAnalise.SelectMany(x => x))
                    {
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.ID_REGISTRO]].Value = item.ID_REGISTRO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.ID_DEFEITO]].Value = item.ID_DEFEITO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.ID_RONDA]].Value = item.ID_RONDA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.ATUALIZACAO]].Value = item.ATUALIZACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.DATA]].Value = item.DATA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.RESPONSAVEL]].Value = item.RESPONSAVEL;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.STATUS]].Value = item.STATUS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.SUB]].Value = item.SUB;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.KM]].Value = item.KM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.EQUIP]].Value = item.EQUIP;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.KM_INICIO]].Value = item.KM_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.KM_FIM]].Value = item.KM_FIM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.EXTENSAO]].Value = item.EXTENSAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.LATITUDE_INICIO]].Value = item.LATITUDE_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.LONGITUDE_INICIO]].Value = item.LONGITUDE_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.TIPO_TERRENO]].Value = item.TIPO_TERRENO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.LADO]].Value = item.LADO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.DEFEITO]].Value = item.DEFEITO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.PRIORIDADE]].Value = item.PRIORIDADE;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.OBSERVACAO]].Value = item.OBSERVACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.FOTOS]].Value = item.FOTOS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.OS]].Value = item.OS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.SUB_TRECHO]].Value = item.SUB_TRECHO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.POWERAPPSID]].Value = item.POWERAPPSID;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.ENG]].Value = item.ENG;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPN.GRADE]].Value = item.GRADE;
                        linha++;
                    }
                }

                if (itensRemover.Any())
                {
                    var planilhaExcluidos = pacote.Workbook.Worksheets.Add("PN_excluidos");

                    // Preenche os cabeçalhos
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.ID_REGISTRO]].Value = "ID_Registro";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.ID_DEFEITO]].Value = "ID_Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.ID_RONDA]].Value = "ID_Ronda";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.ATUALIZACAO]].Value = "Atualização";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.TIPO_INSPECAO]].Value = "Tipo de Inspeção";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.DATA]].Value = "DATA ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.RESPONSAVEL]].Value = "Responsável";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.STATUS]].Value = "Status Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.SUB]].Value = "SUB_Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.KM]].Value = "Km_Nominal";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.EQUIP_SUPER]].Value = "Equip_Super ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.EQUIP]].Value = "Equip_Infra";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.KM_INICIO]].Value = "km Início Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.KM_FIM]].Value = "km Fim Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.EXTENSAO]].Value = "Extensão Defeito(m)";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.LATITUDE_INICIO]].Value = "Latitude Inicio";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.LONGITUDE_INICIO]].Value = "Longitude Inicio";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.TIPO_TERRENO]].Value = "TIPO TERRENO";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.LADO]].Value = "Lado";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.DEFEITO]].Value = "DEFEITO";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.PRIORIDADE]].Value = "Prioridade defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.OBSERVACAO]].Value = "Observação";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.FOTOS]].Value = "Fotos";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.OS]].Value = "OS ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.SUB_TRECHO]].Value = "Sub_Trecho";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.POWERAPPSID]].Value = "__PowerAppsId__";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.ENG]].Value = "Eng";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPN.GRADE]].Value = "Grade";
                    linha = 2;

                    foreach (var item in itensRemover)
                    {
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.ID_REGISTRO]].Value = item.ID_REGISTRO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.ID_DEFEITO]].Value = item.ID_DEFEITO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.ID_RONDA]].Value = item.ID_RONDA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.ATUALIZACAO]].Value = item.ATUALIZACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.DATA]].Value = item.DATA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.RESPONSAVEL]].Value = item.RESPONSAVEL;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.STATUS]].Value = item.STATUS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.SUB]].Value = item.SUB;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.KM]].Value = item.KM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.EQUIP]].Value = item.EQUIP;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.KM_INICIO]].Value = item.KM_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.KM_FIM]].Value = item.KM_FIM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.EXTENSAO]].Value = item.EXTENSAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.LATITUDE_INICIO]].Value = item.LATITUDE_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.LONGITUDE_INICIO]].Value = item.LONGITUDE_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.TIPO_TERRENO]].Value = item.TIPO_TERRENO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.LADO]].Value = item.LADO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.DEFEITO]].Value = item.DEFEITO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.PRIORIDADE]].Value = item.PRIORIDADE;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.OBSERVACAO]].Value = item.OBSERVACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.FOTOS]].Value = item.FOTOS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.OS]].Value = item.OS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.SUB_TRECHO]].Value = item.SUB_TRECHO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.POWERAPPSID]].Value = item.POWERAPPSID;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.ENG]].Value = item.ENG;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPN.GRADE]].Value = item.GRADE;
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
                        logErro.Error("Erro ao salvar arquivo - GravaArquivoPN", e);
                        throw e;
                    }
            }
        }

        private Dictionary<string, int> LayoutExcel()
        {
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
                                    {ELayoutExcelPN.LONGITUDE_INICIO, 17 },
                                    {ELayoutExcelPN.TIPO_TERRENO, 18 },
                                    {ELayoutExcelPN.LADO, 19 },
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
        }
    }
}
