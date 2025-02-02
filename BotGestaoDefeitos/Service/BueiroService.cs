using BotGestaoDefeitos;
using BotGestaoDefeitos.Service;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BotGestaoDefeitos.Service{
    public class BueiroService : BaseService
    {
        public string LeArquivo(string path, string pathDefeito)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var listBueiros = new List<Bueiro>();
                var itensRemover = new List<Bueiro>();
                var itensAnalise = new List<IGrouping<long, Bueiro>>();
                var layout = LayoutExcel();

                using (var pacote = new ExcelPackage(new FileInfo(path)))
                {
                    var planilha = pacote.Workbook.Worksheets[0]; // Obtém a primeira planilha

                    int totalLinhas = planilha.Dimension.Rows;

                    listBueiros = LeArquivoBueiro(totalLinhas, planilha, layout);
                    VerificaRepetidosBueiros(listBueiros, ref itensAnalise, ref itensRemover);
                    RemoveItens(itensRemover.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);
                }
                AtualizarPowerQuery(pathDefeito);

                GravaArquivoBueiros(itensAnalise, itensRemover, layout);
                return MontaLayoutEmail(itensAnalise, itensRemover);
            }
            catch(Exception e)
            {
                throw e;
            }
        }

        //private List<Bueiro> GravaItensDefeitos(List<Bueiro> historicoBueiroFinal, string pathDefeito, Dictionary<string, int> layout)
        //{


        //    var historicoBueirosCopiados = new List<Bueiro>();
        //    var defeitoBueirosRemovidos = new List<Bueiro>();
        //    var defeitoBueiros = new List<Bueiro>();

        //    using (var pacote = new ExcelPackage(new FileInfo(pathDefeito)))
        //    {
        //        var planilha = pacote.Workbook.Worksheets[0];

        //        int totalLinhas = planilha.Dimension.Rows;
        //        for (int linha = 2; linha <= totalLinhas; linha++)
        //        {
        //            defeitoBueiros.Add(new Bueiro()
        //            {
        //                linha = linha,
        //                ID_REGISTRO = Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelBueiro.ID_REGISTRO]].Text),
        //                ID_DEFEITO = Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelBueiro.ID_DEFEITO]].Text)
        //            });
        //        }

        //        var historicoBueiroAgrupados = historicoBueiroFinal.Where(x=> x.STATUS != "Excluido" && x.STATUS != "Reclassificado").GroupBy(x => x.ID_DEFEITO)
        //                                     .Select(x => new Bueiro { ID_REGISTRO = x.Max(y => y.ID_REGISTRO), ID_DEFEITO = x.Key }).ToList();

        //        var idsHistoricoBueiroAgrupados = historicoBueiroAgrupados.Select(y => y.ID_REGISTRO);

        //        historicoBueirosCopiados = historicoBueiroFinal.Where(x=> x.ID_REGISTRO.HasValue && idsHistoricoBueiroAgrupados.Contains(x.ID_REGISTRO) && !defeitoBueiros.Select(y=> y.ID_REGISTRO).Contains(x.ID_REGISTRO.Value) && x.STATUS != "Excluido" && x.STATUS != "Reclassificado").ToList();

        //        defeitoBueirosRemovidos = defeitoBueiros.Where(x => historicoBueiroAgrupados.Select(y => y.ID_DEFEITO).Contains(x.ID_DEFEITO) && !idsHistoricoBueiroAgrupados.Contains(x.ID_REGISTRO)).ToList();

        //        var defeitosRepetidos = defeitoBueiros.GroupBy(x => x.ID_DEFEITO).Where(x => x.Count() > 1);

        //        // listBueirosativos + listBueirosCopiados - listBueirosRemovidos = IdsRegistros
        //        //tenho a lista ativos agrupada p


        //        totalLinhas++;

        //        planilha.InsertRow(totalLinhas, historicoBueirosCopiados.Count());

        //        foreach (var item in historicoBueirosCopiados)
        //        {
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.ID_REGISTRO]].Value = item.ID_REGISTRO;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.ID_DEFEITO]].Value = item.ID_DEFEITO;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.ID_RONDA]].Value = item.ID_RONDA;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.ATUALIZACAO]].Value = item.ATUALIZACAO;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.DATA]].Value = item.DATA;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.RESPONSAVEL]].Value = item.RESPONSAVEL;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.STATUS]].Value = item.STATUS;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.SUB]].Value = item.SUB;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.KM]].Value = item.KM;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.EQUIP]].Value = item.EQUIP;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.LOCAL]].Value = item.LOCAL;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.DEFEITO]].Value = item.DEFEITO;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.PRIORIDADE]].Value = item.PRIORIDADE;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.OBSERVACAO]].Value = item.OBSERVACAO;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.FOTOS]].Value = item.FOTOS;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.OS]].Value = item.OS;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.SUB_TRECHO]].Value = item.SUB_TRECHO;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.POWERAPPSID]].Value = item.POWERAPPSID;
        //            planilha.Cells[totalLinhas, layout[ELayoutExcelBueiro.ENG]].Value = item.ENG;
        //            totalLinhas++;
        //        }

        //        pacote.Save();
        //    }
        //    return historicoBueirosCopiados;
        //}

        private List<Bueiro> LeArquivoBueiro(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            var listBueiros = new List<Bueiro>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                try
                {
                    if (string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Text))
                        break;
                    listBueiros.Add(new Bueiro
                    {
                        linha = linha,
                        ID_REGISTRO = Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelBueiro.ID_REGISTRO]].Text.Replace(",00", "")),
                        ID_DEFEITO = Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelBueiro.ID_DEFEITO]].Text.Replace(",00", "")),
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
                catch (Exception e) 
                {
                    throw e;
                }


            }
            return listBueiros;
        }
        
        private void VerificaRepetidosBueiros(List<Bueiro> listBueiros, ref List<IGrouping<long, Bueiro>> itensAnalise, ref List<Bueiro> itensRemover)
        {
            var itensagrupados = listBueiros.Where(x => x.ID_REGISTRO.HasValue).GroupBy(x => x.ID_REGISTRO.Value).ToList();

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
                            if (!itensAnalise.Contains(rep))
                                itensAnalise.Add(rep);
                        }
                    }
                }
            }

        }
       
        private void GravaArquivoBueiros(List<IGrouping<long, Bueiro>> itensAnalise, List<Bueiro> itensRemover, Dictionary<string, int> layout)
        {
            // Configura a licença do EPPlus (obrigatório desde a versão 5)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var pacote = new ExcelPackage(new FileInfo(_pathaux)))
            {
                var planilha = pacote.Workbook.Worksheets["Bueiros_analise"];
                if (planilha != null)
                    pacote.Workbook.Worksheets.Delete("Bueiros_analise");
                var planilha2 = pacote.Workbook.Worksheets["Bueiros_excluidos"];
                if (planilha2 != null)
                    pacote.Workbook.Worksheets.Delete("Bueiros_excluidos");

                int linha = 2;
                if (itensAnalise.Any())
                {
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

                    foreach (var item in itensAnalise.SelectMany(x => x))
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
                }

                if (itensRemover.Any())
                {
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

                    foreach (var item in itensRemover)
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
                }

                pacote.Save();
            }
        }

        private Dictionary<string, int> LayoutExcel()
        {
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
       
    }

}
