using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BotGestaoDefeitos.Service
{
    public class InfraestruturaService : BaseService
    {
        public string LeArquivo(string path, string pathDefeito)
        {
            try
            {
                var listInfraestrutura = new List<Infraestrutura>();
                var itensRemover = new List<Infraestrutura>();
                var itensAnalise = new List<IGrouping<long, Infraestrutura>>();
                var layout = LayoutExcel();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var pacote = new ExcelPackage(new FileInfo(path)))
                {
                    var planilha = pacote.Workbook.Worksheets[0]; // Obtém a primeira planilha

                    int totalLinhas = planilha.Dimension.Rows;

                    listInfraestrutura = LeArquivoInfraestrutura(totalLinhas, planilha, layout);
                    VerificaRepetidosInfraestrutura(listInfraestrutura, ref itensAnalise, ref itensRemover);
                    RemoveItens(itensRemover.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);
                }
                AtualizarPowerQuery(pathDefeito);

                GravaArquivoInfraestrutura(itensAnalise, itensRemover, layout);

                return MontaLayoutEmail(itensAnalise, itensRemover);
            }
            catch (Exception e)
            {
                logErro.Error($"Erro ao ler arquivo - LeArquivo {path}", e);
                throw e;
            }
        }

        private List<Infraestrutura> LeArquivoInfraestrutura(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            var listInfraestrutura = new List<Infraestrutura>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                try
                {
                    if (string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Text))
                        break;
                    listInfraestrutura.Add(new Infraestrutura
                    {
                        linha = linha,
                        ID_REGISTRO = string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_REGISTRO]].Text) ? (long?)null : Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_REGISTRO]].Text.Replace(",00", "")),
                        ID_DEFEITO = string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_DEFEITO]].Text) ? (long?)null : Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_DEFEITO]].Text.Replace(",00", "")),
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
                        LADO = planilha.Cells[linha, layout[ELayoutExcelInfraestrutura.LADO]].Text,
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
                catch (Exception e)
                {
                    logErro.Error($"Erro ao ler linha {linha} - LeArquivoInfraestrutura", e);
                    throw e;
                }
            }
            return listInfraestrutura;
        }

        private void VerificaRepetidosInfraestrutura(List<Infraestrutura> listInfraestrutura, ref List<IGrouping<long, Infraestrutura>> itensEmail, ref List<Infraestrutura> itensRemover)
        {
            var itensagrupados = listInfraestrutura.Where(x => x.ID_REGISTRO.HasValue).GroupBy(x => x.ID_REGISTRO.Value).ToList();

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

        private void GravaArquivoInfraestrutura(List<IGrouping<long, Infraestrutura>> itensAnalise, List<Infraestrutura> itensRemover, Dictionary<string, int> layout)
        {
            logInfo.Info("Gravando arquivo Infraestrutura");
            // Configura a licença do EPPlus (obrigatório desde a versão 5)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var pacote = new ExcelPackage(new FileInfo(_pathaux)))
            {
                var planilha = pacote.Workbook.Worksheets["Infraestrutura_analise"];
                if (planilha != null)
                    pacote.Workbook.Worksheets.Delete("Infraestrutura_analise");
                var planilha2 = pacote.Workbook.Worksheets["Infraestrutura_excluidos"];
                if (planilha2 != null)
                    pacote.Workbook.Worksheets.Delete("Infraestrutura_excluidos");
                int linha = 2;

                if (itensAnalise.Any())
                {
                    var planilhaAnalise = pacote.Workbook.Worksheets.Add("Infraestrutura_analise");

                    // Preenche os cabeçalhos
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.ID_REGISTRO]].Value = "ID_Registro";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.ID_DEFEITO]].Value = "ID_Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.ID_RONDA]].Value = "ID_Ronda";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.ATUALIZACAO]].Value = "Atualização";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.TIPO_INSPECAO]].Value = "Tipo de Inspeção";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.DATA]].Value = "DATA ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.RESPONSAVEL]].Value = "Responsável";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.STATUS]].Value = "Status Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.SUB]].Value = "SUB_Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.KM]].Value = "Km_Nominal";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.EQUIP_SUPER]].Value = "Equip_Super ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.EQUIP]].Value = "Equip_Infra";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.KM_INICIO]].Value = "km Início Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.KM_FIM]].Value = "km Fim Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.EXTENSAO]].Value = "Extensão Defeito(m)";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.LATITUDE_INICIO]].Value = "Latitude Inicio";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.LATITUDE_FIM]].Value = "Latitude Fim";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.LONGITUDE_INICIO]].Value = "Longitude Inicio";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.LONGITUDE_FIM]].Value = "Longitude Fim";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.LADO]].Value = "Lado";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.DEFEITO]].Value = "DEFEITO";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.PRIORIDADE]].Value = "Prioridade";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.OBSERVACAO]].Value = "Observação";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.FOTOS]].Value = "Fotos";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.OS]].Value = "OS ";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.SUB_TRECHO]].Value = "Sub_Trecho";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.IDENT]].Value = "Ident_Elemen";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.POWERAPPSID]].Value = "__PowerAppsId__";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelInfraestrutura.ENG]].Value = "Eng";


                    foreach (var item in itensAnalise.SelectMany(x => x))
                    {
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_REGISTRO]].Value = item.ID_REGISTRO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_DEFEITO]].Value = item.ID_DEFEITO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_RONDA]].Value = item.ID_RONDA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.ATUALIZACAO]].Value = item.ATUALIZACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.DATA]].Value = item.DATA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.RESPONSAVEL]].Value = item.RESPONSAVEL;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.STATUS]].Value = item.STATUS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.SUB]].Value = item.SUB;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.KM]].Value = item.KM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.EQUIP]].Value = item.EQUIP;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.KM_INICIO]].Value = item.KM_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.KM_FIM]].Value = item.KM_FIM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.EXTENSAO]].Value = item.EXTENSAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.LATITUDE_INICIO]].Value = item.LATITUDE_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.LATITUDE_FIM]].Value = item.LATITUDE_FIM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.LONGITUDE_INICIO]].Value = item.LONGITUDE_INICIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.LONGITUDE_FIM]].Value = item.LONGITUDE_FIM;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.LADO]].Value = item.LADO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.DEFEITO]].Value = item.DEFEITO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.PRIORIDADE]].Value = item.PRIORIDADE;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.OBSERVACAO]].Value = item.OBSERVACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.FOTOS]].Value = item.FOTOS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.OS]].Value = item.OS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.SUB_TRECHO]].Value = item.SUB_TRECHO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.IDENT]].Value = item.IDENT;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.POWERAPPSID]].Value = item.POWERAPPSID;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelInfraestrutura.ENG]].Value = item.ENG;
                        linha++;
                    }
                }
                if (itensRemover.Any())
                {

                    var planilhaExcluidos = pacote.Workbook.Worksheets.Add("Infraestrutura_excluidos");

                    // Preenche os cabeçalhos
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.ID_REGISTRO]].Value = "ID_Registro";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.ID_DEFEITO]].Value = "ID_Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.ID_RONDA]].Value = "ID_Ronda";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.ATUALIZACAO]].Value = "Atualização";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.TIPO_INSPECAO]].Value = "Tipo de Inspeção";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.DATA]].Value = "DATA ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.RESPONSAVEL]].Value = "Responsável";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.STATUS]].Value = "Status Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.SUB]].Value = "SUB_Defeito_Co";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.KM]].Value = "Km_Nominal";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.EQUIP_SUPER]].Value = "Equip_Super ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.EQUIP]].Value = "Equip_Contencao";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.KM_INICIO]].Value = "km Início Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.KM_FIM]].Value = "km Fim Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.EXTENSAO]].Value = "Extensão Defeito(m)";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.LATITUDE_INICIO]].Value = "Latitude Inicio";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.LATITUDE_FIM]].Value = "Latitude Fim";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.LONGITUDE_INICIO]].Value = "Longitude Inicio";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.LONGITUDE_FIM]].Value = "Longitude Fim";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.LADO]].Value = "Lado";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.DEFEITO]].Value = "DEFEITO";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.PRIORIDADE]].Value = "Prioridade";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.OBSERVACAO]].Value = "Observação";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.FOTOS]].Value = "Fotos";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.OS]].Value = "OS ";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.SUB_TRECHO]].Value = "Sub_Trecho";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.IDENT]].Value = "Ident_Elemen";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.POWERAPPSID]].Value = "__PowerAppsId__";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelInfraestrutura.ENG]].Value = "Eng";
                    linha = 2;

                    foreach (var item in itensRemover)
                    {
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_REGISTRO]].Value = item.ID_REGISTRO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_DEFEITO]].Value = item.ID_DEFEITO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.ID_RONDA]].Value = item.ID_RONDA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.ATUALIZACAO]].Value = item.ATUALIZACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.TIPO_INSPECAO]].Value = item.TIPO_INSPECAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.DATA]].Value = item.DATA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.RESPONSAVEL]].Value = item.RESPONSAVEL;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.STATUS]].Value = item.STATUS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.SUB]].Value = item.SUB;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.KM]].Value = item.KM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.EQUIP]].Value = item.EQUIP;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.KM_INICIO]].Value = item.KM_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.KM_FIM]].Value = item.KM_FIM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.EXTENSAO]].Value = item.EXTENSAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.LATITUDE_INICIO]].Value = item.LATITUDE_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.LATITUDE_FIM]].Value = item.LATITUDE_FIM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.LONGITUDE_INICIO]].Value = item.LONGITUDE_INICIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.LONGITUDE_FIM]].Value = item.LONGITUDE_FIM;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.LADO]].Value = item.LADO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.DEFEITO]].Value = item.DEFEITO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.PRIORIDADE]].Value = item.PRIORIDADE;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.OBSERVACAO]].Value = item.OBSERVACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.FOTOS]].Value = item.FOTOS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.OS]].Value = item.OS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.SUB_TRECHO]].Value = item.SUB_TRECHO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.IDENT]].Value = item.IDENT;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.POWERAPPSID]].Value = item.POWERAPPSID;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelInfraestrutura.ENG]].Value = item.ENG;
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
                        logErro.Error("Erro ao salvar arquivo - GravaArquivoInfraestrutura", e);
                        throw e;
                    }
            }
        }

        private Dictionary<string, int> LayoutExcel()
        {
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
                                    {ELayoutExcelInfraestrutura.LADO, 21 },
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
        }
    }
}
