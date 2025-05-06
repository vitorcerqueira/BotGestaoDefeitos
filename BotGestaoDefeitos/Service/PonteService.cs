using log4net;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace BotGestaoDefeitos.Service
{
    public class PonteService
    {
        public readonly string _pathaux;
        public static readonly ILog logInfo = LogManager.GetLogger("Processamento.Geral.Info");
        public static readonly ILog logErro = LogManager.GetLogger("Processamento.Geral.Erro");
        public PonteService()
        {

            _pathaux = ConfigurationManager.AppSettings["pathaux"];
        }
        public string LeArquivo(string path, string pathDefeito)
        {
            try
            {
                var listPonte = new List<Ponte>();
                var itensRemover = new List<Ponte>();
                var itensAnalise = new List<IGrouping<long, Ponte>>();
                var layout = LayoutExcel();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var pacote = new ExcelPackage(new FileInfo(path)))
                {
                    var planilha = pacote.Workbook.Worksheets[0]; // Obtém a primeira planilha

                    int totalLinhas = planilha.Dimension.Rows;

                    listPonte = LeArquivoPonte(totalLinhas, planilha, layout);
                    VerificaRepetidosPonte(listPonte, ref itensAnalise, ref itensRemover);
                    RemoveItens(itensRemover.Select(y => y.linha).OrderByDescending(x => x).ToList(), planilha, pacote);

                    pacote.Dispose();
                }

                AtualizarPowerQuery(pathDefeito);

                GravaArquivoPonte(itensAnalise, itensRemover, layout);
                return MontaLayoutEmail(itensAnalise, itensRemover);
            }
            catch (Exception e)
            {
                logErro.Error($"Erro ao ler arquivo - LeArquivo {path}", e);
                throw e;
            }
        }

        private List<Ponte> LeArquivoPonte(int totalLinhas, ExcelWorksheet planilha, Dictionary<string, int> layout)
        {
            var listPonte = new List<Ponte>();

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                try
                {
                    if (string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelContencao.ID_REGISTRO]].Text))
                        break;
                    listPonte.Add(new Ponte
                    {
                        linha = linha,
                        ID_REGISTRO = string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelPonte.ID_REGISTRO]].Text) ? (long?)null : Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelPonte.ID_REGISTRO]].Text.Replace(",00", "")),
                        ID_DEFEITO = string.IsNullOrEmpty(planilha.Cells[linha, layout[ELayoutExcelPonte.ID_DEFEITO]].Text) ? (long?)null :  Convert.ToInt64(planilha.Cells[linha, layout[ELayoutExcelPonte.ID_DEFEITO]].Text.Replace(",00", "")),
                        STATUS = planilha.Cells[linha, layout[ELayoutExcelPonte.STATUS]].Text,
                        ATUALIZACAO = planilha.Cells[linha, layout[ELayoutExcelPonte.ATUALIZACAO]].Text,
                        SUBDIVISAO = planilha.Cells[linha, layout[ELayoutExcelPonte.SUBDIVISAO]].Text,
                        SUBDIVISAOD2 = planilha.Cells[linha, layout[ELayoutExcelPonte.SUBDIVISAOD2]].Text,
                        KM_NOMINAL = planilha.Cells[linha, layout[ELayoutExcelPonte.KM_NOMINAL]].Text,
                        EXTENSAO = planilha.Cells[linha, layout[ELayoutExcelPonte.EXTENSAO]].Text,
                        PATIO = planilha.Cells[linha, layout[ELayoutExcelPonte.PATIO]].Text,
                        TRAMO = planilha.Cells[linha, layout[ELayoutExcelPonte.TRAMO]].Text,
                        EQUIP_SUPER = planilha.Cells[linha, layout[ELayoutExcelPonte.EQUIP_SUPER]].Text,
                        EQUIP = planilha.Cells[linha, layout[ELayoutExcelPonte.EQUIP]].Text,
                        RAPEL = planilha.Cells[linha, layout[ELayoutExcelPonte.RAPEL]].Text,
                        RESPONSAVEL = planilha.Cells[linha, layout[ELayoutExcelPonte.RESPONSAVEL]].Text,
                        DATA = planilha.Cells[linha, layout[ELayoutExcelPonte.DATA]].Text,
                        LOCALOAE = planilha.Cells[linha, layout[ELayoutExcelPonte.LOCALOAE]].Text,
                        MATERIAL = planilha.Cells[linha, layout[ELayoutExcelPonte.MATERIAL]].Text,
                        ELEMENTO = planilha.Cells[linha, layout[ELayoutExcelPonte.ELEMENTO]].Text,
                        IDENTIFICACAO = planilha.Cells[linha, layout[ELayoutExcelPonte.IDENTIFICACAO]].Text,
                        MANIFESTACAO = planilha.Cells[linha, layout[ELayoutExcelPonte.MANIFESTACAO]].Text,
                        GRAVIDADE = planilha.Cells[linha, layout[ELayoutExcelPonte.GRAVIDADE]].Text,
                        ABRANGENCIA = planilha.Cells[linha, layout[ELayoutExcelPonte.ABRANGENCIA]].Text,
                        NOTA = planilha.Cells[linha, layout[ELayoutExcelPonte.NOTA]].Text,
                        DISCIPLINA = planilha.Cells[linha, layout[ELayoutExcelPonte.DISCIPLINA]].Text,
                        OBSERVACAO = planilha.Cells[linha, layout[ELayoutExcelPonte.OBSERVACAO]].Text,
                        FOTOS = planilha.Cells[linha, layout[ELayoutExcelPonte.FOTOS]].Text,
                        OS = planilha.Cells[linha, layout[ELayoutExcelPonte.OS]].Text,
                        ENG = planilha.Cells[linha, layout[ELayoutExcelPonte.ENG]].Text,
                        QTDEDORMENTES = planilha.Cells[linha, layout[ELayoutExcelPonte.QTDEDORMENTES]].Text,
                        DATAAUX = planilha.Cells[linha, layout[ELayoutExcelPonte.DATAAUX]].Text,
                    });
                }
                catch (Exception e)
                {
                    logErro.Error($"Erro ao ler linha {linha} - LeArquivoPonte", e);
                    throw e;
                }
            }
            return listPonte;
        }

        private void VerificaRepetidosPonte(List<Ponte> listPonte, ref List<IGrouping<long, Ponte>> itensAnalise, ref List<Ponte> itensRemover)
        {
            var itensagrupados = listPonte.Where(x => x.ID_REGISTRO.HasValue).GroupBy(x => x.ID_REGISTRO.Value).ToList();

            var repetidos = itensagrupados.Where(x => x.Count() > 1);

            foreach (var rep in repetidos)
            {
                Ponte PonteInicial = null;
                foreach (var item in rep)
                {
                    if (PonteInicial == null)
                        PonteInicial = item;
                    else
                    {
                        if (PonteInicial.ID_DEFEITO == item.ID_DEFEITO
                          && PonteInicial.DATA == item.DATA
                          && PonteInicial.RESPONSAVEL == item.RESPONSAVEL
                          && PonteInicial.STATUS == item.STATUS)
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

        private void GravaArquivoPonte(List<IGrouping<long, Ponte>> itensAnalise, List<Ponte> itensRemover, Dictionary<string, int> layout)
        {
            logInfo.Info("Gravando arquivo Ponte");
            // Configura a licença do EPPlus (obrigatório desde a versão 5)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var pacote = new ExcelPackage(new FileInfo(_pathaux)))
            {
                var planilha = pacote.Workbook.Worksheets["Ponte_analise"];
                if (planilha != null)
                    pacote.Workbook.Worksheets.Delete("Ponte_analise");
                var planilha2 = pacote.Workbook.Worksheets["Ponte_excluidos"];
                if (planilha2 != null)
                    pacote.Workbook.Worksheets.Delete("Ponte_excluidos");

                int linha = 2;

                if (itensAnalise.Any())
                {
                    var planilhaAnalise = pacote.Workbook.Worksheets.Add("Ponte_analise");

                    // Preenche os cabeçalhos
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.ID_REGISTRO]].Value = "ID_Registro";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.ID_DEFEITO]].Value = "ID_Defeito";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.STATUS]].Value = "Status";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.ATUALIZACAO]].Value = "Atualização";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.SUBDIVISAO]].Value = "Subdivisão";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.SUBDIVISAOD2]].Value = "SubdivisãoD2";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.KM_NOMINAL]].Value = "Km_nominal";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.EXTENSAO]].Value = "Extensão";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.PATIO]].Value = "Patio";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.TRAMO]].Value = "Tramo";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.EQUIP_SUPER]].Value = "Equip_Super";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.EQUIP]].Value = "Equip_Pontes";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.RAPEL]].Value = "Rapel";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.RESPONSAVEL]].Value = "Responsavel_Inspeção";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.DATA]].Value = "Data de Inspeção";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.LOCALOAE]].Value = "Local_na_OAE";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.MATERIAL]].Value = "Material";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.ELEMENTO]].Value = "Elemento";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.IDENTIFICACAO]].Value = "Identificacao_do_elemento";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.MANIFESTACAO]].Value = "Manifestacao_Patologica";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.GRAVIDADE]].Value = "Gravidade";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.ABRANGENCIA]].Value = "Abrangencia";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.NOTA]].Value = "Nota";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.DISCIPLINA]].Value = "Disciplina";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.OBSERVACAO]].Value = "Observação";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.FOTOS]].Value = "Fotos";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.OS]].Value = "OS";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.ENG]].Value = "Eng";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.QTDEDORMENTES]].Value = "Qtde Dormentes Inservives";
                    planilhaAnalise.Cells[1, layout[ELayoutExcelPonte.DATAAUX]].Value = "Data_Aux";


                    foreach (var item in itensAnalise.SelectMany(x => x))
                    {
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.ID_REGISTRO]].Value = item.ID_REGISTRO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.ID_DEFEITO]].Value = item.ID_DEFEITO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.STATUS]].Value = item.STATUS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.ATUALIZACAO]].Value = item.ATUALIZACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.SUBDIVISAO]].Value = item.SUBDIVISAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.SUBDIVISAOD2]].Value = item.SUBDIVISAOD2;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.KM_NOMINAL]].Value = item.KM_NOMINAL;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.EXTENSAO]].Value = item.EXTENSAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.PATIO]].Value = item.PATIO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.TRAMO]].Value = item.TRAMO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.EQUIP]].Value = item.EQUIP;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.RAPEL]].Value = item.RAPEL;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.RESPONSAVEL]].Value = item.RESPONSAVEL;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.DATA]].Value = item.DATA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.LOCALOAE]].Value = item.LOCALOAE;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.MATERIAL]].Value = item.MATERIAL;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.ELEMENTO]].Value = item.ELEMENTO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.IDENTIFICACAO]].Value = item.IDENTIFICACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.MANIFESTACAO]].Value = item.MANIFESTACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.GRAVIDADE]].Value = item.GRAVIDADE;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.ABRANGENCIA]].Value = item.ABRANGENCIA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.NOTA]].Value = item.NOTA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.DISCIPLINA]].Value = item.DISCIPLINA;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.OBSERVACAO]].Value = item.OBSERVACAO;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.FOTOS]].Value = item.FOTOS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.OS]].Value = item.OS;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.ENG]].Value = item.ENG;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.QTDEDORMENTES]].Value = item.QTDEDORMENTES;
                        planilhaAnalise.Cells[linha, layout[ELayoutExcelPonte.DATAAUX]].Value = item.DATAAUX;
                        linha++;
                    }
                }

                if (itensRemover.Any())
                {
                    var planilhaExcluidos = pacote.Workbook.Worksheets.Add("Ponte_excluidos");

                    // Preenche os cabeçalhos

                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.ID_REGISTRO]].Value = "ID_Registro";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.ID_DEFEITO]].Value = "ID_Defeito";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.STATUS]].Value = "Status";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.ATUALIZACAO]].Value = "Atualização";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.SUBDIVISAO]].Value = "Subdivisão";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.SUBDIVISAOD2]].Value = "SubdivisãoD2";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.KM_NOMINAL]].Value = "Km_nominal";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.EXTENSAO]].Value = "Extensão";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.PATIO]].Value = "Patio";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.TRAMO]].Value = "Tramo";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.EQUIP_SUPER]].Value = "Equip_Super";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.EQUIP]].Value = "Equip_Pontes";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.RAPEL]].Value = "Rapel";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.RESPONSAVEL]].Value = "Responsavel_Inspeção";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.DATA]].Value = "Data de Inspeção";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.LOCALOAE]].Value = "Local_na_OAE";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.MATERIAL]].Value = "Material";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.ELEMENTO]].Value = "Elemento";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.IDENTIFICACAO]].Value = "Identificacao_do_elemento";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.MANIFESTACAO]].Value = "Manifestacao_Patologica";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.GRAVIDADE]].Value = "Gravidade";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.ABRANGENCIA]].Value = "Abrangencia";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.NOTA]].Value = "Nota";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.DISCIPLINA]].Value = "Disciplina";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.OBSERVACAO]].Value = "Observação";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.FOTOS]].Value = "Fotos";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.OS]].Value = "OS";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.ENG]].Value = "Eng";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.QTDEDORMENTES]].Value = "Qtde Dormentes Inservives";
                    planilhaExcluidos.Cells[1, layout[ELayoutExcelPonte.DATAAUX]].Value = "Data_Aux";
                    linha = 2;

                    foreach (var item in itensRemover)
                    {
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.ID_REGISTRO]].Value = item.ID_REGISTRO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.ID_DEFEITO]].Value = item.ID_DEFEITO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.STATUS]].Value = item.STATUS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.ATUALIZACAO]].Value = item.ATUALIZACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.SUBDIVISAO]].Value = item.SUBDIVISAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.SUBDIVISAOD2]].Value = item.SUBDIVISAOD2;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.KM_NOMINAL]].Value = item.KM_NOMINAL;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.EXTENSAO]].Value = item.EXTENSAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.PATIO]].Value = item.PATIO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.TRAMO]].Value = item.TRAMO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.EQUIP_SUPER]].Value = item.EQUIP_SUPER;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.EQUIP]].Value = item.EQUIP;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.RAPEL]].Value = item.RAPEL;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.RESPONSAVEL]].Value = item.RESPONSAVEL;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.DATA]].Value = item.DATA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.LOCALOAE]].Value = item.LOCALOAE;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.MATERIAL]].Value = item.MATERIAL;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.ELEMENTO]].Value = item.ELEMENTO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.IDENTIFICACAO]].Value = item.IDENTIFICACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.MANIFESTACAO]].Value = item.MANIFESTACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.GRAVIDADE]].Value = item.GRAVIDADE;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.ABRANGENCIA]].Value = item.ABRANGENCIA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.NOTA]].Value = item.NOTA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.DISCIPLINA]].Value = item.DISCIPLINA;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.OBSERVACAO]].Value = item.OBSERVACAO;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.FOTOS]].Value = item.FOTOS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.OS]].Value = item.OS;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.ENG]].Value = item.ENG;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.QTDEDORMENTES]].Value = item.QTDEDORMENTES;
                        planilhaExcluidos.Cells[linha, layout[ELayoutExcelPonte.DATAAUX]].Value = item.DATAAUX;
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
                        logErro.Error("Erro ao salvar arquivo - GravaArquivoPonte", e);
                        throw e;
                    }
            }
        }

        private Dictionary<string, int> LayoutExcel()
        {
            return new Dictionary<string, int>
                                {
                {ELayoutExcelPonte.ID_REGISTRO      ,1},
                {ELayoutExcelPonte.ID_DEFEITO       ,2},
                {ELayoutExcelPonte.STATUS           ,3},
                {ELayoutExcelPonte.ATUALIZACAO      ,4},
                {ELayoutExcelPonte.SUBDIVISAO       ,5},
                {ELayoutExcelPonte.SUBDIVISAOD2     ,6},
                {ELayoutExcelPonte.KM_NOMINAL       ,7},
                {ELayoutExcelPonte.EXTENSAO         ,8},
                {ELayoutExcelPonte.PATIO            ,9},
                {ELayoutExcelPonte.TRAMO            ,10},
                {ELayoutExcelPonte.EQUIP_SUPER      ,11},
                {ELayoutExcelPonte.EQUIP            ,12},
                {ELayoutExcelPonte.RAPEL            ,13},
                {ELayoutExcelPonte.RESPONSAVEL      ,14},
                {ELayoutExcelPonte.DATA             ,15},
                {ELayoutExcelPonte.LOCALOAE         ,16},
                {ELayoutExcelPonte.MATERIAL         ,17},
                {ELayoutExcelPonte.ELEMENTO         ,18},
                {ELayoutExcelPonte.IDENTIFICACAO    ,19},
                {ELayoutExcelPonte.MANIFESTACAO     ,20},
                {ELayoutExcelPonte.GRAVIDADE        ,21},
                {ELayoutExcelPonte.ABRANGENCIA      ,22},
                {ELayoutExcelPonte.NOTA             ,23},
                {ELayoutExcelPonte.DISCIPLINA       ,24},
                {ELayoutExcelPonte.OBSERVACAO       ,25},
                {ELayoutExcelPonte.FOTOS            ,26},
                {ELayoutExcelPonte.OS               ,27},
                {ELayoutExcelPonte.ENG              ,28},
                {ELayoutExcelPonte.QTDEDORMENTES    ,29},
                {ELayoutExcelPonte.DATAAUX          ,30}
            };
        }

        public void RemoveItens(List<int> itensRemover, ExcelWorksheet planilha, ExcelPackage pacote)
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
                try// Salva as alterações no arquivo original
                {
                    pacote.Save();
                }
                catch (Exception ex)
                {
                    logErro.Error($"Erro ao salvar arquivo - RemoveItens: {ex.Message}", ex);
                }
            }
        }

        public string MontaLayoutEmail(List<IGrouping<long, Ponte>> itensEmail, List<Ponte> itensRemover) 
        {
            return $@"<p>Itens que precisam ser analisados: {itensEmail.Count()}</p>
                        <p>Itens que foram removidos: {itensRemover.Count()}</p>";
        }

        public void AtualizarPowerQuery(string caminhoArquivo)
        {

            try
            {
                // Inicia o Excel
               var excelApp = new Excel.Application();
                excelApp.Visible = false; // Mantém o Excel em segundo plano

                // Abre a planilha
                var workbook = excelApp.Workbooks.Open(caminhoArquivo);

                // Atualiza todas as consultas do Power Query
                foreach (Excel.QueryTable query in workbook.Sheets[1].QueryTables)
                {
                    query.Refresh(false);
                }

                // Atualiza todas as conexões de dados (incluindo Power Query)
                foreach (Excel.WorkbookConnection connection in workbook.Connections)
                {
                    connection.OLEDBConnection.BackgroundQuery = false;
                    connection.Refresh();
                }

                // Salva e fecha a planilha
                try
                {
                    workbook.Save();
                    workbook.Close();
                }
                catch (Exception ex) {
                    logErro.Error($"Erro ao salvar arquivo - AtualizarPowerQuery: {ex.Message}", ex);
                    throw ex;
                }

                logInfo.Info($"Atualização concluída com sucesso. Arquivo {caminhoArquivo}");

                // Fecha o Excel e libera os recursos
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                logErro.Error($"Erro ao atualizar (arquivo : {caminhoArquivo}): {ex.Message}");
                    throw ex;
            }
            finally
            {
            }
        }
    }
}
