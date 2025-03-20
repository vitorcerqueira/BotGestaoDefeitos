using log4net;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace BotGestaoDefeitos.Service
{
    public class BaseService
    {
        public readonly string _pathaux;
        public static readonly ILog logInfo = LogManager.GetLogger("Processamento.Geral.Info");
        public static readonly ILog logErro = LogManager.GetLogger("Processamento.Geral.Erro");
        public BaseService()
        {

            _pathaux = ConfigurationManager.AppSettings["pathaux"];
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

        public string MontaLayoutEmail<T>(List<IGrouping<long, T>> itensEmail, List<T> itensRemover) where T : Disciplina
        {
            return $@"<p>Itens que precisam ser analisados: {itensEmail.Count()}</p>
                        <p>Itens que foram removidos: {itensRemover.Count()}</p>";
        }

        public void AtualizarPowerQuery(string caminhoArquivo)
        {
            logInfo.Info($"Iniciando AtualizarPowerQuery. Arquivo {caminhoArquivo}");

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                // Inicia o Excel
                excelApp = new Excel.Application();
                excelApp.Visible = false; // Mantém o Excel em segundo plano

                // Abre a planilha
                workbook = excelApp.Workbooks.Open(caminhoArquivo);

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
                    logInfo.Info($"AtualizarPowerQuery [Save]. Arquivo {caminhoArquivo}");
                    workbook.Save();

                    logInfo.Info($"AtualizarPowerQuery [Close]. Arquivo {caminhoArquivo}");
                    workbook.Close();
                }
                catch (Exception ex) { logErro.Error($"Erro ao salvar arquivo - AtualizarPowerQuery: {ex.Message}", ex); }

                logInfo.Info($"Atualização AtualizarPowerQuery concluída com sucesso. Arquivo {caminhoArquivo}");

            }
            catch (Exception ex)
            {
                logErro.Error($"Erro ao atualizar (arquivo : {caminhoArquivo}): {ex.Message}");
            }
            finally
            {
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
        }
    }
}