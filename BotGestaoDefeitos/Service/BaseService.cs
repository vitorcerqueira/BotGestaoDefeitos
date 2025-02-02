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
                pacote.Save(); // Salva as alterações no arquivo original
            }
        }

        public string MontaLayoutEmail<T>(List<IGrouping<long, T>> itensEmail, List<T> itensRemover) where T : Disciplina
        {
            return $@"<p>Itens que precisam ser analisados: {itensEmail.Count()}</p>
                        <p>Itens que foram removidos: {itensRemover.Count()}</p>";
        }

        public void AtualizarPowerQuery(string caminhoArquivo)
        {
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
                workbook.Save();
                workbook.Close();

                log4net.LogManager.GetLogger("Processamento.Geral.Info").Info($"Atualização concluída com sucesso. Arquivo {caminhoArquivo}");

            }
            catch (Exception ex)
            {
                log4net.LogManager.GetLogger("Processamento.Geral.Erro").Error($"Erro ao atualizar (arquivo : {caminhoArquivo}): {ex.Message}");
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
