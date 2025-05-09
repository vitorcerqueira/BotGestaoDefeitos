using log4net;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace BotGestaoDefeitos.Service
{
    public class BaseService
    {
        private readonly string _path;
        private readonly string _pathaux;
        private readonly string _user;
        private readonly string _password;
        private readonly string _host;
        private readonly int _port;
        private readonly string _destinatario;

        public static readonly ILog logInfo = LogManager.GetLogger("Processamento.Geral.Info");
        public static readonly ILog logErro = LogManager.GetLogger("Processamento.Geral.Erro");

        public BaseService()
        {
            _path = ConfigurationManager.AppSettings["path"];
            _pathaux = ConfigurationManager.AppSettings["pathaux"];
            _user = ConfigurationManager.AppSettings["user"];
            _password = ConfigurationManager.AppSettings["password"];
            _host = ConfigurationManager.AppSettings["host"];
            _port = Convert.ToInt32(ConfigurationManager.AppSettings["port"]);
            _destinatario = ConfigurationManager.AppSettings["destinatario"];
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

                if (EsperarArquivoLiberado(caminhoArquivo, 15)) // espera até 15 segundos
                {
                    // Abre a planilha
                    workbook = excelApp.Workbooks.Open(caminhoArquivo);

                    // Atualiza todas as consultas do Power Query
                    foreach (Excel.QueryTable query in workbook.Sheets[1].QueryTables)
                    {
                        logInfo.Info($"AtualizarPowerQuery [query.Refresh(false)]. Arquivo {caminhoArquivo}");

                        query.Refresh(false);
                    }

                    // Atualiza todas as conexões de dados (incluindo Power Query)
                    foreach (Excel.WorkbookConnection connection in workbook.Connections)
                    {
                        logInfo.Info($"AtualizarPowerQuery [Refresh Inicio]. Arquivo {caminhoArquivo}");

                        connection.OLEDBConnection.BackgroundQuery = false;
                        connection.Refresh();

                        logInfo.Info($"AtualizarPowerQuery [Refresh Fim]. Arquivo {caminhoArquivo}");
                    }

                    // Salva e fecha a planilha
                    try
                    {
                        logInfo.Info($"AtualizarPowerQuery [Save]. Arquivo {caminhoArquivo}");
                        workbook.Save();

                        logInfo.Info($"AtualizarPowerQuery [Close]. Arquivo {caminhoArquivo}");
                        workbook.Close();

                        // Fecha o Excel e libera os recursos
                        if (workbook != null) Marshal.ReleaseComObject(workbook);
                        if (excelApp != null)
                        {
                            excelApp.Quit();
                            Marshal.ReleaseComObject(excelApp);
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        logInfo.Info($"Atualização AtualizarPowerQuery concluída com sucesso. Arquivo {caminhoArquivo}");
                    }
                    catch (Exception ex)
                    {
                        logErro.Error($"Erro ao salvar arquivo - AtualizarPowerQuery: {ex.Message}", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                logErro.Error($"Erro ao atualizar (arquivo : {caminhoArquivo}): {ex.Message}");
            }
        }

        public static bool EsperarArquivoLiberado(string caminhoArquivo, int timeoutSegundos = 10, int intervaloMs = 500)
        {
            var tempoLimite = DateTime.Now.AddSeconds(timeoutSegundos);

            var tentativa = 0;

            while (DateTime.Now < tempoLimite)
            {
                tentativa++;

                if (EstaEmUso(caminhoArquivo))
                {
                    logInfo.Info($"EsperarArquivoLiberado [Arquivo em uso]. Tentativa: {tentativa} - Arquivo {caminhoArquivo}");
                }
                else
                {
                    return true; // Arquivo disponível
                }

                Thread.Sleep(intervaloMs); // Espera antes de tentar novamente
            }

            new BaseService().EnviarEmail("Arquivo em uso", caminhoArquivo);

            logInfo.Error($"AtualizarPowerQuery [Arquivo em uso]. Arquivo {caminhoArquivo}");

            return false; // Timeout atingido
        }

        private static bool EstaEmUso(string caminhoArquivo)
        {
            try
            {
                using (FileStream stream = new FileStream(caminhoArquivo, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false; // Está disponível
                }
            }
            catch (IOException)
            {
                return true; // Está em uso
            }
        }

        public static bool ObterContaExcelConectada()
        {
            try
            {
                var excelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                var userName = excelApp.UserName;
                var userEmail = excelApp.Application.DisplayFullScreen; // Não existe forma direta de pegar email

                if (userName.Length > 0)
                {
                    return true;
                }
                else
                {
                    logInfo.Error($"ObterContaExcelConectada -> Excel não conectado!");
                }
            }
            catch (Exception e)
            {
                logInfo.Error($"ObterContaExcelConectada -> Erro: {e.Message}");
            }

            return false;
        }

        public void EnviarEmail(string assunto, string corpo, string[] anexos = null)
        {
            try
            {
                MailMessage mensagem = new MailMessage();
                mensagem.From = new MailAddress(_user);

                foreach (var email in _destinatario.Split(';'))
                {
                    mensagem.To.Add(email);
                }

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

                logInfo.Info("Enviando e-mail");

                using (SmtpClient smtp = new SmtpClient(_host, _port))
                {
                    smtp.Credentials = new NetworkCredential(_user, _password);
                    smtp.EnableSsl = true;
                    smtp.Send(mensagem);
                }

                logInfo.Info("E-mail enviado");
            }
            catch (Exception ex)
            {
                logErro.Error($"Erro ao enviar e-mail: {ex.Message}", ex);
            }
        }
    }
}