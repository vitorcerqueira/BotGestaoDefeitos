using System;
using System.IO;
using OfficeOpenXml;

namespace BotGestaoDefeitos.Service
{
    public class ExcelBackupHelper
    {
        public static string CriarArquivoExecucao(string caminhoBase)
        {
            if (string.IsNullOrWhiteSpace(caminhoBase))
            {
                throw new ArgumentException("Caminho base do arquivo auxiliar não informado.", nameof(caminhoBase));
            }

            string pastaArquivo = Path.GetDirectoryName(caminhoBase);
            if (string.IsNullOrWhiteSpace(pastaArquivo))
            {
                throw new InvalidOperationException("Não foi possível determinar a pasta do arquivo auxiliar.");
            }

            Directory.CreateDirectory(pastaArquivo);

            string nomeArquivo = Path.GetFileNameWithoutExtension(caminhoBase);
            string extensao = Path.GetExtension(caminhoBase);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
            string caminhoExecucao = Path.Combine(pastaArquivo, $"{nomeArquivo}_{timestamp}{extensao}");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage pacote = new ExcelPackage())
            {
                ExcelWorksheet planilhaResumo = pacote.Workbook.Worksheets.Add("Resumo");
                planilhaResumo.Cells[1, 1].Value = "Gerado em";
                planilhaResumo.Cells[1, 2].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                pacote.SaveAs(new FileInfo(caminhoExecucao));
            }

            return caminhoExecucao;
        }

        public static void FazerBackupExcel(string caminhoArquivo)
        {
            if (string.IsNullOrWhiteSpace(caminhoArquivo) || !File.Exists(caminhoArquivo))
            {
                throw new FileNotFoundException("Arquivo Excel não encontrado: " + caminhoArquivo);
            }

            string pastaBackup = @"C:\temp\BotGestaoDefeitos\backup";

            if (!Directory.Exists(pastaBackup))
            {
                Directory.CreateDirectory(pastaBackup);
            }

            string nomeArquivo = Path.GetFileNameWithoutExtension(caminhoArquivo);
            string extensao = Path.GetExtension(caminhoArquivo);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            string caminhoBackup = Path.Combine(pastaBackup, $"{nomeArquivo}_backup_{timestamp}{extensao}");

            File.Copy(caminhoArquivo, caminhoBackup, overwrite: true);
        }
    }
}
