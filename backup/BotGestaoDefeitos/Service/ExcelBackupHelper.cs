using System;
using System.IO;

namespace BotGestaoDefeitos.Service
{
    public class ExcelBackupHelper
    {
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
