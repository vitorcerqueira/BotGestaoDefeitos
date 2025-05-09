using BotGestaoDefeitos.Service;
using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web.UI.WebControls;

namespace BotGestaoDefeitos
{
    public class GestaoDefeitos
    {
        private readonly string _path;
        private readonly string _pathaux;

        private List<Tuple<int, string, string>> _itensFiles;

        private static readonly ILog logInfo = LogManager.GetLogger("Processamento.Geral.Info");
        private static readonly ILog logErro = LogManager.GetLogger("Processamento.Geral.Erro");

        public GestaoDefeitos()
        {
            _path = ConfigurationManager.AppSettings["path"];
            _pathaux = ConfigurationManager.AppSettings["pathaux"];
        }

        public void ExecutarGestaoDefeitos()
        {
            if (!Directory.Exists(Path.GetDirectoryName(_pathaux)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_pathaux));
            }

            //EnviarEmail("Itens processados RUMO", "testeeee");

            ListaDocumentos();
        }

        public void ListaDocumentos()
        {
            try
            {
                var pathFileSource = Directory.GetFiles(_path, "*.*", SearchOption.AllDirectories).ToList();
                _itensFiles = new List<Tuple<int, string, string>>();
                foreach (string path in pathFileSource)
                {
                    string[] pathFilePart = path.Split('\\');
                    string fileName = pathFilePart[pathFilePart.Length - 1];
                    string type = fileName.Split('_').Last().Substring(0, fileName.Split('_').Last().IndexOf("."));
                    if (fileName.StartsWith("Histórico") || fileName.StartsWith("Historico"))
                    {
                        if (fileName.Contains("Geral"))
                            _itensFiles.Add(new Tuple<int, string, string>(3, path, type));
                        else
                            _itensFiles.Add(new Tuple<int, string, string>(1, path, type));
                    }
                    if (fileName.StartsWith("Defeitos"))
                    {
                        if (fileName.Contains("Geral"))
                            _itensFiles.Add(new Tuple<int, string, string>(4, path, type));
                        else
                            _itensFiles.Add(new Tuple<int, string, string>(2, path, type));
                    }
                }
                var email = "";
                foreach (var item in _itensFiles.Where(x => x.Item1 == 1))
                {
                    logInfo.Info($"Iniciando arquivo: {item}");
                    email += $"<p>{item.Item3}</p>";
                    email += LeArquivo(item.Item2, item.Item3);
                }
                var pathGeral = _itensFiles.FirstOrDefault(x => x.Item1 == 4).Item2;
                var pathHistGeral = _itensFiles.FirstOrDefault(x => x.Item1 == 3).Item2;

                new BaseService().AtualizarPowerQuery(pathGeral);
                new BaseService().AtualizarPowerQuery(pathHistGeral);

                new BaseService().EnviarEmail("Itens processados RUMO", email, new string[] { _pathaux });

            }
            catch (Exception ex)
            {
                logErro.Error($"Falha ao realizar gestão de defeitos.", ex);
            }
        }

        private string LeArquivo(string path, string type)
        {
            var pathDefeito = _itensFiles.FirstOrDefault(x => x.Item1 == 2 && x.Item3 == type).Item2;

            if (BaseService.EsperarArquivoLiberado(path) && BaseService.EsperarArquivoLiberado(pathDefeito))
            {
                ExcelBackupHelper.FazerBackupExcel(path);

                ExcelBackupHelper.FazerBackupExcel(pathDefeito);

                switch (type)
                {
                    case "Bueiros":
                        return new BueiroService().LeArquivo(path, pathDefeito);
                    case "Contenções":
                        return new ContencaoService().LeArquivo(path, pathDefeito);
                    case "Infraestrutura":
                        return new InfraestruturaService().LeArquivo(path, pathDefeito);
                    case "PN":
                        return new PNService().LeArquivo(path, pathDefeito);
                    case "Túneis":
                        return new TunelService().LeArquivo(path, pathDefeito);
                    case "Pontes":
                        return new PonteService().LeArquivo(path, pathDefeito);
                }
            }

            return "";
        }
    }
}
