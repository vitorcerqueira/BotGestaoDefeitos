using BotGestaoDefeitos.Service;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace BotGestaoDefeitos
{
    public class GestaoDefeitos
    {
        private readonly string _path;
        private readonly string _pathaux;
        private readonly string _user;
        private readonly string _password;
        private readonly string _host;
        private readonly int _port;
        private readonly string _destinatario;
        private List<Tuple<int, string, string>> _itensFiles;
        public GestaoDefeitos()
        {
            _path = ConfigurationManager.AppSettings["path"];
            _pathaux = ConfigurationManager.AppSettings["pathaux"];
            _user = ConfigurationManager.AppSettings["user"];
            _password = ConfigurationManager.AppSettings["password"];
            _host = ConfigurationManager.AppSettings["host"];
            _port = Convert.ToInt32(ConfigurationManager.AppSettings["port"]);
            _destinatario = ConfigurationManager.AppSettings["destinatario"];
        }
        public void ExecutarGestaoDefeitos()
        {
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
                    if (fileName.StartsWith("Histórico"))
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
                    email += $"<p>{item.Item3}</p>";
                    email += LeArquivo(item.Item2, item.Item3);
                }
                var pathGeral = _itensFiles.FirstOrDefault(x => x.Item1 == 4).Item2;
                var pathHistGeral = _itensFiles.FirstOrDefault(x => x.Item1 == 3).Item2;
               
                new BaseService().AtualizarPowerQuery(pathGeral);
                new BaseService().AtualizarPowerQuery(pathHistGeral);
               
                EnviarEmail("Itens processados RUMO", email, new string[] { _pathaux });

            }
            catch (Exception ex)
            {
                log4net.LogManager.GetLogger("Processamento.Geral.Erro").Error($"Falha ao realizar gestão de defeitos.", ex);
            }
        }

        private string LeArquivo(string path, string type)
        {
            var pathDefeito = _itensFiles.FirstOrDefault(x => x.Item1 == 2 && x.Item3 == type).Item2;
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
                    return "";
            }
            return "";
        }

        public void EnviarEmail(string assunto, string corpo, string[] anexos = null)
        {
            try
            {
                MailMessage mensagem = new MailMessage();
                mensagem.From = new MailAddress(_user);
                mensagem.To.Add(_destinatario);
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

                using (SmtpClient smtp = new SmtpClient(_host, _port))
                {
                    smtp.Credentials = new NetworkCredential(_user, _password);
                    smtp.EnableSsl = true;
                    smtp.Send(mensagem);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao enviar e-mail: {ex.Message}");
            }
        }
    }
}
