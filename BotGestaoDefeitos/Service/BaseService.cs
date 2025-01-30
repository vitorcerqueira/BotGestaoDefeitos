using OfficeOpenXml;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;

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

        public string MontaLayoutEmail<T>(List<IGrouping<string, T>> itensEmail, List<T> itensRemover, List<T> itensCopiados) where T : Disciplina
        {
            return $@"<p>Itens que precisam ser analisados: {itensEmail.Count()}</p>
                        <p>Itens que foram removidos: {itensRemover.Count()}</p>
                        <p>Itens que foram copiados: {itensCopiados.Count()}</p>";
        }


    }
}
