using ClosedXML.Excel;
using System;
using System.Linq;

namespace TesteExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //Puxa o excel
            var xls = new XLWorkbook(@"C:\Users\Maycon\Downloads\MOCK_DATA.xlsx");
            
            //Pega a aba da planilha pelo nome
            var planilha = xls.Worksheets.First(w => w.Name == "data");
            
            //Pega o total de linhas
            var totalLinhas = planilha.Rows().Count();

            //Inicia pelo 2 pois a primeira linha é o cabecalho
            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                var codigo = planilha.Cell(linha, 1).Value.ToString();
                var descricao = planilha.Cell(linha, 2).Value.ToString();
                var preco = planilha.Cell(linha, 3).Value.ToString();
                Console.WriteLine($"{codigo} - {descricao} - {preco}");
            }
        }
    }
}
