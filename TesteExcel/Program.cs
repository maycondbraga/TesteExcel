using ClosedXML.Excel;
using System;
using System.Linq;

namespace TesteExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var xls = new XLWorkbook(@"C:\Users\Maycon\Downloads\MOCK_DATA.xlsx");
            var planilha = xls.Worksheets.First(w => w.Name == "data");
            var totalLinhas = planilha.Rows().Count();

            // primeira linha é o cabecalho
            for (int l = 2; l <= totalLinhas; l++)
            {
                var codigo = planilha.Cell($"A{l}").Value.ToString();
                var descricao = planilha.Cell($"B{l}").Value.ToString();
                var preco = planilha.Cell($"C{l}").Value.ToString();
                Console.WriteLine($"{codigo} - {descricao} - {preco}");
            }

            Console.WriteLine($"\nSEGUNDA ABA ----------------------------------------\n");

            var planilha2 = xls.Worksheets.First(w => w.Name == "data2");
            var totalLinhas2 = planilha.Rows().Count();

            // primeira linha é o cabecalho
            for (int l = 2; l <= totalLinhas2; l++)
            {
                var codigo = planilha2.Cell($"A{l}").Value.ToString();
                var descricao = planilha2.Cell($"B{l}").Value.ToString();
                var preco = planilha2.Cell($"C{l}").Value.ToString();
                Console.WriteLine($"{codigo} - {descricao} - {preco}");
            }
        }
    }
}
