
using LereEscreverExcel.MinhasClasses;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace LereEscreverExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Livro livro = new Livro();
            livro.Titulo = "Biblia Sagrada";
            livro.ISBN = 1111;

            List<Livro> Lista = new List<Livro>();

            Lista.Add(livro) ;

            Livro livro2 = new Livro();
            livro2.Titulo = "Harry Potter";
            livro2.ISBN = 222;
            Lista.Add(livro2);

            Livro livro3 = new Livro();
            livro3.Titulo = "O Pequeno Príncipe";
            livro3.ISBN = 333;
            Lista.Add(livro3);

            Gerar(Lista);
            //Add é método, método indica uma ação; título é uma propriedade

        }

        static void LerExcel()
        {

        }
        public static void Gerar(List<Livro> livros)
        {
            // criando/abrindo o arquivo:
            FileInfo caminhoNomeArquivo = new FileInfo(@"C:\Fabiana\teste.xlsx");
            ExcelPackage arquivoExcel = new ExcelPackage(caminhoNomeArquivo);

            // CRIANDO (ADD) uma planilha neste arquivo e obtendo a referência para meu código operá-la.
            ExcelWorksheet planilha = arquivoExcel.Workbook.Worksheets.Add("Livros");

            // (operações para gerar o arquivo)
            // ESCREVENDO O CABECALHO
            // Uma forma de escrever: informando endereco da celula
            planilha.Cells["A1"].Value = livros[0].Titulo;
            planilha.Cells["B1"].Value = livros[0].ISBN;
            planilha.Cells["A2"].Value = livros[1].Titulo;
            planilha.Cells["B2"].Value = livros[1].ISBN;
            planilha.Cells["A3"].Value = livros[2].Titulo;
            planilha.Cells["B3"].Value = livros[2].ISBN;



            // salvando e fechando o arquivo: MUITO IMPORTANTE HEIN!!!
            arquivoExcel.Save();
            arquivoExcel.Dispose();
        }
    }
}
