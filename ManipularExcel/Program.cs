using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ManipularExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // https://www.youtube.com/watch?v=Q-9P6anYkrU&ab_channel=JoseCarlosMacoratti
            //https://www.epplussoftware.com/


            //Definindo o Caminho:

            
            //string pathPlanilha = @"D:\Dados\xls\Vendas.xlsx";
            string pathPlanilha = @"D:\Dados\xls\Chamados.xlsx";

            //Console.Clear();

            //Console.WriteLine("-----< Precisone Enter para Iniciar >-----");
            //Console.ReadKey();
            
            //CriarPlanilhaExcel(pathPlanilha);
            
            Console.WriteLine("-----< Precisone Enter para Iniciar e Visualizar a Planilha >-----");
            Console.ReadKey();
            OpenPlanilhaExcel(pathPlanilha);

            Console.ReadKey();
        }

        

        private static void CriarPlanilhaExcel(string pathPlanilha)
        {
            // Criando um Tipo Anonimo:
            var LancVendas = new[]
            {
                new {id="SP101",Filial="São Paulo", Vendas=980},
                new {id="SP102",Filial="Rio de Janeiro", Vendas=840},
                new {id="SP103",Filial="Minas Gerais", Vendas=790},
                new {id="SP104",Filial="Bahia", Vendas=699},
                new {id="SP105",Filial="Paraná", Vendas=775},
                new {id="SP106",Filial="Porto Alegre", Vendas=660},
            };

            //List<Vendas> LancVendas = new List<Vendas>();

            //foreach (var item in LancVendas)
            //{
            //    item.Id = 01; item.Filial = "São Paulo"; item.VendasRealizadas = 980;
            //    item.Id = 02; item.Filial = "Rio de Janeiro"; item.VendasRealizadas = 840;
            //    item.Id = 03; item.Filial = "Rondonia "; item.VendasRealizadas = 790;
            //    item.Id = 04; item.Filial = "Acre"; item.VendasRealizadas = 699;
            //    item.Id = 05; item.Filial = "Parana"; item.VendasRealizadas = 775;
            //    item.Id = 06; item.Filial = "Paraiba"; item.VendasRealizadas = 660;
            //}


            // Definindo Licençar do ExcelPackege 

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage xls = new ExcelPackage();

            // Nome da Planilha 
            
            var workSheet = xls.Workbook.Worksheets.Add("PlanilhasVendas");
            
            // Definindo propriedade da Planilha
            
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            // Definindo propriedades da primeira linha
            
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Definindo o Cabeçalho da Planilha(Base 1)

            workSheet.Cells[1, 1].Value = "ID";
            workSheet.Cells[1, 2].Value = "Filial";
            workSheet.Cells[1, 3].Value = "Vendas";

            workSheet.Cells["A1:C1"].Style.Font.Italic = true;

            //Inclindo dados na Planilha 
            //Inicia na segunda linha

            int index = 2;
            foreach (var vd in LancVendas)
            {
                workSheet.Cells[index, 1].Value = vd.id;
                workSheet.Cells[index, 2].Value = vd.Filial;
                workSheet.Cells[index, 3].Value = vd.Vendas;
                index++;
            }
            // Ajuste o Tamanho da Coluna
            workSheet.Column(1).AutoFit();
            workSheet.Column(2).AutoFit();
            workSheet.Column(3).AutoFit();

            // Verificar exixtencia do arquivo
            if (File.Exists(pathPlanilha))      { File.Delete(pathPlanilha);    }
            
            //Criar Arquivo excel no Disco Fisico
            FileStream fileStream = File.Create(pathPlanilha);
            fileStream.Close();

            //Escrever o Conteudo no Arquivo
            File.WriteAllBytes(pathPlanilha, xls.GetAsByteArray());

            //Fechar arquivo
            xls.Dispose();
            Console.WriteLine($"Planilha Criada com Sucesso em:  {pathPlanilha}\n");
        }

        private static void OpenPlanilhaExcel(string pathPlanilha)
        {
            

            //Abre Planilha
            var arquivoXLS = new ExcelPackage(new FileInfo(pathPlanilha));

            //Localizar a Planilha a ser acessada:

            ExcelWorksheet PlanilhaVD = arquivoXLS.Workbook.Worksheets.FirstOrDefault();
            //ExcelWorksheet PlanilhaVD = arquivoXLS.Workbook.Worksheets["PlanilhasVendas"];

            //Obtendo o Numero de linhas e Colunas:
            int rows = PlanilhaVD.Dimension.Rows;
            int cols = PlanilhaVD.Dimension.Columns;


            //Percorrendo as linhas e colunas da Planilha

            for (int l = 1; l <=rows ; l++)
            {
                for (int c = 1; c <=cols; c++)
                {
                    var result = PlanilhaVD.Cells[l, c].Value.ToString();
                    Console.WriteLine(result);                  

                }
            }
        }
    }
}
