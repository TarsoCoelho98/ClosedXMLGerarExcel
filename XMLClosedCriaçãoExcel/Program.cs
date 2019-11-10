using System;
using System.Data.SqlClient;
using ClosedXML.Excel;


namespace XMLClosedCriaçãoExcel
{
    class Program
    {
        const string stringConexao = @"Data Source = DESKTOP-HL277SI; Initial Catalog = DBSQLTeste; Integrated Security = true";
        const string comando = "SELECT * FROM teste";
        const string caminhoExcel = @"C:\Users\Dell\Desktop\arquivosTeste\novo.xlsx";
        const string nomePlan = "plan";
        const int primeiraLinha = 1;

        static void Main(string[] args)
        { 
            SqlConnection conexao = new SqlConnection(stringConexao);
            SqlCommand comandoSql = new SqlCommand(comando, conexao);

            var workbook = new XLWorkbook();
            var plan = workbook.AddWorksheet(nomePlan);

            // Cabeçalho 

            plan.Cell("A1").Value = "Relatorio de Teste";
            // var range = workbook.Range("A1:C1");
            // range.Merge().Style.Font.SetBold().Font.FontSize = 20;

            plan.Cell("A2").Value = "Id";
            plan.Cell("B2").Value = "Letra";

            Console.WriteLine();

            try
            {
                conexao.Open();
                SqlDataReader leitor = comandoSql.ExecuteReader();

                int linha = 3;

                while (leitor.Read())
                {
                    string id = leitor["id"].ToString();
                    string letra = leitor["letra"].ToString();

                    plan.Cell("A" + linha).Value = id;
                    plan.Cell("B" + linha).Value = letra;

                    linha++;
                }

                plan.Columns("1-2").AdjustToContents();
                workbook.SaveAs(caminhoExcel);
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                conexao.Close();
                workbook.Dispose();
            }

            Console.WriteLine("fim...");
            Console.ReadKey();
        }       
    }
}
