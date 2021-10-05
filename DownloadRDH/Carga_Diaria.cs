using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using DownloadCompass.DB;


namespace DownloadCompass
{
    class Carga_Diaria
    { 

        public void CarregaCarga(string path,DateTime data_carga, string banco = "local")
        {

            
           // path = @"H:\Middle - Preço\Acompanhamento de vazões\Vazoes_Observadas\2020\11_2020\Vazões Observadas - " + data1.ToString("dd-MM-yyy") + " a " + data2.ToString("dd-MM-yyy") + ".xlsx";

            Workbook wb = null;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            try
            {

                excel.DisplayAlerts = false;
                excel.Visible = false;
                excel.ScreenUpdating = true;
                Workbook workbook = excel.Workbooks.Open(path);

                wb = excel.ActiveWorkbook;

                Sheets sheets = wb.Worksheets;

                var N_Sheets = sheets.Count;

                var Dados = new List<(string tabela, string[] campos, object[,] valores)>();
               
                for (int i = 1; i <= N_Sheets; i++)
                {
                    Worksheet worksheet = (Worksheet)sheets.get_Item(i);
                    string sheetName = worksheet.Name;//Get the name of worksheet.

                    if (sheetName.Contains("Carga Horária"))
                    {
                        var range_dados = wb.Worksheets[sheetName].Range[wb.Worksheets[sheetName].Cells[8, 1], wb.Worksheets[sheetName].Cells[31,13]].Value;

                       for(int j = 1; j<=24; j++)
                        {
                            int hora = Convert.ToInt32(range_dados[j, 1]);
                            double Previsto;
                            double Verificado;
                            double Desvio;
                            int Submercado = 0;
                            for (int z = 2; z <= 13; z = z + 3)
                            {
                                
                                switch (z)
                                {
                                    case 2:
                                        Submercado = 1;
                                        break;
                                    case 5:
                                        Submercado = 2;
                                        break;
                                    case 8:
                                        Submercado = 3;
                                        break;
                                    case 11:
                                        Submercado = 4;
                                        break;
                                }

                                Previsto = Convert.ToDouble(range_dados[j, z]);
                                Verificado = Convert.ToDouble(range_dados[j, z+1]);
                                Desvio = Convert.ToDouble(range_dados[j, z+2]);

                                string[] campos = { "[Data]", "[Hora]", "[Submercado]", "[Previsto]", "[Verificado]", "[Desvio]" };

                                object[,] valores = new object[1, 6]    {
                                                {
                                                    data_carga,
                                                    hora,
                                                    Submercado,
                                                    Previsto,
                                                    Verificado,
                                                    Desvio
                                                }
                                            };
                                string tabela = "[dbo].[Carga_Diaria]";

                                Dados.Add((tabela, campos, valores));
                            }


                        }


                    }

                   
                }


                wb.Close();
                //workbook.Close();
                excel.Quit();

                inserir_Banco("local", Dados, data_carga);

                //inserir_Banco("azure", Dados, data_carga);


            }
            catch (Exception e)
            {
                wb.Close();
                excel.Quit();
            }
           

                
        }

       public void inserir_Banco(string banco, List<(string tabela, string[] campos, object[,] valores)> Dados, DateTime Data)
        {
            string query_Insert = "";
            IDB objSQL = new SQLServerDBCompass(banco);


           
            objSQL.Execute("DELETE FROM [IPDO].[dbo].[Carga_Diaria] WHERE Data ='" + Convert.ToDateTime(Data).ToString("yyyy-MM-dd HH:mm:ss") + "'");
            

            int i = 0;
            
            foreach (var Info in Dados)
            {
                if (i <= 300)
                {
                    query_Insert = query_Insert + "INSERT INTO [IPDO].[dbo].[Carga_Diaria] ( [Data], [Hora], [Submercado], [Previsto], [Verificado], [Desvio] ) Values ('" + Convert.ToDateTime(Info.valores[0, 0]).ToString("yyyy-MM-dd HH:mm:ss") + "', '" + Convert.ToInt32(Info.valores[0, 1]).ToString().Replace(',', '.') + "', '" + Convert.ToInt32(Info.valores[0, 2]).ToString().Replace(',', '.') + "', '" + Convert.ToDouble(Info.valores[0, 3]).ToString().Replace(',', '.') + "', '" + Convert.ToDouble(Info.valores[0, 4]).ToString().Replace(',', '.') + "', '" + Convert.ToDouble(Info.valores[0, 5]).ToString().Replace(',', '.') + "');";
                    i++;
                }
                else
                {
                    query_Insert = query_Insert + "INSERT INTO [IPDO].[dbo].[Carga_Diaria] ( [Data], [Hora], [Submercado], [Previsto], [Verificado], [Desvio] ) Values ('" + Convert.ToDateTime(Info.valores[0, 0]).ToString("yyyy-MM-dd HH:mm:ss") + "', '" + Convert.ToInt32(Info.valores[0, 1]).ToString().Replace(',', '.') + "', '" + Convert.ToInt32(Info.valores[0, 2]).ToString().Replace(',', '.') + "', '" + Convert.ToDouble(Info.valores[0, 3]).ToString().Replace(',', '.') + "', '" + Convert.ToDouble(Info.valores[0, 4]).ToString().Replace(',', '.') + "', '" + Convert.ToDouble(Info.valores[0, 5]).ToString().Replace(',', '.') + "');";
                    objSQL.Execute(query_Insert);
                    i = 0;
                    query_Insert = "";
                }
                //objSQL.Insert(Info.tabela, Info.campos, Info.valores);

            }
            objSQL.Execute(query_Insert);


        }
    }
}
