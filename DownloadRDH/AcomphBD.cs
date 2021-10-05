using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using DownloadCompass.DB;


namespace DownloadCompass
{
    class AcomphBD
    {

        public void CarregaAcomph(string path, string banco = "local")
        {
            Workbook wb = null;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //path = @"D:\Compass\Acomph\ACOMPH_31.03.2020.xls";
            try
            {

                Object[,] data_dados = null;
                excel.DisplayAlerts = false;
                excel.Visible = false;
                excel.ScreenUpdating = true;
                Workbook workbook = excel.Workbooks.Open(path);

                wb = excel.ActiveWorkbook;

                Sheets sheets = wb.Worksheets;

                var N_Sheets = sheets.Count;

                var Dados = new List<(string tabela, string[] campos, object[,] valores)>();
                var Postos = new List<(int Posto, object data)>();
                for (int i = 1; i <= N_Sheets; i++)
                {
                    Worksheet worksheet = (Worksheet)sheets.get_Item(i);
                    string sheetName = worksheet.Name;//Get the name of worksheet.
                    int num_posto = 9;
                    var posto = wb.Worksheets[sheetName].Cells[1, num_posto].Value;
                    int num_Reser = 3;
                    int num_Incre = 8;
                    int num_Nat = 9;
                    int num_Data = 1;

                    var datas = wb.Worksheets[sheetName].Range["A6", "A35"].Value as Object[,];
                    
                    data_dados = datas;

                    do
                    {
                        var range_Reserv = wb.Worksheets[sheetName].Range[wb.Worksheets[sheetName].Cells[6, num_Reser], wb.Worksheets[sheetName].Cells[35, num_Reser]].Value;
                        var range_Increm = wb.Worksheets[sheetName].Range[wb.Worksheets[sheetName].Cells[6, num_Incre], wb.Worksheets[sheetName].Cells[35, num_Incre]].Value;
                        var range_Natural = wb.Worksheets[sheetName].Range[wb.Worksheets[sheetName].Cells[6, num_Nat], wb.Worksheets[sheetName].Cells[35, num_Nat]].Value;
                        foreach (var data in datas)
                        {
                            var Reserv = range_Reserv[num_Data, 1];
                            var Increm = range_Increm[num_Data, 1];
                            var Natural = range_Natural[num_Data, 1];
                            num_Data = num_Data + 1;
                            //Inserte Aqui
                         //   IDB objSQL = new SQLServerDBCompass(banco);
                            string[] campos = { "[Data]", "[Posto]", "[Vaz_nat]", "[Vaz_Inc]", "[Reserv]" };
                            object[,] valores = new object[1, 5]    {
                                                        {
                                                            data,
                                                            posto,
                                                            Natural,
                                                            Increm,
                                                            Reserv
                                                        }
                                                    };
                            string tabela = "[dbo].[ACOMPH]";
                            Postos.Add((Convert.ToInt32(posto), data));
                           // objSQL.Execute("DELETE FROM [IPDO].[dbo].[ACOMPH] WHERE Posto = '" + posto + "' and Data ='" + Convert.ToDateTime(data).ToString("yyyy-MM-dd HH:mm:ss") + "'");

                            Dados.Add((tabela, campos, valores));
                        //    objSQL.Insert(tabela, campos, valores);




                        }


                        num_Reser = num_Reser + 8;
                        num_Incre = num_Incre + 8;
                        num_Nat = num_Nat + 8;
                        num_Data = 1;
                        num_posto = num_posto + 8;
                        posto = wb.Worksheets[sheetName].Cells[1, num_posto].Value;

                    } while (posto != null);


                }
              

                wb.Close();
                //workbook.Close();
                excel.Quit();

                inserir_Banco("local", Dados, Postos, data_dados);

                //inserir_Banco("azure", Dados, Postos, data_dados);
               

            }
            catch(Exception e)
            {
                wb.Close();
                excel.Quit();
            }




        }

       public void inserir_Banco(string banco, List<(string tabela, string[] campos, object[,] valores)> Dados, List<(int Posto, object data)> Postos, Object[,] Datas)
        {
            string query_Delete = "";
            string query_Insert = "";
            IDB objSQL = new SQLServerDBCompass(banco);


            foreach (var data in Datas)
            {
                objSQL.Execute("DELETE FROM [IPDO].[dbo].[ACOMPH] WHERE Data ='" + Convert.ToDateTime(data).ToString("yyyy-MM-dd HH:mm:ss") + "'");
            }


            // foreach (var p in Postos)
            //{
            //   query_Delete = query_Delete + "DELETE FROM [IPDO].[dbo].[ACOMPH] WHERE Posto = '" + p.Posto + "' and Data ='" + Convert.ToDateTime(p.data).ToString("yyyy-MM-dd HH:mm:ss") + "';";

            //}
            //objSQL.Execute("DELETE FROM [IPDO].[dbo].[ACOMPH] WHERE Posto = '" + p.Posto + "' and Data ='" + Convert.ToDateTime(p.data).ToString("yyyy-MM-dd HH:mm:ss") + "'");


            int i = 0;
            
            foreach (var Info in Dados)
            {
                if (i <= 300)
                {
                    query_Insert = query_Insert + "INSERT INTO [IPDO].[dbo].[ACOMPH] ( [Data],[Posto],[Vaz_nat],[Vaz_Inc],[Reserv] ) Values ('" + Convert.ToDateTime(Info.valores[0, 0]).ToString("yyyy-MM-dd HH:mm:ss") + "', '" + Info.valores[0, 1] + "', '" + Convert.ToInt32(Info.valores[0, 2]).ToString().Replace(',', '.')+ "', '" + Convert.ToInt32(Info.valores[0, 3]).ToString().Replace(',', '.') + "', '" + Convert.ToDouble(Info.valores[0, 4]).ToString().Replace(',', '.') + "');";
                    i++;
                }
                else
                {
                    query_Insert = query_Insert + "INSERT INTO [IPDO].[dbo].[ACOMPH] ( [Data],[Posto],[Vaz_nat],[Vaz_Inc],[Reserv] ) Values ('" + Convert.ToDateTime(Info.valores[0, 0]).ToString("yyyy-MM-dd HH:mm:ss") + "', '" + Info.valores[0, 1] + "', '" + Convert.ToInt32(Info.valores[0, 2]).ToString().Replace(',', '.') + "', '" + Convert.ToInt32(Info.valores[0, 3]).ToString().Replace(',', '.') + "', '" + Convert.ToDouble(Info.valores[0, 4]).ToString().Replace(',', '.') + "');";
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
