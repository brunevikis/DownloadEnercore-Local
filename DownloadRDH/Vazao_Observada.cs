using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using DownloadCompass.DB;


namespace DownloadCompass
{
    class Vazao_Observada
    { 

        public void CarregaVazao(string path, string banco = "local")
        {

            
           // path = @"H:\Middle - Preço\Acompanhamento de vazões\Vazoes_Observadas\2020\11_2020\Vazões Observadas - " + data1.ToString("dd-MM-yyy") + " a " + data2.ToString("dd-MM-yyy") + ".xlsx";

            Workbook wb = null;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

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
                    var datas = wb.Worksheets[sheetName].Range["A8", "A97"].Value as Object[,];
                    data_dados = datas;
                    // Achar numero de tipos da Vazão
                    int num_vazao = 1;
                    int inicio_posto = 2;
                    int num_coluna = inicio_posto;
                    int num_posto = 0;

                    int falhas = 0;

                    while (falhas <= 7)
                    {



                        var Cel_posto = wb.Worksheets[sheetName].Cells[4, num_coluna].Value;
                        var Encontrou = Int32.TryParse(Convert.ToString(Cel_posto), out num_posto);
                        if (!Encontrou)
                        {
                            falhas++;
                            num_coluna++;
                            num_vazao++;
                        }
                        else
                        {
                            var posto = wb.Worksheets[sheetName].Cells[5, inicio_posto].Value;
                            for (int j = 0; j < num_vazao; j++)
                            {
                                var tipo_vazao = wb.Worksheets[sheetName].Cells[7, inicio_posto + j].Value;

                                int num_Vaz = inicio_posto;
                                int num_Data = 1;
                                var range_vazao = wb.Worksheets[sheetName].Range[wb.Worksheets[sheetName].Cells[8, inicio_posto + j], wb.Worksheets[sheetName].Cells[97, inicio_posto + j]].Value;

                                foreach (var data in datas)
                                {

                                    var Vazao = range_vazao[num_Data, 1];
                                    num_Data = num_Data + 1;

                                    string[] campos = { "[Data]", "[Cod_Posto]", "[Nome_Posto]", "[Bacia]", "[Tipo_Vazao]]", "[Vazao]", "[Data_Update]" };

                                    object[,] valores = new object[1, 6]    {
                                                {
                                                    data,
                                                    Cel_posto,
                                                    posto,
                                                    sheetName,
                                                    tipo_vazao,
                                                    Vazao
                                                }
                                            };
                                    string tabela = "[dbo].[Vazoes_Observadas]";

                                    Dados.Add((tabela, campos, valores));

                                }

                            }
                            inicio_posto = ++num_coluna;
                            num_vazao = 1;
                            falhas = 0;
                        }
                    }
                }


                wb.Close();
                //workbook.Close();
                excel.Quit();

                inserir_Banco("local", Dados, Postos, data_dados);

              //  inserir_Banco("azure", Dados, Postos, data_dados);


            }
            catch (Exception e)
            {
                wb.Close();
                excel.Quit();
            }
           

                
        }

       public void inserir_Banco(string banco, List<(string tabela, string[] campos, object[,] valores)> Dados, List<(int Posto, object data)> Postos, Object[,] Datas)
        {
            string query_Insert = "";
            IDB objSQL = new SQLServerDBCompass(banco);


            foreach (var data in Datas)
            {
                objSQL.Execute("DELETE FROM [IPDO].[dbo].[Vazoes_Observadas] WHERE Data ='" + Convert.ToDateTime(data).ToString("yyyy-MM-dd HH:mm:ss") + "'");
            }

            int i = 0;
            
            foreach (var Info in Dados)
            {
                if (i <= 300)
                {
                    query_Insert = query_Insert + "INSERT INTO [IPDO].[dbo].[Vazoes_Observadas] ( [Data], [Cod_Posto], [Nome_Posto], [Bacia], [Tipo_Vazao], [Vazao] ) Values ('" + Convert.ToDateTime(Info.valores[0, 0]).ToString("yyyy-MM-dd HH:mm:ss") + "', '" + Convert.ToInt32(Info.valores[0, 1]).ToString().Replace(',', '.') + "', '" + Info.valores[0, 2].ToString() + "', '" + Info.valores[0, 3].ToString() + "', '" + Info.valores[0, 4].ToString() + "', '" + Convert.ToDouble(Info.valores[0, 5]).ToString().Replace(',', '.') + "');";
                    i++;
                }
                else
                {
                    query_Insert = query_Insert + "INSERT INTO [IPDO].[dbo].[Vazoes_Observadas] ( [Data], [Cod_Posto], [Nome_Posto], [Bacia], [Tipo_Vazao], [Vazao] ) Values ('" + Convert.ToDateTime(Info.valores[0, 0]).ToString("yyyy-MM-dd HH:mm:ss") + "', '" + Convert.ToInt32(Info.valores[0, 1]).ToString().Replace(',', '.') + "', '" + Info.valores[0, 2].ToString() + "', '" + Info.valores[0, 3].ToString() + "', '" + Info.valores[0, 4].ToString() + "', '" + Convert.ToDouble(Info.valores[0, 5]).ToString().Replace(',', '.') + "');";
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
