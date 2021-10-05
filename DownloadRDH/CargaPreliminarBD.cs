using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using DownloadCompass.DB;
using System.IO;

namespace DownloadCompass
{
    class CargaPreliminarBD
    {

        public void CarregaCarga(List<Tuple<string, DateTime, int, Nullable<decimal>>> dadosCarga, string banco = "local")
        {

            try
            {


                var Dados = new List<(string tabela, string[] campos, object[,] valores)>();
                var Postos = new List<(int Posto, object data)>();

                foreach (var dadosC in dadosCarga)
                {

                    //Inserte Aqui
                    //   IDB objSQL = new SQLServerDBCompass(banco);
                    string[] campos = { "[Data]", "[Minuto]", "[Submercado]", "[Carga]" };
                    object[,] valores = new object[1, 4]    {
                                                        {
                                                            dadosC.Item2,
                                                            dadosC.Item3,
                                                            dadosC.Item1,
                                                            dadosC.Item4

                                                        }
                                                    };
                    string tabela = "[dbo].[Carga_Preliminar]";

                    Dados.Add((tabela, campos, valores));

                }


                inserir_Banco("local", Dados);

                //inserir_Banco("azure", Dados);

            }
            catch (Exception e)
            {

            }




        }

        public void inserir_Banco(string banco, List<(string tabela, string[] campos, object[,] valores)> Dados)
        {
            //string query_Delete = "";
            string query_Insert = "";
            IDB objSQL = new SQLServerDBCompass(banco);
            //List<DateTime> datasI = new List<DateTime>();
            //var dates = Dados.Select(x => x.valores[0, 0]).Distinct().ToList();
            //foreach (var d in dates)
            //{

            //    query_Delete = query_Delete + "DELETE FROM [IPDO].[dbo].[Carga_Preliminar] WHERE Data ='" + Convert.ToDateTime(d).ToString("yyyy-MM-dd HH:mm:ss") + "';";

            //}
            //objSQL.Execute(query_Delete);




            int i = 0;

            foreach (var Info in Dados)
            {
                if (i <= 300)
                {
                    query_Insert = query_Insert + "INSERT INTO [IPDO].[dbo].[Carga_Preliminar] ( [Data], [Minuto], [Submercado], [Carga] ) Values ('" + Convert.ToDateTime(Info.valores[0, 0]).ToString("yyyy-MM-dd HH:mm:ss") + "', " + Info.valores[0, 1] + ", '" + Info.valores[0, 2].ToString() + "', " + Info.valores[0, 3].ToString().Replace(',', '.') + ");";
                    i++;
                }
                else
                {
                    query_Insert = query_Insert + "INSERT INTO [IPDO].[dbo].[Carga_Preliminar] ( [Data], [Minuto], [Submercado], [Carga] ) Values ('" + Convert.ToDateTime(Info.valores[0, 0]).ToString("yyyy-MM-dd HH:mm:ss") + "', " + Info.valores[0, 1] + ", '" + Info.valores[0, 2].ToString() + "', " + Info.valores[0, 3].ToString().Replace(',', '.') + ");";

                    //query_Insert = query_Insert + "INSERT INTO [IPDO].[dbo].[ACOMPH] ( [Data],[Posto],[Vaz_nat],[Vaz_Inc],[Reserv] ) Values ('" + Convert.ToDateTime(Info.valores[0, 0]).ToString("yyyy-MM-dd HH:mm:ss") + "', '" + Info.valores[0, 1] + "', '" + Convert.ToInt32(Info.valores[0, 2]).ToString().Replace(',', '.') + "', '" + Convert.ToInt32(Info.valores[0, 3]).ToString().Replace(',', '.') + "', '" + Convert.ToDouble(Info.valores[0, 4]).ToString().Replace(',', '.') + "');";
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
