using DownloadCompass.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownloadCompass
{
    public class CarregaPrecip
    {

        public void Carrega_Chuvas(List<string> mapas, bool realizada = false)
        {
            List<string> linhas = new List<string>();

            var Paths = Directory.GetDirectories(@"P:\Alex Freires\anos CV\dat_modeloR\");
            var file_Sub = @"H:\TI - Sistemas\UAT\Download Compass\Arquivos Auxiliares\SubBacias.txt";

            foreach (var Pasta in Paths)
            {
                var files = Directory.GetFiles(Pasta);

                foreach (var file in files)
                {
                    //List<string> linhas = new List<string>();
                    System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex(@"(\d{2})(\d{2})(\d{2})");

                    var fMatch = r.Match(file);
                    if (fMatch.Success)
                    {
                        var data = new DateTime(
                            int.Parse(fMatch.Groups[3].Value) + 2000,
                            int.Parse(fMatch.Groups[2].Value),
                            int.Parse(fMatch.Groups[1].Value))
                            ;

                        using (TextReader tr = new StreamReader(file, Encoding.Default))
                        {
                            string sLinha = null;
                            while ((sLinha = tr.ReadLine()) != null)
                            {

                                var dados = sLinha.Split(' ');


                                if (dados[2] == "NaN")
                                {
                                    dados[2] = "0";
                                }

                                string Nome_usina = "";

                                using (TextReader trt = new StreamReader(file_Sub, Encoding.Default))
                                {
                                    string sLinhat = null;
                                    while ((sLinhat = trt.ReadLine()) != null)
                                    {

                                        var dadost = sLinhat.Split('#');

                                        if (dados[0] == dadost[2] && dados[1] == dadost[1])
                                        {
                                            Nome_usina = dadost[0];
                                        }



                                    }
                                    trt.Close();
                                }




                                IDB objSQL = new SQLServerDBCompass("local","CHUVAS");
                                string[] campos = { "[Fonte]", "[Data]", "[Sub_Bacia]", "[Precip]" };
                                object[,] valores = new object[1, 4]    {
                                                        {
                                                            "3",// lembrar de adicionar o tipo do modelo
                                                            data,
                                                            Nome_usina,
                                                            dados[2]
                                                        }
                                                    };
                                string tabela = "[dbo].[Chuvas_Sub_Bacias]";
                                objSQL.Insert(tabela, campos, valores);



                                // linhas.Add(sLinha); //adiciona cada linha do arquivo à lista
                            }



                            tr.Close();
                        }
                    }


                }
            }



        }

        public void Sub_Bacias()
        {



            var file = @"H:\TI - Sistemas\UAT\Download Compass\Arquivos Auxiliares\incluiBacias.txt";


            using (TextReader tr = new StreamReader(file, Encoding.Default))
            {
                string sLinha = null;
                while ((sLinha = tr.ReadLine()) != null)
                {

                    var dados = sLinha.Split('#');



                    IDB objSQL = new SQLServerDBCompass("local","CHUVAS");
                    string[] campos = { "[Nome]", "[Latitude]", "[Longitude]" };
                    object[,] valores = new object[1, 3]    {
                                                        {
                                                            dados[0],
                                                            double.Parse(dados[1]),
                                                            double.Parse(dados[2])
                                                        }
                                                    };
                    string tabela = "[dbo].[Sub_Bacias]";
                    objSQL.Insert(tabela, campos, valores);



                    // linhas.Add(sLinha); //adiciona cada linha do arquivo à lista
                }



                tr.Close();
            }








        }



        public void Patamares()
        {
            var caminho = @"C:\teste\patamar.csv";

            List<string> linhas = new List<string>();



            int semana = 0;
            int pesada = 0;
            int media = 0;
            int leve = 0;

            using (TextReader tr = new StreamReader(caminho, Encoding.Default))
            {
                string sLinha = null;
                while ((sLinha = tr.ReadLine()) != null)
                {
                    linhas.Add(sLinha); //adiciona cada linha do arquivo à lista
                }

                tr.Close();
            }

            foreach (var linha in linhas)
            {
                var separar = linha.Split(';');
                semana = Convert.ToInt32(separar[0]);
                pesada = Convert.ToInt32(separar[1]);
                media = Convert.ToInt32(separar[2]);
                leve = Convert.ToInt32(separar[3]);

                IDB objSQL = new SQLServerDBCompass();
                string[] campos = { "[Semana]", "[pesado]", "[medio]", "[leve]" };
                object[,] valores = new object[1, 4]    {
                                                        {
                                                            semana,
                                                            pesada,
                                                            media,
                                                            leve
                                                        }
                                                    };
                string tabela = "[dbo].[semanas_patamares]";
                objSQL.Insert(tabela, campos, valores);

            }

        }


    }
}
