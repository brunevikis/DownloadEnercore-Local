using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DownloadCompass.DB;
using System.Data.Common;


namespace DownloadCompass
{
    class IPDODB
    {
        public static int id_ipdo = 0;
        public static double[,] termo = new double[5, 15];

        public void LoadProcess(string PathFileName, DateTime dta)
        {
            carregaIPDO(PathFileName, dta, "local");
            GerTerm(PathFileName, "local");
            IPDO_Sub(PathFileName, "local");
        }

        public static void carregaIPDO(string PathFileName, DateTime dta, string banco = "local")
        {



            string FileName = PathFileName.Split('\\').Last();


            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(PathFileName)))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets["IPDO"]; //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                var sb = new StringBuilder(); //this is your data 
                var Culture = System.Globalization.CultureInfo.GetCultureInfo("pt-BR");

                //DateTime data_ipdo = Convert.ToDateTime(myWorksheet.Cells["X6"].Value.ToString(), Culture.DateTimeFormat);
                DateTime data_ipdo = dta;
                int prod_itaipu = Convert.ToInt32(myWorksheet.Cells["O9"].Value.ToString());
                int prod_nuclear = Convert.ToInt32(myWorksheet.Cells["O10"].Value.ToString());

                int inter_se_fc = Convert.ToInt32(myWorksheet.Cells["V21"].Value.ToString());
                int inter_se_ne = 0;
                int inter_s_se = Convert.ToInt32(myWorksheet.Cells["V22"].Value.ToString());
                int inter_n_fc = Convert.ToInt32(myWorksheet.Cells["V19"].Value.ToString());
                int inter_fc_ne = Convert.ToInt32(myWorksheet.Cells["V20"].Value.ToString());



                // Dados Tabela IPDO INFO
                object[,] valores_Info = new object[1, 9]{
                                                        {
                                                            data_ipdo,
                                                            FileName,
                                                            prod_itaipu,
                                                            prod_nuclear,
                                                            inter_se_fc,
                                                            inter_se_ne,
                                                            inter_s_se,
                                                            inter_n_fc,
                                                            inter_fc_ne

                                                        }
                                                    };

                IPDO_Info(valores_Info, data_ipdo, banco);

            }

        }

        public static void IPDO_Info(object[,] valores, DateTime data_ipdo, string banco = "local")
        {
            IDB objSQL = new SQLServerDBCompass(banco, "IPDO");//"IPDO", banco
            DbDataReader reader = null;
            string[] campos = { "[dt_ipdo]", "[ds_arquivo]", "[prod_itaipu]", "[prod_nuclear]", "[inter_se-fc]", "[inter_se-ne]", "[inter_s-se]", "[inter_n-fc]", "[inter_fc-ne]" };

            string tabela = "[IPDO].[dbo].[IPDO_INFO]";

            string strQuery = String.Format(@"SELECT TOP 1 [id]
  FROM [IPDO].[dbo].[IPDO_INFO] order by dt_update desc");

            try
            {
                int id_ipdo_atual = 0;
                string strQuer2 = String.Format(@"SELECT TOP 1 [id]
  FROM [IPDO].[dbo].[IPDO_INFO] WHERE dt_ipdo ='" + data_ipdo.ToString("yyyy-MM-dd HH:mm:ss") + "'");

                reader = objSQL.GetReader(strQuer2);
                while (reader.Read())
                {
                    id_ipdo_atual = Convert.ToInt32(reader[0]);
                }
                if (id_ipdo_atual != 0)
                {
                    objSQL.Execute("DELETE FROM [IPDO].[dbo].[IPDO_gerTerm] WHERE id_ipdo = " + id_ipdo_atual);
                    objSQL.Execute("DELETE FROM [IPDO].[dbo].[IPDO_Submercado] WHERE id_ipdo = " + id_ipdo_atual);
                    objSQL.Execute("DELETE FROM [IPDO].[dbo].[IPDO_INFO] WHERE dt_ipdo ='" + data_ipdo.ToString("yyyy-MM-dd HH:mm:ss") + "'");

                }

                objSQL.Insert(tabela, campos, valores);

                reader = objSQL.GetReader(strQuery);

                while (reader.Read())
                {
                    id_ipdo = Convert.ToInt32(reader[0]);
                }
            }
            finally
            {
                // Fecha o datareader
                if (reader != null)
                {
                    reader.Close();
                }
            }


        }

        public static void GerTerm(string PathFileName, string banco = "local")
        {



            string FileName = PathFileName.Split('\\').Last();

            string usina = "";
            string tipo = "";
            int tipo_int = 0;
            double valor_verificado = 0;
            string submercado = "";
            double valor_prog = 0.00;


            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(PathFileName)))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets["IPDO"]; //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                for (int i = 468; i <= 576; i++)
                {
                    try
                    {
                        usina = myWorksheet.Cells["A" + i].Value.ToString();

                        if (usina != "")
                        {
                            tipo = myWorksheet.Cells["G" + i].Value.ToString();
                            tipo_int = verifica_tipo(tipo);


                            int j = i;
                            while (tipo_int != 0)
                            {
                                valor_verificado = Convert.ToDouble(myWorksheet.Cells["C" + j].Value.ToString());
                                valor_prog = Convert.ToDouble(myWorksheet.Cells["E" + j].Value.ToString());
                                termo[1, tipo_int] = termo[1, tipo_int] + valor_verificado;

                                submercado = "1";
                                IDB objSQL = new SQLServerDBCompass(banco, "IPDO");
                                string[] campos = { "[id_ipdo]", "[usina]", "[tipo]", "[valor_verif]", "[submercado]", "[valor_prog]" };
                                object[,] valores = new object[1, 6]    {
                                                        {
                                                            id_ipdo,
                                                            usina,
                                                            tipo_int,
                                                            valor_verificado,
                                                            submercado,
                                                            valor_prog
                                                        }
                                                    };
                                string tabela = "[dbo].[IPDO_gerTerm]";
                                objSQL.Insert(tabela, campos, valores);
                                j++;
                                tipo = myWorksheet.Cells["G" + j].Value.ToString();
                                tipo_int = verifica_tipo(tipo);
                                // termo[1, tipo_int] = termo[1, tipo_int] + valor_verificado;
                            }
                        }
                    }
                    catch (Exception e) { }

                }

                for (int i = 261; i <= 286; i++)
                {
                    try
                    {
                        usina = myWorksheet.Cells["A" + i].Value.ToString();
                        if (usina != "")
                        {
                            valor_verificado = Convert.ToDouble(myWorksheet.Cells["G" + i].Value.ToString());
                            valor_prog = Convert.ToDouble(myWorksheet.Cells["F" + i].Value.ToString());
                            tipo = myWorksheet.Cells["C" + i].Value.ToString();
                            tipo_int = verifica_tipo(tipo);


                            submercado = "1";
                            if (tipo_int != 0)
                            {
                                termo[1, tipo_int] = termo[1, tipo_int] + valor_verificado;
                                IDB objSQL = new SQLServerDBCompass(banco, "IPDO");
                                string[] campos = { "[id_ipdo]", "[usina]", "[tipo]", "[valor_verif]", "[submercado]", "[valor_prog]" };
                                object[,] valores = new object[1, 6]    {
                                                        {
                                                            id_ipdo,
                                                            usina,
                                                            tipo_int,
                                                            valor_verificado,
                                                            submercado,
                                                            valor_prog
                                                        }
                                                    };
                                string tabela = "[dbo].[IPDO_gerTerm]";
                                objSQL.Insert(tabela, campos, valores);
                            }
                        }
                    }
                    catch (Exception e)
                    {


                    }



                }

                for (int i = 317; i <= 325; i++)
                {
                    try
                    {
                        usina = myWorksheet.Cells["A" + i].Value.ToString();

                        if (usina != "")
                        {

                            valor_verificado = Convert.ToDouble(myWorksheet.Cells["G" + i].Value.ToString());
                            valor_prog = Convert.ToDouble(myWorksheet.Cells["F" + i].Value.ToString());
                            tipo = myWorksheet.Cells["C" + i].Value.ToString();
                            tipo_int = verifica_tipo(tipo);

                            termo[2, tipo_int] = termo[2, tipo_int] + valor_verificado;

                            submercado = "2";
                            if (tipo_int != 0)
                            {
                                IDB objSQL = new SQLServerDBCompass(banco, "IPDO");
                                string[] campos = { "[id_ipdo]", "[usina]", "[tipo]", "[valor_verif]", "[submercado]", "[valor_prog]" };
                                object[,] valores = new object[1, 6]    {
                                                        {
                                                            id_ipdo,
                                                            usina,
                                                            tipo_int,
                                                            valor_verificado,
                                                            submercado,
                                                            valor_prog
                                                        }
                                                    };
                                string tabela = "[dbo].[IPDO_gerTerm]";
                                objSQL.Insert(tabela, campos, valores);
                            }
                        }
                    }
                    catch (Exception e)
                    {


                    }

                }

                for (int i = 336; i <= 360; i++)
                {
                    try
                    {
                        usina = myWorksheet.Cells["A" + i].Value.ToString();

                        if (usina != "")
                        {
                            valor_verificado = Convert.ToDouble(myWorksheet.Cells["G" + i].Value.ToString());
                            valor_prog = Convert.ToDouble(myWorksheet.Cells["F" + i].Value.ToString());
                            tipo = myWorksheet.Cells["C" + i].Value.ToString();
                            tipo_int = verifica_tipo(tipo);

                            termo[3, tipo_int] = termo[3, tipo_int] + valor_verificado;

                            submercado = "3";
                            if (tipo_int != 0)
                            {
                                IDB objSQL = new SQLServerDBCompass(banco, "IPDO");
                                string[] campos = { "[id_ipdo]", "[usina]", "[tipo]", "[valor_verif]", "[submercado]", "[valor_prog]" };
                                object[,] valores = new object[1, 6]    {
                                                        {
                                                            id_ipdo,
                                                            usina,
                                                            tipo_int,
                                                            valor_verificado,
                                                            submercado,
                                                            valor_prog
                                                        }
                                                    };
                                string tabela = "[dbo].[IPDO_gerTerm]";
                                objSQL.Insert(tabela, campos, valores);
                            }
                        }
                    }
                    catch (Exception e)
                    {


                    }


                }

                for (int i = 379; i <= 393; i++)
                {
                    try
                    {
                        usina = myWorksheet.Cells["A" + i].Value.ToString();
                        if (usina != "")
                        {
                            valor_verificado = Convert.ToDouble(myWorksheet.Cells["G" + i].Value.ToString());
                            valor_prog = Convert.ToDouble(myWorksheet.Cells["F" + i].Value.ToString());
                            tipo = myWorksheet.Cells["C" + i].Value.ToString();
                            tipo_int = verifica_tipo(tipo);



                            termo[4, tipo_int] = termo[4, tipo_int] + valor_verificado;

                            submercado = "4";
                            if (tipo_int != 0)
                            {
                                IDB objSQL = new SQLServerDBCompass(banco, "IPDO");
                                string[] campos = { "[id_ipdo]", "[usina]", "[tipo]", "[valor_verif]", "[submercado]", "[valor_prog]" };
                                object[,] valores = new object[1, 6]    {
                                                        {
                                                            id_ipdo,
                                                            usina,
                                                            tipo_int,
                                                            valor_verificado,
                                                            submercado,
                                                            valor_prog
                                                        }
                                                    };
                                string tabela = "[dbo].[IPDO_gerTerm]";
                                objSQL.Insert(tabela, campos, valores);
                            }
                        }


                    }
                    catch (Exception e)
                    {


                    }


                }




            }



        }
        public static int verifica_tipo(string tipo)
        {
            switch (tipo)
            {
                case "REL":
                    return 1;
                case "OME":
                    return 2;
                case "INF":
                    return 3;
                case "EXP":
                    return 4;
                case "TE":
                    return 5;
                case "CGE":
                    return 6;
                case "PCI":
                    return 7;
                case "GFM":
                    return 8;
                case "GSB":
                    return 9;
                case "ERP":
                    return 10;
                case "---":
                    return 11;
                case "RRO":
                    return 12;
                case "UCM":
                    return 13;
                case "GEN":
                    return 14;
            }
            return 0;
        }

        public static void IPDO_Sub(string PathFileName, string banco = "local")
        {


            string FileName = PathFileName.Split('\\').Last();

            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(PathFileName)))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets["IPDO"]; //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;


                double ENA_SE = Convert.ToDouble(myWorksheet.Cells["M65"].Value.ToString());
                double ENA_S = Convert.ToDouble(myWorksheet.Cells["M64"].Value.ToString());
                double ENA_NE = Convert.ToDouble(myWorksheet.Cells["M63"].Value.ToString());
                double ENA_N = Convert.ToDouble(myWorksheet.Cells["M62"].Value.ToString());

                double hidro_SE = Math.Round(Convert.ToDouble(myWorksheet.Cells["O35"].Value.ToString()), 0);
                double hidro_S = Math.Round(Convert.ToDouble(myWorksheet.Cells["O43"].Value.ToString()), 0);
                double hidro_NE = Math.Round(Convert.ToDouble(myWorksheet.Cells["O27"].Value.ToString()), 0);
                double hidro_N = Math.Round(Convert.ToDouble(myWorksheet.Cells["O19"].Value.ToString()), 0);

                int termVeri_SE = Convert.ToInt32(myWorksheet.Cells["O36"].Value.ToString());
                int termVeri_S = Convert.ToInt32(myWorksheet.Cells["O44"].Value.ToString());
                int termVeri_NE = Convert.ToInt32(myWorksheet.Cells["O28"].Value.ToString());
                int termVeri_N = Convert.ToInt32(myWorksheet.Cells["O20"].Value.ToString());

                double Dif_term_SE = Math.Round(Convert.ToDouble(myWorksheet.Cells["H630"].Value.ToString()), 0);
                double Dif_term_S = Math.Round(Convert.ToDouble(myWorksheet.Cells["H631"].Value.ToString()), 0);
                double Dif_term_NE = Math.Round(Convert.ToDouble(myWorksheet.Cells["H632"].Value.ToString()), 0);
                double Dif_term_N = Math.Round(Convert.ToDouble(myWorksheet.Cells["H633"].Value.ToString()), 0);

                double termProg_SE = termVeri_SE - Dif_term_SE;
                double termProg_S = termVeri_S - Dif_term_S;
                double termProg_NE = termVeri_NE - Dif_term_NE;
                double termProg_N = termVeri_N - Dif_term_N;

                int Eolica_SE = 0;
                int Eolica_S = Convert.ToInt32(myWorksheet.Cells["O45"].Value.ToString());
                int Eolica_NE = Convert.ToInt32(myWorksheet.Cells["O29"].Value.ToString());
                int Eolica_N = Convert.ToInt32(myWorksheet.Cells["O21"].Value.ToString());

                int prodSolar_SE = Convert.ToInt32(myWorksheet.Cells["O38"].Value.ToString());
                int prodSolar_S = 0;
                int prodSolar_NE = Convert.ToInt32(myWorksheet.Cells["O30"].Value.ToString());
                int prodSolar_N = 0;

                double carga_SE = Math.Round(Convert.ToDouble(myWorksheet.Cells["O40"].Value.ToString()), 0);
                double carga_S = Math.Round(Convert.ToDouble(myWorksheet.Cells["O48"].Value.ToString()), 0);
                double carga_NE = Math.Round(Convert.ToDouble(myWorksheet.Cells["O32"].Value.ToString()), 0);
                double carga_N = Math.Round(Convert.ToDouble(myWorksheet.Cells["O24"].Value.ToString()), 0);

                double ReserMax_SE = Convert.ToDouble(myWorksheet.Cells["M73"].Value.ToString());
                double ReserMax_S = Convert.ToDouble(myWorksheet.Cells["M72"].Value.ToString());
                double ReserMax_NE = Convert.ToDouble(myWorksheet.Cells["M71"].Value.ToString());
                double ReserMax_N = Convert.ToDouble(myWorksheet.Cells["M70"].Value.ToString());

                double ReserAtual_SE = Convert.ToDouble(myWorksheet.Cells["Q65"].Value.ToString());
                double ReserAtual_S = Convert.ToDouble(myWorksheet.Cells["Q64"].Value.ToString());
                double ReserAtual_NE = Convert.ToDouble(myWorksheet.Cells["Q63"].Value.ToString());
                double ReserAtual_N = Convert.ToDouble(myWorksheet.Cells["Q62"].Value.ToString());

                double ReserPorc_SE = Convert.ToDouble(myWorksheet.Cells["R65"].Value.ToString()) <= 1 ? Convert.ToDouble(myWorksheet.Cells["R65"].Value.ToString()) * 100 : Convert.ToDouble(myWorksheet.Cells["R65"].Value.ToString());
                double ReserPorc_S = Convert.ToDouble(myWorksheet.Cells["R64"].Value.ToString()) <= 1 ? Convert.ToDouble(myWorksheet.Cells["R64"].Value.ToString()) * 100 : Convert.ToDouble(myWorksheet.Cells["R64"].Value.ToString());
                double ReserPorc_NE = Convert.ToDouble(myWorksheet.Cells["R63"].Value.ToString()) <= 1 ? Convert.ToDouble(myWorksheet.Cells["R63"].Value.ToString()) * 100 : Convert.ToDouble(myWorksheet.Cells["R63"].Value.ToString());
                double ReserPorc_N = Convert.ToDouble(myWorksheet.Cells["R62"].Value.ToString()) <= 1 ? Convert.ToDouble(myWorksheet.Cells["R62"].Value.ToString()) * 100 : Convert.ToDouble(myWorksheet.Cells["R62"].Value.ToString());


                IDB objSQL = new SQLServerDBCompass(banco, "IPDO");
                string[] campos = { "[id_ipdo]", "[num_sub]", "[prod_hidro_veri]", "[prod_term_veri]", "[prod_term_prog]", "[prod_eolica_veri]", "[carga_veri]", "[reservatorio_max]", "[reservatorio_atual]", "[reservatorio_atual_porc]", "[ENA]", "[term_el]", "[term_en]", "[term_in]", "[term_ex]", "[term_te]", "[term_ge]", "[term_pe]", "[term_gfom]", "[term_gsub]", "[term_er]", "[term_null]", "[prod_solar_veri]", "[term_rro]", "[term_ucm]", "[term_gen]" };
                object[,] valores = new object[1, 26]    {
                                                        {
                                                            id_ipdo,
                                                            1,
                                                            hidro_SE,
                                                            termVeri_SE,
                                                            termProg_SE,
                                                            Eolica_SE,
                                                            carga_SE,
                                                            ReserMax_SE,
                                                            ReserAtual_SE,
                                                            ReserPorc_SE,
                                                            ENA_SE,
                                                            termo[1,1],
                                                            termo[1,2],
                                                            termo[1,3],
                                                            termo[1,4],
                                                            termo[1,5],
                                                            termo[1,6],
                                                            termo[1,7],
                                                            termo[1,8],
                                                            termo[1,9],
                                                            termo[1,10],
                                                            termo[1,11],
                                                            prodSolar_SE,
                                                            termo[1,12],
                                                            termo[1,13],
                                                            termo[1,14]


                                                        }
                                                    };
                object[,] valores2 = new object[1, 26]    {
                                                        {
                                                            id_ipdo,
                                                            2,
                                                            hidro_S,
                                                            termVeri_S,
                                                            termProg_S,
                                                            Eolica_S,
                                                            carga_S,
                                                            ReserMax_S,
                                                            ReserAtual_S,
                                                            ReserPorc_S,
                                                            ENA_S,
                                                            termo[2,1],
                                                            termo[2,2],
                                                            termo[2,3],
                                                            termo[2,4],
                                                            termo[2,5],
                                                            termo[2,6],
                                                            termo[2,7],
                                                            termo[2,8],
                                                            termo[2,9],
                                                            termo[2,10],
                                                            termo[2,11],
                                                            prodSolar_S,
                                                            termo[2,12],
                                                            termo[1,13],
                                                            termo[1,14]


                                                        }
                                                    };

                object[,] valores3 = new object[1, 26]    {
                                                        {
                                                            id_ipdo,
                                                            3,
                                                            hidro_NE,
                                                            termVeri_NE,
                                                            termProg_NE,
                                                            Eolica_NE,
                                                            carga_NE,
                                                            ReserMax_NE,
                                                            ReserAtual_NE,
                                                            ReserPorc_NE,
                                                            ENA_NE,
                                                            termo[3,1],
                                                            termo[3,2],
                                                            termo[3,3],
                                                            termo[3,4],
                                                            termo[3,5],
                                                            termo[3,6],
                                                            termo[3,7],
                                                            termo[3,8],
                                                            termo[3,9],
                                                            termo[3,10],
                                                            termo[3,11],
                                                            prodSolar_NE,
                                                            termo[3,12],
                                                            termo[1,13],
                                                            termo[1,14]


                                                        }
                                                    };

                object[,] valores4 = new object[1, 26]    {
                                                        {
                                                            id_ipdo,
                                                            4,
                                                            hidro_N,
                                                            termVeri_N,
                                                            termProg_N,
                                                            Eolica_N,
                                                            carga_N,
                                                            ReserMax_N,
                                                            ReserAtual_N,
                                                            ReserPorc_N,
                                                            ENA_N,
                                                            termo[4,1],
                                                            termo[4,2],
                                                            termo[4,3],
                                                            termo[4,4],
                                                            termo[4,5],
                                                            termo[4,6],
                                                            termo[4,7],
                                                            termo[4,8],
                                                            termo[4,9],
                                                            termo[4,10],
                                                            termo[4,11],
                                                            prodSolar_N,
                                                            termo[4,12],
                                                            termo[1,13],
                                                            termo[1,14]


                                                        }
                                                    };

                string tabela = "[dbo].[IPDO_Submercado]";
                objSQL.Insert(tabela, campos, valores);
                objSQL.Insert(tabela, campos, valores2);
                objSQL.Insert(tabela, campos, valores3);
                objSQL.Insert(tabela, campos, valores4);





            }
        }

        public static void carregaUEE(string PathExcel)
        {

            string FileName = PathExcel.Split('\\').Last();


            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(PathExcel)))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets["Totais"]; //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                var sb = new StringBuilder(); //this is your data 

                string Submercado = myWorksheet.Cells["C88"].Value.ToString();
                for (int i = 102; i <= 106; i++)
                {
                    int Ano = Convert.ToInt32(myWorksheet.Cells["C" + i].Value.ToString());
                    var janeiro = myWorksheet.Cells["D" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["D" + i].Value));

                    var fevereiro = myWorksheet.Cells["E" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["E" + i].Value));

                    var marco = myWorksheet.Cells["F" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["F" + i].Value));
                    var abril = myWorksheet.Cells["G" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["G" + i].Value));
                    var maio = myWorksheet.Cells["H" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["H" + i].Value));
                    var junho = myWorksheet.Cells["I" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["I" + i].Value));
                    var julho = myWorksheet.Cells["J" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["J" + i].Value));
                    var agosto = myWorksheet.Cells["K" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["K" + i].Value));
                    var setembro = myWorksheet.Cells["L" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["L" + i].Value));
                    var outubro = myWorksheet.Cells["M" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["M" + i].Value));
                    var novembro = myWorksheet.Cells["N" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["N" + i].Value));
                    var dezembro = myWorksheet.Cells["O" + i].Value.ToString() == "" ? 0 : Math.Round(Convert.ToDouble(myWorksheet.Cells["O" + i].Value));
                    var data = DateTime.Today;


                    // Dados Tabela UEE
                    object[,] valores_UEE = new object[1, 15]{
                                                        {  data,
                                                           Submercado,
                                                           Ano,
                                                           janeiro,
                                                           fevereiro,
                                                           marco,
                                                           abril,
                                                           maio,
                                                           junho,
                                                           julho,
                                                           agosto,
                                                           setembro,
                                                           outubro,
                                                           novembro,
                                                           dezembro

                                                        }
                                                    };


                    IDB objSQL = new SQLServerDBCompass("ESTUDO_PV");

                    string[] campos = { "[Data]", "[submercado]", "[Ano]", "[Janeiro]", "[Fevereiro]", "[Março]", "[Abril]", "[Maio]", "[Junho]", "[Julho]", "[Agosto]", "[Setembro]", "[Outubro]", "[Novembro]", "[Dezembro]" };

                    string tabela = "[ESTUDO_PV].[dbo].[UEE]";
                    objSQL.Insert(tabela, campos, valores_UEE);

                }


            }
        }
    }
}
