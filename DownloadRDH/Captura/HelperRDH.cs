using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DownloadCompass.Database;

namespace DownloadCompass
{
    public class HelperRDH
    {
        #region Public Methods
        public static void GravarRDHInfo(RDHInfo p_objRDHInfo)
        {
            try
            {
                GravarRDHGeral(p_objRDHInfo);

                List<RDHHidroSubData> arrListHidroSubData = new List<RDHHidroSubData>();
                arrListHidroSubData.Add(p_objRDHInfo.GetHidroSubData(Enums.Submercado.SudesteCentroOeste));
                arrListHidroSubData.Add(p_objRDHInfo.GetHidroSubData(Enums.Submercado.Sul));
                arrListHidroSubData.Add(p_objRDHInfo.GetHidroSubData(Enums.Submercado.Nordeste));
                arrListHidroSubData.Add(p_objRDHInfo.GetHidroSubData(Enums.Submercado.Norte));

                GravarRDHHidroSubData(arrListHidroSubData);

                GravarRDHHidraHidroData(p_objRDHInfo.HidraulicoHidrologica);

                //p_objRDHInfo.e
            }
            catch(Exception e)
            {
                Console.WriteLine("O RDH deu pau aqui: " + e.Message);
            }
        }

        #endregion

        #region Private Methods

        private static void GravarRDHHidroSubData(List<RDHHidroSubData> p_arrDados)
        {
            IDB objSQL = new MySQLDBCompass();
            string[] campos = { "dt_rdh", "id_submercado", "perc_reservatorio_dia", "mw_semana", "mw_mes", "dt_semana_inicio", "dt_semana_fim", "dt_atualizacao" };
            object[,] valores = new object[p_arrDados.Count, 8];

            int i = 0;
            int j;
            foreach (RDHHidroSubData objDados in p_arrDados)
            {
                j = -1;
                valores[i, ++j] = objDados.Data;
                valores[i, ++j] = (int)objDados.Submercado;
                valores[i, ++j] = objDados.Reservatorio_Dia;
                valores[i, ++j] = objDados.TotalMW_MediaSemanaAteData;
                valores[i, ++j] = objDados.TotalMW_MesAteData;
                valores[i, ++j] = objDados.SemanaAtualInicio;
                valores[i, ++j] = objDados.SemanaAtualFim;
                valores[i, ++j] = DateTime.Now;
                i++;
            }

            string tabela = "fat_rdh_hidro_subsist";
            objSQL.Replace(tabela, campos, valores);
        }

        private static void GravarRDHHidraHidroData(Dictionary<int, RDHHidraHidroData> p_arrDados)
        {
            IDB objSQL = new MySQLDBCompass();
            string[] campos = { "dt_rdh", "id_posto", "vl_vazao_dia", "vl_vazao_ult_max", "vl_vazao_ult_min", "vl_earm", "vl_vol_espera", "vl_vazao_defluente", "vl_vazao_incremental", "dt_atualizacao" };
            object[,] valores = new object[p_arrDados.Count, 10];

            int i = 0;
            int j;
            foreach (int key in p_arrDados.Keys)
            {

                RDHHidraHidroData objDados = p_arrDados[key];

                j = -1;
                valores[i, ++j] = objDados.Data;
                valores[i, ++j] = (int)objDados.Posto;
                valores[i, ++j] = objDados.VazaoNaturalDia;
                valores[i, ++j] = objDados.VazaoNaturalUltMax;
                valores[i, ++j] = objDados.VazaoNaturalUltMin;
                valores[i, ++j] = objDados.EnergiaArmazenada;
                valores[i, ++j] = objDados.VolumeEspera;
                valores[i, ++j] = objDados.VazaoDefluente;
                valores[i, ++j] = objDados.VazaoIncremental;
                valores[i, ++j] = DateTime.Now;
                i++;
            }

            string tabela = "fat_rdh_hidra_hidro";
            objSQL.Replace(tabela, campos, valores);
        }
        private static void GravarRDHGeral(RDHInfo p_objRDHInfo)
        {
                IDB objSQL = new MySQLDBCompass();
            string[] campos = { "dt_rdh", "ds_arquivo", "dt_atualizacao" };
            object[,] valores = new object[1, 3]    {
                                                        {
                                                            p_objRDHInfo.DataRDH,
                                                            p_objRDHInfo.ArquivoOriginal,
                                                            DateTime.Now
                                                        }
                                                    };
            string tabela = "fat_rdh_carga";
            objSQL.Replace(tabela, campos, valores);
            //objSQL.Execute(p_objRDHInfo);
        }
        #endregion

    }
}
