using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DownloadCompass
{
    class ExcelParserRDH
    {
        #region Public Methods
        public static void PreencherGeral(string p_strArquivoParaProcessar, RDHInfo objRDHInfoEmProcessamento)
        {
            objRDHInfoEmProcessamento.DataRDH = ExcelParserRDH.GetDataRDH(p_strArquivoParaProcessar);
            objRDHInfoEmProcessamento.ArquivoOriginal = p_strArquivoParaProcessar;
        }

        public static void PreencherHidroenergeticaSubsistemas(String p_strPathOrigem, RDHInfo p_objRDHInfo)
        {
            Microsoft.Office.Interop.Excel.Application objExcel;
            Microsoft.Office.Interop.Excel.Workbook objWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet objWorksheet;

            objExcel = new Microsoft.Office.Interop.Excel.Application();

            objWorkbook = objExcel.Workbooks.Open(p_strPathOrigem);
            objWorksheet = objWorkbook.Worksheets[1];

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in objWorkbook.Worksheets)
            {
                if (sheet.Name.ToLower().Replace("é", "e").Replace("-", "") == "hidroenergeticasubsistemas")
                {
                    objWorksheet = sheet;
                    break;
                }
            }

            ExcelParserRDH.PreencherHidroenergeticaSubsistemas(objWorksheet, p_objRDHInfo);

            objWorkbook.Close(false);

            objExcel.Quit();

            releaseObject(objWorksheet);
            releaseObject(objWorkbook);
        }

        /// <summary>
        /// Classe capaz de, recebendo um Worksheet, extrair o conteúdo da aba Hidroenergérico-Subsistemas e preencher um objeto do tipo RDHInfo.
        /// </summary>
        /// <param name="objWorksheet">Objeto do tipo Worksheet referente à aba Hidroenergérico-Subsistemas do RDH.</param>
        /// <param name="p_objRDHInfo">Objeto já instanciado para preencher com as informações extraídas.</param>
        public static void PreencherHidroenergeticaSubsistemas(Microsoft.Office.Interop.Excel.Worksheet objWorksheet, RDHInfo p_objRDHInfo)
        {
            string col1 = string.Empty;
            RDHHidroSubData data;

            DateTime dtConsiderada;

            //var teste = objWorksheet.Range["J2"].Value.ToString();

            string strDataConsiderada;
            strDataConsiderada = objWorksheet.Range["J2"].Value.ToString(); // antes estava J4, estava dando referencia nulla
            strDataConsiderada = strDataConsiderada.Split(' ')[2];

            dtConsiderada = DateTime.Parse(strDataConsiderada);

            for (int i = 1; i < 100; i++)
            {

                data = null;
                col1 = "";

                if (objWorksheet.Range["A" + i.ToString()].Value == null)
                {
                    continue;
                }

                col1 = objWorksheet.Range["A" + i.ToString()].Value.ToString();



                if (col1.Length > 2)
                {
                    if (col1.ToLower().Contains("sudeste"))
                    {
                        data = p_objRDHInfo.GetHidroSubData(Enums.Submercado.SudesteCentroOeste);
                        data.Submercado = Enums.Submercado.SudesteCentroOeste;
                    }
                    else if (col1.ToLower().Contains("sul"))
                    {
                        data = p_objRDHInfo.GetHidroSubData(Enums.Submercado.Sul);
                        data.Submercado = Enums.Submercado.Sul;
                    }
                    else if (col1.ToLower().Contains("nordeste"))
                    {
                        data = p_objRDHInfo.GetHidroSubData(Enums.Submercado.Nordeste);
                        data.Submercado = Enums.Submercado.Nordeste;
                    }
                    else if (col1.ToLower().Contains("norte"))
                    {
                        data = p_objRDHInfo.GetHidroSubData(Enums.Submercado.Norte);
                        data.Submercado = Enums.Submercado.Norte;
                    }
                    else
                        continue;


                    /*var teste = col1.Split('/').First();
                    switch (teste.ToLower())
                    {
                        case "sudeste":

                            break;


                        case "sul":

                            break;

                        case "nordeste":

                            break;

                        case "norte":

                            break;

                        default:
                            continue;
                    }*/

                    data.Reservatorio_Dia = ToDouble(objWorksheet.Range["L" + (i + 3).ToString()].Value.ToString());
                    data.TotalMW_MesAteData = Convert.ToInt32(objWorksheet.Range["H" + (i + 3).ToString()].Value.ToString());
                    data.TotalMW_MediaSemanaAteData = Convert.ToInt32(objWorksheet.Range["G" + (i + 3).ToString()].Value.ToString());

                    data.Data = dtConsiderada;

                    string strDiasSemanaConsiderados = objWorksheet.Range["G" + (i + 2).ToString()].Value.ToString();
                    strDiasSemanaConsiderados = strDiasSemanaConsiderados.Replace(" ", "");

                    string[] arrSemanas = strDiasSemanaConsiderados.Split('-');
                    string[] arrDataInicio = arrSemanas[0].Split('/');
                    string[] arrDataFim = arrSemanas[1].Split('/');

                    data.SemanaAtualFim = DateTime.Parse(arrSemanas[1] + "/" + dtConsiderada.Year);

                    if (Convert.ToInt32(arrDataInicio[1]) > Convert.ToInt32(arrDataFim[1]))
                    {
                        data.SemanaAtualInicio = DateTime.Parse(arrSemanas[0] + "/" + (dtConsiderada.Year - 1).ToString());
                    }
                    else
                    {
                        data.SemanaAtualInicio = DateTime.Parse(arrSemanas[0] + "/" + dtConsiderada.Year.ToString());
                    }
                }
            }
        }

        public static void PreencherHidraulicoHidrologica(String p_strPathOrigem, RDHInfo p_objRDHInfo)
        {
            Microsoft.Office.Interop.Excel.Application objExcel;
            Microsoft.Office.Interop.Excel.Workbook objWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet objWorksheet;

            objExcel = new Microsoft.Office.Interop.Excel.Application();

            objWorkbook = objExcel.Workbooks.Open(p_strPathOrigem, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            objWorksheet = objWorkbook.Worksheets[1];

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in objWorkbook.Worksheets)
            {
                if (sheet.Name.ToLower().Replace("á", "a").Replace("ó", "o").Replace("-", "") == "hidraulicohidrologica")
                {
                    objWorksheet = sheet;
                    break;
                }
            }

            ExcelParserRDH.PreencherHidraulicoHidrologica(objWorksheet, p_objRDHInfo);

            objWorkbook.Close(false);

            objExcel.Quit();

            releaseObject(objWorksheet);
            releaseObject(objWorkbook);
        }

        public static void PreencherHidraulicoHidrologica(Microsoft.Office.Interop.Excel.Worksheet objWorksheet, RDHInfo p_objRDHInfo)
        {
            // Pega a data na célula U2
            string strData = objWorksheet.Range["U2"].Value.ToString();
            string strDataExtraida = "";
            foreach (char c in strData)
            {
                if (char.IsDigit(c) || c == '/' || c == '-')
                {
                    strDataExtraida += c;
                }
            }
            DateTime dtConsiderada = Convert.ToDateTime(strDataExtraida);

            // Varre a coluna "E" (id dos postos) procurando por um número maior que 0 e menor que 1000, e descobre o range de linhas a subir para a memória.

            int intLinhaFirst = 0, intLinhaLast = 0;

            object objAuxProcuraLinha;
            int intAuxProcuraLinha = 100;
            for (int i = 1; i <= 100; i++)
            {
                objAuxProcuraLinha = objWorksheet.Range["E" + i.ToString()].Value;
                if (objAuxProcuraLinha != null && int.TryParse(objAuxProcuraLinha.ToString(), out intAuxProcuraLinha))
                {
                    if (intAuxProcuraLinha > 0 && intAuxProcuraLinha < 1000)
                    {
                        // Achei!
                        intAuxProcuraLinha = i;
                        break;
                    }
                }
            }

            // se não parou no break, certamente saiu do for com intAuxProcuraPrimeiraLinha valendo 100. Se for o caso, aborta.
            if (intAuxProcuraLinha == 100) return;
            intLinhaFirst = intAuxProcuraLinha;

            for (int i = intLinhaFirst; i <= 1000; i++)
            {
                objAuxProcuraLinha = objWorksheet.Range["E" + i.ToString()].Value;
                if (
                        objAuxProcuraLinha == null ||
                        (
                            objAuxProcuraLinha.ToString().ToLower().Trim() != "nd" &&
                            !int.TryParse(objAuxProcuraLinha.ToString(), out intAuxProcuraLinha)
                        )
                    )
                {
                    if (intAuxProcuraLinha < 1000)
                    {
                        // Achei!
                        intAuxProcuraLinha = i;
                        break;
                    }
                }
            }

            if (intAuxProcuraLinha == 1000) return;
            intLinhaLast = intAuxProcuraLinha - 1;

            object[,] arrDados = objWorksheet.Range[String.Format("E{0}:AB{1}", intLinhaFirst, intLinhaLast)].Value;

            for (int row = 1; row <= arrDados.GetLength(0); row++)
            {
                if (IsInteger(arrDados[row, 1]))
                {
                    RDHHidraHidroData objData = new RDHHidraHidroData();
                    objData.Data = dtConsiderada;
                    objData.Posto = Convert.ToInt32(arrDados[row, 1]);
                    objData.VazaoNaturalDia = ToIntSmart(arrDados[row, 10]);
                    objData.VazaoNaturalUltMax = ToIntSmart(arrDados[row, 6]);
                    objData.VazaoNaturalUltMin = ToIntSmart(arrDados[row, 8]);
                    objData.EnergiaArmazenada = ToDoubleSmart(arrDados[row, 12]);
                    objData.VolumeEspera = ToDoubleSmart(arrDados[row, 13]);
                    objData.VazaoDefluente = ToDoubleSmart(arrDados[row, 17]);
                    objData.VazaoIncremental = ToDoubleSmart(arrDados[row, 20]);
                    p_objRDHInfo.HidraulicoHidrologica.Add(objData.Posto, objData);
                }
            }
        }
        #endregion

        #region Private Methods

        private static double ToDouble(string p_strNumber)
        {
            string sep = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            return Convert.ToDouble(p_strNumber.Replace(",", sep).Replace(".", sep));
        }

        private static DateTime GetDataRDH(string p_strCaminhoArquivo)
        {
            string strData = ExcelParserRDH.GetCellValue(p_strCaminhoArquivo, "", "U2").ToString();
            string strDataExtraida = "";

            foreach (char c in strData)
            {
                if (char.IsDigit(c) || c == '/' || c == '-')
                {
                    strDataExtraida += c;
                }
            }

            return Convert.ToDateTime(strDataExtraida);
        }

        private static object GetCellValue(string p_strCaminhoArquivo, string p_strNomeWorksheet, string p_strCelula)
        {
            Microsoft.Office.Interop.Excel.Application objExcel;
            Microsoft.Office.Interop.Excel.Workbook objWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet objWorksheet = null;

            objExcel = new Microsoft.Office.Interop.Excel.Application();

            objWorkbook = objExcel.Workbooks.Open(p_strCaminhoArquivo, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            if (p_strNomeWorksheet == null || p_strNomeWorksheet == "")
            {
                objWorksheet = objWorkbook.Worksheets[1];
            }
            else
            {
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in objWorkbook.Worksheets)
                {
                    if (sheet.Name.ToLower().Replace("_", "").Replace(" ", "").Replace("-", "") == p_strNomeWorksheet.ToLower().Replace("_", "").Replace("", " ").Replace("-", ""))
                    {
                        objWorksheet = sheet;
                        break;
                    }
                }
            }

            object objRetorno = null;

            if (p_strCelula.Contains(';'))
            {
                objRetorno = new List<object>();
                foreach (string strCelula in p_strCelula.Split(';'))
                {
                    if (strCelula != "")
                    {
                        ((List<object>)objRetorno).Add(objWorksheet.Range[strCelula].Value);
                    }
                }
            }
            else
            {
                objRetorno = objWorksheet.Range[p_strCelula].Value;
            }

            objWorkbook.Close(false);

            objExcel.Quit();

            releaseObject(objWorksheet);
            releaseObject(objWorkbook);

            return objRetorno;
        }

        private async static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static bool IsInteger(object obj)
        {
            if (obj == null) return false;
            int i;
            return int.TryParse(obj.ToString(), out i);
        }
        private static int ToIntSmart(object obj)
        {
            if (obj == null) return int.MinValue;
            int i = 0; ;
            if (!int.TryParse(obj.ToString(), out i))
            {
                return int.MinValue;
            }
            return i;
        }
        private static double ToDoubleSmart(object obj)
        {
            if (obj == null) return double.MinValue;
            double i = 0; ;
            if (!double.TryParse(obj.ToString(), out i))
            {
                return double.MinValue;
            }
            return i;
        }
        #endregion
    }
}
