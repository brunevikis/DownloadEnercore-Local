using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownloadCompass.DB
{
    public class SQLServerQueryBuilder
    {
        private static string BuildInsertReplaceQuery(string p_strTable, string[] p_arrColumns, object[,] p_arrValues, bool p_blnReplace)
        {
            // Arrays:
            // http://msdn.microsoft.com/en-us/library/aa288453(v=vs.71).aspx

            // Utiliza o StringBuilder pois concatenação abusiva de string gasta mta memória, e nesse caso faz diferença.
            StringBuilder sb = new StringBuilder();

            if (p_blnReplace)
            {
                sb.AppendFormat("UPDATE \t{0}\n", p_strTable);
            }
            else
            {
                sb.AppendFormat("INSERT INTO \t{0}\n", p_strTable);
            }

            sb.AppendFormat("\t({0})\n", String.Join(",", p_arrColumns.ToArray()));
            sb.Append("VALUES\n");

            int row, col;
            object objValor;
            string tipo;

            for (row = 0; row < p_arrValues.GetLength(0); row++)
            {
                sb.Append("(");
                for (col = 0; col < p_arrValues.GetLength(1); col++)
                {
                    objValor = p_arrValues[row, col];
                    tipo = objValor.GetType().Name;

                    // verifica o tipo do valor, colocando aspas e formatando se necessário.
                    switch (tipo.ToLower())
                    {
                        case "string":
                            sb.AppendFormat("'{0}'", objValor.ToString().Replace("\\", "\\\\"));
                            break;

                        case "datetime":
                            sb.AppendFormat("'{0}'", ((DateTime)objValor).ToString("yyyy-M-d H:m:s"));
                            break;

                        case "int":
                            if ((int)objValor != int.MinValue)
                            {
                                sb.Append(objValor.ToString().Replace(',', '.'));
                            }
                            else
                            {
                                sb.Append("null");
                            }
                            break;


                        case "double":
                            if ((double)objValor != double.MinValue)
                            {
                                sb.Append(objValor.ToString().Replace(',', '.'));
                            }
                            else
                            {
                                sb.Append("null");
                            }
                            break;

                        case "float":
                            if ((float)objValor != float.MinValue)
                            {
                                sb.Append(objValor.ToString().Replace(',', '.'));
                            }
                            else
                            {
                                sb.Append("null");
                            }
                            break;

                        case "decimal":
                            if ((decimal)objValor != decimal.MinValue)
                            {
                                sb.Append(objValor.ToString().Replace(',', '.'));
                            }
                            else
                            {
                                sb.Append("null");
                            }
                            break;

                        default:
                            sb.Append(objValor.ToString());
                            break;
                    }

                    if (col != p_arrValues.GetLength(1) - 1)
                    {
                        sb.Append(",");
                    }
                }
                sb.Append(")");

                if (row != p_arrValues.GetLength(0) - 1)
                {
                    sb.Append(",");
                }

            }
            sb.Append(";");

            return sb.ToString();
        }

        public static string BuildInsertQuery(string p_strTable, string[] p_arrColumns, object[,] p_arrValues)
        {
            return BuildInsertReplaceQuery(p_strTable, p_arrColumns, p_arrValues, false);
        }

        public static string BuildReplaceQuery(string p_strTable, string[] p_arrColumns, object[,] p_arrValues)
        {
            return BuildInsertReplaceQuery(p_strTable, p_arrColumns, p_arrValues, true);
        }



    }
}
