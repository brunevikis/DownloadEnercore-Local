using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;
using System.Data.SqlClient;
using System.Collections;
using System.Data;
using System.Windows.Forms;


namespace DownloadCompass.DB
{
    public class SQLServerDB : IDB
    {

        private string m_strServidor = "";
        private int m_intPorta = 0;
        private string m_strDatabase = "";
        private string m_strUsuario = "";
        private string m_strSenha = "";


        protected void SetServidor(string host) { this.m_strServidor = host; }
        //  protected void SetPorta(int porta) { this.m_intPorta = porta; }
        protected void SetDatabase(string database) { this.m_strDatabase = database; }
        protected void SetUsuario(string usuario) { this.m_strUsuario = usuario; }
        protected void SetSenha(string senha) { this.m_strSenha = senha; }

        public string GetServidor() { return m_strServidor; }
        // public int GetPorta() { return m_intPorta; }
        public string GetDatabase() { return m_strDatabase; }
        public string GetUsuario() { return m_strUsuario; }
        public string GetSenha() { return m_strSenha; }

        /// <summary>
        /// Lista de palavras proibidas, para segurança contra intervenções mal-intencionadas.
        /// AINDA NÃOIMPLEMENTADA
        /// </summary>
        private static string[] blackList = {"--",";--",";","/*","*/","@@",
                                            " char"," nchar"," varchar"," nvarchar"," int",
                                            "alter ","begin ","cast","create ","cursor ","declare ","delete ","drop ","end","exec ","execute ",
                                            "fetch ","insert ","kill ","open ",
                                            "select ", " sys","sysobjects","syscolumns", "table ","update "};

        public string GetConnectionString()
        {
            return "server=" + GetServidor() + ";" +
                    //  "port=" + GetPorta() + ";" +
                    "database=" + GetDatabase() + ";" +
                    "uid=" + GetUsuario() + ";" +
                    "pwd=" + GetSenha();
        }

        public DbConnection GetConnection()
        {
            try
            {
                SqlConnection objConn = new SqlConnection(this.GetConnectionString());
                objConn.Open();
                return objConn;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha na conexão com o banco de dados:" + ex);
                return null;
            }
        }

        #region Leitura do banco de dados

        public object GetScalar(string p_strComandoSQL)
        {
            object objRetorno;
            DbConnection objConn = GetConnection();
            DbCommand objCmd = new SqlCommand(p_strComandoSQL, (SqlConnection)objConn);
            objRetorno = objCmd.ExecuteScalar();
            objConn.Close();
            return objRetorno;
        }

        #region Variações sobre GetScalar

        public string GetString(string p_strComandoSQL)
        {
            object objScalar = GetScalar(p_strComandoSQL);
            if (objScalar == null) return null;
            return objScalar.ToString();
        }

        public int GetInt(string p_strComandoSQL)
        {
            object objScalar = GetScalar(p_strComandoSQL);
            if (objScalar == null) return int.MinValue;
            return Convert.ToInt32(objScalar);
        }

        public double GetDouble(string p_strComandoSQL)
        {
            object objScalar = GetScalar(p_strComandoSQL);
            if (objScalar == null) return double.MinValue;
            return Convert.ToDouble(objScalar);
        }

        public DateTime GetDateTime(string p_strComandoSQL)
        {
            object objScalar = GetScalar(p_strComandoSQL);
            if (objScalar == null) return DateTime.MinValue;
            return Convert.ToDateTime(objScalar);
        }

        public bool GetBoolean(string p_strComandoSQL)
        {
            object objScalar = GetScalar(p_strComandoSQL);
            if (objScalar == null) return false;
            return Convert.ToBoolean(objScalar);
        }

        public byte GetByte(string p_strComandoSQL)
        {
            object objScalar = GetScalar(p_strComandoSQL);
            if (objScalar == null) return byte.MinValue;
            return Convert.ToByte(objScalar);
        }

        #endregion

        public List<List<object>> GetList(string p_strComandoSQL)
        {
            DbDataReader objReader = GetReader(p_strComandoSQL);
            if (objReader == null) return null;

            List<List<object>> arrRetorno = new List<List<object>>();

            while (objReader.Read())
            {
                List<object> objLinha = new List<object>(objReader.FieldCount);
                for (int i = 0; i < objReader.FieldCount; i++)
                {
                    objLinha.Add(objReader[i]);

                }
                arrRetorno.Add(objLinha);
            }

            objReader.Close();

            return arrRetorno;
        }

        public DataTable GetDataTable(string p_strComandoSQL)
        {
            SqlDataAdapter objAdapter = new SqlDataAdapter(p_strComandoSQL, this.GetConnectionString());
            DataTable objResult = new DataTable("table");
            objAdapter.Fill(objResult);
            return objResult;
        }

        public DbDataReader GetReader(string p_strComandoSQL)
        {
            DbDataReader objRetorno;
            DbConnection objConn = GetConnection();
            DbCommand objCmd = new SqlCommand(p_strComandoSQL, (SqlConnection)objConn);
            objRetorno = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
            return objRetorno;
        }

        #endregion

        #region Execuções sem resultado

        public int Execute(string p_strComandoSQL)
        {
            int intRetorno;
            DbConnection objConn = GetConnection();
            DbCommand objCmd = new SqlCommand(p_strComandoSQL, (SqlConnection)objConn);
            intRetorno = objCmd.ExecuteNonQuery();
            objConn.Close();
            return intRetorno;
        }

        #endregion

        #region Escrita no banco de dados

        public int Insert(string p_strTable, string[] p_arrColumns, object[,] p_arrValues)
        {
            return this.Execute(SQLServerQueryBuilder.BuildInsertQuery(p_strTable, p_arrColumns, p_arrValues));
        }

        public int Replace(string p_strTable, string[] p_arrColumns, object[,] p_arrValues)
        {
            return this.Execute(SQLServerQueryBuilder.BuildReplaceQuery(p_strTable, p_arrColumns, p_arrValues));
        }

        #endregion
    }
}
