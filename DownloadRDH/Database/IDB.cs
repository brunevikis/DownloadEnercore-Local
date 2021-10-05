using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data;

namespace DownloadCompass.Database
{
    public interface IDB
    {
        string GetServidor();
        int GetPorta();
        string GetDatabase();
        string GetUsuario();
        string GetSenha();

        string GetConnectionString();

        object GetScalar(string p_strComandoSQL);
        string GetString(string p_strComandoSQL);
        int GetInt(string p_strComandoSQL);
        double GetDouble(string p_strComandoSQL);
        DateTime GetDateTime(string p_strComandoSQL);
        bool GetBoolean(string p_strComandoSQL);
        byte GetByte(string p_strComandoSQL);

        List<List<object>> GetList(string p_strComandoSQL);
        DataTable GetDataTable(string p_strComandoSQL);
        DbDataReader GetReader(string p_strComandoSQL);

        int Execute(string p_strComandoSQL);

        int Insert(string p_strTable, string[] p_arrColumns, object[,] p_arrValues);
        int Replace(string p_strTable, string[] p_arrColumns, object[,] p_arrValues);

    }
}
