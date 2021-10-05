using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownloadCompass.DB
{
    public class SQLServerDBCompass : SQLServerDB
    {
        public SQLServerDBCompass(string banco = "local",string bd = "IPDO")
        {
            string endereco = "10.206.194.187";
            if (banco == "local")
            {
                endereco = "10.206.194.187";
                this.SetUsuario("sa");


            }
            else if(banco == "azure"){

                endereco = "bdcompass.database.windows.net";
                this.SetUsuario("compass");
            }

            this.SetServidor(endereco);
            //this.SetPorta(1433);
            
            this.SetSenha(this.GetPassword(this.GetUsuario()));
            if (bd != "IPDO")
            {
                this.SetDatabase(this.GetDatabase(bd));

            }
            else
            {
                this.SetDatabase(this.GetDatabase("IPDO"));
            }
        }

        private string GetPassword(string p_strUsuario)
        {
            switch (p_strUsuario)
            {
                case "sa":
                    return "cp@s9876";
                case "compass":
                    return "cpas#9876";
                case "captura":
                    return "c@ptura9876";

                case "captura_read":
                    return "captur@leitur@";

                default:
                    return "";
            }
        }

        private string GetDatabase(string p_banco)
        {
            switch (p_banco)
            {
                case "IPDO":
                    return "IPDO";
                case "ESTUDO_PV":
                    return "ESTUDO_PV";
                case "CHUVAS":
                    return "CHUVAS";



                default:
                    return "";
            }
        }
    }
}
