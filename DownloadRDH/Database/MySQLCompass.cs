using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownloadCompass.Database
{
    public class MySQLDBCompass : MySQLDB
    {
        public MySQLDBCompass()
        {
            this.SetServidor("192.168.0.41");
            this.SetPorta(3306); //1433
            this.SetUsuario("captura");
            this.SetSenha(this.GetPassword(this.GetUsuario()));
            this.SetDatabase("captura");

            /*switch (RunInfo.Ambiente)
            {
                case RunInfo.NivelAmbiente.DEV:
                    this.SetDatabase("dev_captura");
                    break;

                case RunInfo.NivelAmbiente.QA:
                    this.SetDatabase("qa_captura");
                    break;

                case RunInfo.NivelAmbiente.PROD:
                    this.SetDatabase("captura");
                    break;

                case RunInfo.NivelAmbiente.Indefinido:
                    break;
            }*/

        }

        private string GetPassword(string p_strUsuario)
        {
            switch (p_strUsuario)
            {
                case "captura":
                    return "c@ptura9876";

                case "captura_read":
                    return "captur@leitur@";

                default:
                    return "";
            }
        }
    }
}
