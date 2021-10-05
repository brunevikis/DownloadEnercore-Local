using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices.AccountManagement;
using System.Security.Principal;
using System.IO;
using System.Windows.Forms;

namespace DownloadCompass.Database
{
    public static class Ambiente
    {

        #region Atributos privados

        /// <summary>
        /// "Cache" de grupos. Para atualiza, utilize o método UpdateGroups desta classe.
        /// </summary>
        private static List<String> m_arrUserGroups = null;

        #endregion

        /// <summary>
        /// Construtor estático. Ninguém chama, ele é executado no primeiro acesso que houver à classe.
        /// </summary>
        static Ambiente()
        {
            UpdateUserGroupsList();
        }

        #region Métodos privados

        private static string GetActiveDirectoryProperty(string p_strPropertyName)
        {
            System.DirectoryServices.DirectoryEntry ADEntry = new System.DirectoryServices.DirectoryEntry("WinNT://" + Ambiente.Domain + "/" + Ambiente.Username);
            return ADEntry.Properties[p_strPropertyName].Value.ToString();
        }

        private static List<String> GetGroups()
        {
            List<String> arrRetorno = new List<String>();
            PrincipalContext objDomain = new PrincipalContext(ContextType.Domain);
            UserPrincipal objUser = UserPrincipal.FindByIdentity(objDomain, Ambiente.Username);
            if (objUser != null)
            {
                PrincipalSearchResult<Principal> arrGrupos = objUser.GetAuthorizationGroups();
                foreach (Principal p in arrGrupos)
                {
                    if (p is GroupPrincipal)
                    {
                        arrRetorno.Add(p.Name);
                    }
                }
            }
            return arrRetorno;
        }

        #endregion

        #region Métodos públicos

        /// <summary>
        /// Atualiza a lista de grupos do usuário.
        /// </summary>
        public static void UpdateUserGroupsList()
        {
            if (m_arrUserGroups == null)
            {
                Ambiente.m_arrUserGroups = new List<string>();
            }
            else
            {
                Ambiente.m_arrUserGroups.Clear();
            }

            // TODO: Descomentar quando houver controle de privilégios por grupos
            //Ambiente.m_arrUserGroups.AddRange(Ambiente.GetGroups());

            // ------------------------------------------------------------------------------
            // TODO: comentar para colocar em produção. Está aqui para agilizar a inicialização do programa
            if (!Ambiente.m_arrUserGroups.Contains("Todos"))
            {
                Ambiente.m_arrUserGroups.Add("Todos");
            }
            // ------------------------------------------------------------------------------
        }

        /// <summary>
        /// Consulta se o usuário ativo tem acesso ao um determinado grupo. Este método não é case-sensitive.
        /// </summary>
        /// <param name="p_strGroupName">Nome do grupo</param>
        /// <returns></returns>
        public static bool UserHasGroup(string p_strGroupName)
        {
            string strGroupName = p_strGroupName.ToLower();
            foreach (string strGrupo in Ambiente.UserGroups)
            {
                if (strGrupo.ToLower().CompareTo(strGroupName) == 0)
                {
                    return true;
                }
            }

            return false;
        }

        #endregion

        #region Dados do usuário

        /// <summary>
        /// Nome completo do usuário
        /// </summary>
        public static string NomeCompleto
        {
            get { return Ambiente.GetActiveDirectoryProperty("FullName"); }
        }

        /// <summary>
        /// Nome de usuário
        /// </summary>
        public static string Username
        {
            get { return WindowsIdentity.GetCurrent().Name.Split('\\')[1]; }
        }

        /// <summary>
        /// Domínio do usuário logado
        /// </summary>
        public static string Domain
        {
            get { return WindowsIdentity.GetCurrent().Name.Split('\\')[0]; }
        }

        /// <summary>
        /// Lista de grupos aos quais o usuário pertence. Para atualizar a lista, chame o método UpdateGroups desta classe.
        /// </summary>
        public static List<String> UserGroups
        {
            get
            {
                if (Ambiente.m_arrUserGroups == null)
                {
                    Ambiente.UpdateUserGroupsList();
                }
                return m_arrUserGroups;
            }
        }

        #endregion

        #region Dados de máquina (nome, paths, etc)

        /// <summary>
        /// Nome do computador
        /// </summary>
        public static string MachineName
        {
            get { return System.Environment.MachineName; }
        }

        /// <summary>
        /// Diretório corrente, onde serão considerados os arquivos a abrir e salvar na ausência de definição de um path para tal.
        /// </summary>
        public static string CurrentDir
        {
            get { return System.Environment.CurrentDirectory; }
        }

        /// <summary>
        /// Caminho do executável da aplicação.
        /// </summary>
        public static string ApplicationExePath
        {
            get { return Path.GetDirectoryName(Application.ExecutablePath); }
        }

        /// <summary>
        /// Nome do arquivo executável da aplicação
        /// </summary>
        public static string ApplicationExeName
        {
            get { return Path.GetFileName(Application.ExecutablePath); }
        }

        /// <summary>
        /// Caminho para a pasta do Desktop do usuário, terminando sem barra. Geralmente algo como C:\Users\nome.de.usuario\Desktop
        /// </summary>
        public static string DesktopPath
        {
            get { return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); }
        }

        /// <summary>
        /// Caminho para a pasta raiz do usuário, terminando sem barra. Geralmente algo como C:\Users\nome.de.usuario
        /// </summary>
        public static string UserPath
        {
            get { return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile); }
        }

        /// <summary>
        /// Caminho para a pasta "Meus Documentos" do usuário, terminando sem barra. Geralmente algo como C:\Users\nome.de.usuario\Documents
        /// </summary>
        public static string MyDocumentsPath
        {
            get { return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); }
        }

        #endregion

        #region Dados da Aplicação

        private static DateTime m_dtAppBuild = DateTime.MinValue;
        public static DateTime ApplicationBuildDate
        {
            get
            {
                if (m_dtAppBuild == DateTime.MinValue)
                {

                    // Como extrair a data do build de um executável
                    // http://stackoverflow.com/questions/1600962/displaying-the-build-date

                    string filePath = System.Reflection.Assembly.GetCallingAssembly().Location;
                    const int c_PeHeaderOffset = 60;
                    const int c_LinkerTimestampOffset = 8;
                    byte[] b = new byte[2048];
                    System.IO.Stream s = null;

                    try
                    {
                        s = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                        s.Read(b, 0, 2048);
                    }
                    finally
                    {
                        if (s != null)
                        {
                            s.Close();
                        }
                    }

                    int i = System.BitConverter.ToInt32(b, c_PeHeaderOffset);
                    int secondsSince1970 = System.BitConverter.ToInt32(b, i + c_LinkerTimestampOffset);
                    DateTime dt = new DateTime(1970, 1, 1, 0, 0, 0);
                    dt = dt.AddSeconds(secondsSince1970);
                    dt = dt.AddHours(TimeZone.CurrentTimeZone.GetUtcOffset(dt).Hours);
                    m_dtAppBuild = dt;
                }

                return m_dtAppBuild;
            }
        }

        #endregion

        public static string GenerateTempFilePath(string p_strBaseName)
        {
            return Ambiente.UserPath + "\\" + p_strBaseName + string.Format("_{0}_{1}", DateTime.Now.ToString("yyyyMMdd_HHmmss"), new Random(DateTime.Now.Millisecond).Next(0, int.MaxValue).ToString());
        }

        public static string GenerateTempFilePath()
        {
            return Ambiente.GenerateTempFilePath(Path.GetFileNameWithoutExtension(Ambiente.ApplicationExeName).Replace(".vshost", ""));
        }

    }
}
