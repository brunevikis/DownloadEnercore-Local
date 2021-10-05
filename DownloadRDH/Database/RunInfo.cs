using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace DownloadCompass.Database
{
    public static class RunInfo
    {
        #region Atributos Privados

        private static string m_strTitulo;
        private static string m_strExecutar;
        private static string m_strIconeArquivo;
        private static Image m_objIconeImagem;
        private static NivelAmbiente m_objNivelAmbiente = NivelAmbiente.Indefinido;
        private static int m_intVersao;
        private static List<string> m_arrArquivosIgnorados = new List<string>();

        private static IniFile m_objIniFile;

        #endregion

        #region Propriedades públicas

        /// <summary>
        /// Título da aplicação
        /// </summary>
        public static string Titulo { get { return m_strTitulo; } }

        /// <summary>
        /// Arquivo a executar
        /// </summary>
        public static string Executar { get { return m_strExecutar; } }

        /// <summary>
        /// Nome do arquivo de ícone
        /// </summary>
        public static string IconePath { get { return m_strIconeArquivo; } }

        /// <summary>
        /// Objeto Image que contém a imagem do ícone
        /// </summary>
        public static Image IconeImagem { get { return m_objIconeImagem; } }

        /// <summary>
        /// Ambiente de execução: DEV, QA, PROD
        /// </summary>
        public static NivelAmbiente Ambiente { get { return m_objNivelAmbiente; } }

        /// <summary>
        /// Versão em execução
        /// </summary>
        public static int Versao { get { return m_intVersao; } }

        /// <summary>
        /// Arquivos que o processo de execução não copia para a pasta local
        /// </summary>
        public static IList<string> ArquivosIgnorados { get { return m_arrArquivosIgnorados.AsReadOnly(); } }

        #endregion

        #region Métodos públicos

        public static string Read(string section, string key)
        {
            return m_objIniFile.Read(section, key);
        }

        public static int ReadInt(string section, string key)
        {
            return m_objIniFile.ReadInt(section, key);
        }

        #endregion

        #region Contrução

        /// <summary>
        /// Cria uma instância de RunInfo que tenta apontar para um arquivo "run.ini", na pasta do executável.
        /// </summary>
        static RunInfo()
        {
            m_objIniFile = new IniFile(DownloadCompass.Database.Ambiente.ApplicationExePath + "\\run.ini");
            CarregarInformacoes();
        }

        /// <summary>
        /// Carrega as informãções do arquivo para o qual este objeto aponta
        /// </summary>
        private static void CarregarInformacoes()
        {
            m_strTitulo = m_objIniFile.Read("run", "titulo");
            m_strExecutar = m_objIniFile.Read("run", "executar");
            m_strIconeArquivo = m_objIniFile.Read("run", "icone");

            if (File.Exists(m_objIniFile.Folder + "\\" + m_strIconeArquivo))
            {
                m_objIconeImagem = Image.FromFile(m_objIniFile.Folder + "\\" + m_strIconeArquivo);
            }

            switch (m_objIniFile.Read("run", "ambiente").ToLower())
            {
                case "dev":
                    m_objNivelAmbiente = RunInfo.NivelAmbiente.DEV;
                    break;

                case "qa":
                    m_objNivelAmbiente = RunInfo.NivelAmbiente.QA;
                    break;

                case "prod":
                    m_objNivelAmbiente = RunInfo.NivelAmbiente.PROD;
                    break;

                default:
                    m_objNivelAmbiente = RunInfo.NivelAmbiente.Indefinido;
                    break;
            }

            if (!int.TryParse(m_objIniFile.Read("versao", "versao"), out m_intVersao))
            {
                m_intVersao = 0;
            }

            foreach (string strIgnorado in m_objIniFile.Read("run", "ignorar").Split(';'))
            {
                if (strIgnorado.Trim() != "")
                {
                    m_arrArquivosIgnorados.Add(strIgnorado);
                }
            }
        }

        #endregion

        /// <summary>
        /// Diferentes tipos de ambiente de execução
        /// </summary>
        public enum NivelAmbiente
        {
            Indefinido = 0,
            DEV = 1,
            QA = 2,
            PROD = 3
        }
    }
}
