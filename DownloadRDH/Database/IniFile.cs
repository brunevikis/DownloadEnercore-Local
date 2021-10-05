using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DownloadCompass.Database
{
    /// <summary>
	/// Classe para leitura de arquivos INI.
	/// </summary>
	public class IniFile
    {
        private string m_strFilePath;
        private string m_strFolder;

        /// <summary>
        /// Caminho do arquivo
        /// </summary>
        public string FilePath
        {
            get { return this.m_strFilePath; }
            set
            {
                this.m_strFilePath = value;
                this.m_strFolder = Path.GetDirectoryName(value);
            }
        }

        /// <summary>
        /// Pasta do arquivo
        /// </summary>
        public string Folder
        {
            get { return this.m_strFolder; }
        }

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section,
        string key,
        string val,
        string filePath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section,
        string key,
        string def,
        StringBuilder retVal,
        int size,
        string filePath);

        public IniFile(string filePath)
        {
            this.FilePath = filePath;
            CheckFile();
        }

        /// <summary>
        /// Escreve uma informação no arquivo INI e salva
        /// </summary>
        /// <param name="section">[Seção]</param>
        /// <param name="key">Chave</param>
        /// <param name="value">Valor</param>
        public void Write(string section, string key, string value)
        {
            CheckFile();
            WritePrivateProfileString(section, key, value.ToLower(), this.m_strFilePath);
        }

        /// <summary>
        /// Lê um valor no arquivo
        /// </summary>
        /// <param name="section">[Seção]</param>
        /// <param name="key">Chave</param>
        /// <returns></returns>
        public string Read(string section, string key)
        {
            CheckFile();
            StringBuilder SB = new StringBuilder(255);
            int i = GetPrivateProfileString(section, key, "", SB, 255, this.m_strFilePath);
            return SB.ToString();
        }

        /// <summary>
        /// Lê um valor inteiro
        /// </summary>
        /// <param name="section">[Seção]</param>
        /// <param name="key">Chave</param>
        /// <returns>O inteiro correpondente, ou int.MinValue se não existir</returns>
        public int ReadInt(string section, string key)
        {
            int retorno = int.MinValue;
            int.TryParse(Read(section, key), out retorno);
            return retorno;
        }

        /// <summary>
        /// Dispara uma FileNotFoundException se o arquivo não existir
        /// </summary>
        public void CheckFile()
        {
            if (!File.Exists(this.m_strFilePath))
            {
                throw new FileNotFoundException("O Arquivo " + this.m_strFilePath + " não foi encontrado !!", this.m_strFilePath);
            }
        }

    }
}
