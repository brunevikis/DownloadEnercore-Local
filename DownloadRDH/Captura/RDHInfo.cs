using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownloadCompass
{
    public class RDHInfo
    {
        private Dictionary<Enums.Submercado, RDHHidroSubData> arrHidroSubData = new Dictionary<Enums.Submercado, RDHHidroSubData>();
        private Dictionary<int, RDHHidraHidroData> arrHidraHidroData = new Dictionary<int, RDHHidraHidroData>();

        private string m_strPathArquivoOriginal = "";
        private DateTime m_dtDataRDH = DateTime.MinValue;

        /// <summary>
        /// Arquivo que originou a carga
        /// </summary>
        public string ArquivoOriginal
        {
            get { return m_strPathArquivoOriginal; }
            set { this.m_strPathArquivoOriginal = value; }
        }

        /// <summary>
        /// Data do RDH
        /// </summary>
        public DateTime DataRDH
        {
            get { return m_dtDataRDH; }
            set { this.m_dtDataRDH = value; }
        }

        #region Hidráulico-Hidrológica

        [CategoryAttribute("Hidráulico-Hidrológica"), DescriptionAttribute("Dados extraídos da aba Hidroenergética Subsistemas - Submercado Norte"), Browsable(true)]
        public Dictionary<int, RDHHidraHidroData> HidraulicoHidrologica
        {
            get { return arrHidraHidroData; }
        }

        public RDHHidraHidroData GetHidroSubData(int p_intIdPosto)
        {
            return arrHidraHidroData[p_intIdPosto];
        }

        #endregion

        #region Hidroenergética-Subsistemas

        [Browsable(false)]
        public Dictionary<Enums.Submercado, RDHHidroSubData> HidroenergeticaSubsistemas
        {
            get { return arrHidroSubData; }
        }

        [CategoryAttribute("Hidroenergetica-Subsistemas"), DescriptionAttribute("Dados extraídos da aba Hidroenergética Subsistemas - Submercado Norte"), Browsable(true)]
        public RDHHidroSubData HidroSubN
        {
            get { return arrHidroSubData[Enums.Submercado.Norte]; }
        }

        [CategoryAttribute("Hidroenergetica-Subsistemas"), DescriptionAttribute("Dados extraídos da aba Hidroenergética Subsistemas - Submercado Nordeste"), Browsable(true)]
        public RDHHidroSubData HidroSubNE
        {
            get { return arrHidroSubData[Enums.Submercado.Nordeste]; }
        }

        [CategoryAttribute("Hidroenergetica-Subsistemas"), DescriptionAttribute("Dados extraídos da aba Hidroenergética Subsistemas - Enums.Submercado Sul"), Browsable(true)]
        public RDHHidroSubData HidroSubS
        {
            get { return arrHidroSubData[Enums.Submercado.Sul]; }
        }

        [CategoryAttribute("Hidroenergetica-Subsistemas"), DescriptionAttribute("Dados extraídos da aba Hidroenergética Subsistemas - Enums.Submercado Sudeste/Centro-Oeste"), Browsable(true)]
        public RDHHidroSubData HidroSubSE
        {
            get { return arrHidroSubData[Enums.Submercado.SudesteCentroOeste]; }
        }

        public RDHHidroSubData GetHidroSubData(Enums.Submercado p_objSubmercado)
        {
            return arrHidroSubData[p_objSubmercado];
        }

        #endregion




        public RDHInfo()
        {
            arrHidroSubData.Add(Enums.Submercado.Nordeste, new RDHHidroSubData());
            arrHidroSubData.Add(Enums.Submercado.Norte, new RDHHidroSubData());
            arrHidroSubData.Add(Enums.Submercado.Sul, new RDHHidroSubData());
            arrHidroSubData.Add(Enums.Submercado.SudesteCentroOeste, new RDHHidroSubData());
        }
    }
}
