using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownloadCompass
{
    // <summary>
    /// Pacote de informações sobre a aba Hidráulico-Hidrológica, para uma linha específica
    /// </summary>
    [TypeConverterAttribute(typeof(RDHHidraHidroDataObjectConverter)), DescriptionAttribute("Expandir para ver mais dados...")]
    public class RDHHidraHidroData : ExpandableObjectConverter
    {
        private DateTime m_dtData;
        private int m_intIdPosto;

        private int m_intVazaoNaturalDia;
        private int m_intVazaoUltMin;
        private int m_intVazaoUltMax;
        private double m_dblEnergiaArmazenada;
        private Double m_dblVolumeEspera;
        private double m_dblVazaoDefluente;
        private double m_dblVazaoIncremental;

        [CategoryAttribute("Dados gerais"), DescriptionAttribute("Data das informações"), Browsable(true)]
        public DateTime Data
        {
            get { return this.m_dtData; }
            set { this.m_dtData = value; }
        }

        [CategoryAttribute("Dados Gerais"), DescriptionAttribute("ID do posto"), Browsable(true)]
        public int Posto
        {
            get { return this.m_intIdPosto; }
            set { this.m_intIdPosto = value; }
        }

        [CategoryAttribute("Vazão Natural"), DescriptionAttribute("Vazão registrada no dia"), Browsable(true)]
        public int VazaoNaturalDia
        {
            get { return this.m_intVazaoNaturalDia; }
            set { this.m_intVazaoNaturalDia = value; }
        }

        [CategoryAttribute("Vazão Natural"), DescriptionAttribute("Vazão Ultimos dias (mais dias)"), Browsable(true)]
        public int VazaoNaturalUltMax
        {
            get { return this.m_intVazaoUltMax; }
            set { this.m_intVazaoUltMax = value; }
        }

        [CategoryAttribute("Vazão Natural"), DescriptionAttribute("Vazão Ultimos dias (menos dias)"), Browsable(true)]
        public int VazaoNaturalUltMin
        {
            get { return this.m_intVazaoUltMin; }
            set { this.m_intVazaoUltMin = value; }
        }

        [CategoryAttribute("Valores do Dia"), DescriptionAttribute("Vazão registrada no dia"), Browsable(true)]
        public double EnergiaArmazenada
        {
            get { return this.m_dblEnergiaArmazenada; }
            set { this.m_dblEnergiaArmazenada = value; }
        }

        [CategoryAttribute("Valores do Dia"), DescriptionAttribute("Volume de espera"), Browsable(true)]
        public Double VolumeEspera
        {
            get { return this.m_dblVolumeEspera; }
            set { this.m_dblVolumeEspera = value; }
        }

        [CategoryAttribute("Valores do Dia"), DescriptionAttribute("Vazão defluente"), Browsable(true)]
        public double VazaoDefluente
        {
            get { return this.m_dblVazaoDefluente; }
            set { this.m_dblVazaoDefluente = value; }
        }

        [CategoryAttribute("Valores do Dia"), DescriptionAttribute("Vazão defluente"), Browsable(true)]
        public double VazaoIncremental
        {
            get { return this.m_dblVazaoIncremental; }
            set { this.m_dblVazaoIncremental = value; }
        }

    }

    public class RDHHidraHidroDataObjectConverter : ObjectConverterGenerico<RDHHidraHidroData> { }

}
