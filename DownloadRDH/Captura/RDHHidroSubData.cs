using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownloadCompass
{
    /// <summary>
	/// Pacote de informações sobre a aba Hidroenergética-Subsistemas, para um Submercado específico.
	/// </summary>
	[TypeConverterAttribute(typeof(RDHHidroSubDataObjectConverter)), DescriptionAttribute("Expandir para ver mais dados...")]
    public class RDHHidroSubData : ExpandableObjectConverter
    {
        private Enums.Submercado m_objSubmercado;
        private int m_intTotalMW_MesAteData;
        private int m_intTotalMW_MediaSemanaAteData;
        private double m_dblReservatorio_Dia;
        private DateTime m_dtSemanaAtualInicio;
        private DateTime m_dtSemanaAtualFim;
        private DateTime m_dtData;

        [CategoryAttribute("Dados gerais"), DescriptionAttribute("Submercado a que se referem as informações"), Browsable(true)]
        public Enums.Submercado Submercado
        {
            get { return this.m_objSubmercado; }
            set { this.m_objSubmercado = value; }
        }

        [CategoryAttribute("Dados gerais"), DescriptionAttribute("Data de início da semana"), Browsable(true)]
        public DateTime SemanaAtualInicio
        {
            get { return this.m_dtSemanaAtualInicio; }
            set { this.m_dtSemanaAtualInicio = value; }
        }

        [CategoryAttribute("Dados gerais"), DescriptionAttribute("Data de fim da semana"), Browsable(true)]
        public DateTime SemanaAtualFim
        {
            get { return this.m_dtSemanaAtualFim; }
            set { this.m_dtSemanaAtualFim = value; }
        }

        [CategoryAttribute("Dados gerais"), DescriptionAttribute("Data das informações"), Browsable(true)]
        public DateTime Data
        {
            get { return this.m_dtData; }
            set { this.m_dtData = value; }
        }

        [CategoryAttribute("MW médio"), DescriptionAttribute("MWmed médio desde o começo do mês, até a Data"), Browsable(true)]
        public int TotalMW_MesAteData
        {
            get { return this.m_intTotalMW_MesAteData; }
            set { this.m_intTotalMW_MesAteData = value; }
        }

        [CategoryAttribute("MW médio"), DescriptionAttribute("MWmed médio desde o começo da semana, até a Data"), Browsable(true)]
        public int TotalMW_MediaSemanaAteData
        {
            get { return this.m_intTotalMW_MediaSemanaAteData; }
            set { this.m_intTotalMW_MediaSemanaAteData = value; }
        }

        [CategoryAttribute("Reservatório"), DescriptionAttribute("Reservatório % na data"), Browsable(true)]
        public double Reservatorio_Dia
        {
            get { return this.m_dblReservatorio_Dia; }
            set { this.m_dblReservatorio_Dia = value; }
        }

    }

    public class RDHHidroSubDataObjectConverter : ObjectConverterGenerico<RDHHidroSubData> { }

}
