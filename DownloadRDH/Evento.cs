using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DownloadCompass
{
    public class Evento
    {
        private string _key;

        public string Texto { get; set; }
        public string Key { get => _key; set => _key = Regex.Replace(value.Replace('\r', ' ').Replace('\n', ' '), "\\s{2,}", " "); }
        public string Href { get; set; }

        public override int GetHashCode()
        {
            return Texto.GetHashCode() + Href.GetHashCode();
        }

    }
}
