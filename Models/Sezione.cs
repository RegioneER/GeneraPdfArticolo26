using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneraPdfArticolo26.Models
{
    public enum TipoSezione { semplice, multipla, condizionale, dichiarazioni, singolovalore }
  
    class Sezione
    {
        public string NomeSezione { get; set; } = "";

        public string NomeFoglio { get; set; } = "";

        public string OrdinaPer { get; set; } = "";

        public TipoSezione Tipo { get; set; } = TipoSezione.semplice;

        public bool SezioneIndice { get; set; } = false;

        public int RigaIntestazione { get; set; } = 1;

        public int RigaDati { get; set; } =  2;

        public string[] CampoMultiplo { get; set; } = Array.Empty<String>();

        public string TestoSeVuota { get; set; } = "";

        public string[] CampiFacoltativi { get; set; } = Array.Empty<String>();

        public string[] EscludiCampi { get; set; } = Array.Empty<String>();

        public string[] LinkElixTraParentesi { get; set; } = Array.Empty<String>();

		public string[] Link { get; set; } = Array.Empty<String>();

		public float[] Tabella { get; set; } = Array.Empty<float>();

        public IList<DatoSezione> Dati = new List<DatoSezione>();

    }


}
