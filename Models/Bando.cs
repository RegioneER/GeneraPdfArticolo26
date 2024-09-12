using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace GeneraPdfArticolo26.Models
{
    class Bando
    {
        public string NomeBando { get; set; } = "";

        public string BandoTAG { get; set; } = "";

        public string Titolo { get; set; } = "";

        public string NomeCampoSottotitolo { get; set; } = "";

        public string TitoloGruppo { get; set; } = "";

        public string NomeFoglioIndice { get; set; } = "";

        public string NomeCampoCF { get; set; } = "";

        public string NomeCampoID { get; set; } = "";

        public string NomeCampoFooter { get; set; } = "";

        public string NomeCampoPercorso { get; set; } = "";
 
        public bool EsportaPerGruppo { get; set; } = false;

        public string NomeCampoGruppo { get; set; } = "";

        public string NomeCampoOrdinamentoGruppo { get; set; } = "";

        public string NomeCampoSottotitoloGruppo { get; set; } = "";

        public int DimensioneFontTitolo { get; set; } = 16;

        public int DimensioneFontSezione { get; set; } = 12;

        public int DimensioneFontTesto { get; set; } = 10;

        public string NomeColonnaAmmessi { get; set; } = "";

        public string ValorePositivo { get; set; } = "";

        public string ValoreNegativo { get; set; } = "";

        public long MaxDocumenti { get; set; } = 32000;

        public bool Numerazione { get; set; } = false;

        public string Logo { get; set; } = "";

        public string LogoAlt { get; set; } = "";

        public IList<DatoSezione> Indice { get; set; } = new List<DatoSezione>();

        public IList<Sezione> Sezioni { get; set; } = new List<Sezione>();

    }

}
