using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GeneraPdfArticolo26.Models;
using System.IO;
using OfficeOpenXml;
using iText.Kernel.Pdf;
using iText.Kernel.Font;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Pdfa;
using iText.IO.Font;
using iText.Kernel.Events;
using iText.Kernel.Pdf.Navigation;
using iText.Kernel.Geom;
using iText.IO.Font.Constants;
using iText.IO.Image;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Net;
using System.Diagnostics;




namespace GeneraPdfArticolo26.Servizi
{
    class ExcelToPdf
    {
        private Bando _bando;

        private string _cartellaDestinazione;

        private string _fileOrigine;

        private string _fileConfigurazione;

        private bool _pdfa;

        
        public ExcelToPdf(string fileOrigine, string cartellaDestinazione, string fileConfigurazione,bool pdfa)
        {

            _cartellaDestinazione = cartellaDestinazione;
            _fileOrigine = fileOrigine;
            _fileConfigurazione = fileConfigurazione;
            _pdfa = pdfa;

            Configura();


            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Converters =
                    {  
                        new JsonStringEnumConverter(JsonNamingPolicy.CamelCase)
                    }
            };


            //check
            var prova = JsonSerializer.Serialize<Bando>(_bando,options);


        }

        public string[] EstraiValoriMultipli(string valoriComposti)
        {
            string[] valori;

            string multivalore;

            multivalore = valoriComposti.Replace(", ", "__");
            valori = multivalore.Split(",");
            for (int i = 0; i < valori.Length; i++) valori[i] = valori[i].Replace("__", ", ");

            return valori;
        }

        private void Configura()
        {

            if (_fileConfigurazione == null || _fileConfigurazione.Length==0) {

                throw new ArgumentNullException("File di configurazione");
            };

            FileInfo fileConfig = new FileInfo(_fileConfigurazione);

            if(fileConfig.Exists) {

                //imposto le opzioni per deserializzare gli enum
                var options = new JsonSerializerOptions
                {
                    Converters =
                    {
                        new JsonStringEnumConverter(JsonNamingPolicy.CamelCase)
                    }
                };

                //leggo il config e lo deserializzo
                string testoConfigurazione = File.ReadAllText(_fileConfigurazione,Encoding.UTF8);
                _bando = JsonSerializer.Deserialize<Bando>(testoConfigurazione, options);

                if (_bando.NomeColonnaAmmessi !="" )
                {
                    _bando.ValorePositivo = _bando.ValorePositivo.ToLower().Trim();
                    _bando.ValoreNegativo = _bando.ValoreNegativo.ToLower().Trim();

                    if (_bando.ValorePositivo == "" && _bando.ValoreNegativo == "") throw new Exception("i valori di selezione non possono essere entrambi nulli");
                    

                     if ( _bando.ValorePositivo == _bando.ValoreNegativo) throw new Exception("i valori di selezione non possono essere uguali");
                }
                    
                var cercaSezioneIndice = from Sezione sezione in _bando.Sezioni where sezione.NomeFoglio == _bando.NomeFoglioIndice select sezione;

                var sezioneindice = cercaSezioneIndice.FirstOrDefault<Sezione>();

                if(sezioneindice is not null )
                {
                    if (sezioneindice.Tipo == TipoSezione.semplice)
                        sezioneindice.SezioneIndice = true;
                    else
                        throw new Exception("La sezione indice può essere solo di tipo semplice");
                }
                else
                {
                    throw new Exception("Il foglio della sezione indice specificato non esiste");
                }


            } else
            {
                throw new FileNotFoundException("File di configurazione non esiste", _fileConfigurazione);
            };

        }

        public void Importa()
        {
            if (_fileOrigine == null || _fileOrigine.Length == 0)
            {
                throw new ArgumentNullException("File di configurazione");
            };


            FileInfo fileExcel = new FileInfo(_fileOrigine);
            if (fileExcel.Exists)
            {
                //Imposto la licenza a non commericale in qualità di progetto open
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                //Leggo il file excel
                using (var package = new ExcelPackage(fileExcel))
                {
                    var wb = package.Workbook;

                    //importo le sezioni
                    foreach (Sezione sezione in _bando.Sezioni)
                    {
                        ExcelWorksheet foglioSezione;
                        foglioSezione = wb.Worksheets[sezione.NomeFoglio];
                        int colonnaCF = 0;
                        int colonnaGruppo = 0;
                        int colonnaOrdinamentoPerGruppo = 0;
                        int colonnaSottotitolo = 0;
                        int colonnaSottotitoloGruppo = 0;
                        int colonnaAmmessi = 0;
                        int colonnaCampoFooter = 0;
                        int colonnaID = 0;
                        int colonnaPercorso = 0;


                        if (foglioSezione is not null)
                        {
 

                            for(int riga=sezione.RigaDati; riga<= foglioSezione.Dimension.End.Row; riga++)
                            {
                                DatoSezione dato = new DatoSezione();
                                DatoSezione datoIndice= new DatoSezione();
                                bool ammesso;
                                string ammissibilita;
                                string sottotitoloGruppoTesto="";
                                string nomeGruppoTesto="";


                                // se è la sezione indice cerco le colonne speciali
                                if (sezione.SezioneIndice) { 
                                    for (int col = 2; col <= foglioSezione.Dimension.End.Column; col++) { 
                                        if (_bando.NomeCampoCF.ToLower().Trim() != "" && foglioSezione.Cells[1, col].Text.ToLower().Trim() == _bando.NomeCampoCF.ToLower().Trim()) colonnaCF = col;
                                        if (_bando.NomeColonnaAmmessi.ToLower().Trim() != "" && foglioSezione.Cells[1, col].Text.ToLower().Trim() == _bando.NomeColonnaAmmessi.ToLower().Trim()) colonnaAmmessi = col;
                                        if (_bando.NomeCampoGruppo.ToLower().Trim() != "" && foglioSezione.Cells[1, col].Text.ToLower().Trim() == _bando.NomeCampoGruppo.ToLower().Trim()) colonnaGruppo = col;
                                        if (_bando.NomeCampoOrdinamentoGruppo.ToLower().Trim()!= "" && foglioSezione.Cells[1, col].Text.ToLower().Trim() == _bando.NomeCampoOrdinamentoGruppo.ToLower().Trim()) colonnaOrdinamentoPerGruppo = col;
                                        if (_bando.NomeCampoFooter.ToLower().Trim() != "" && foglioSezione.Cells[1, col].Text.ToLower().Trim() == _bando.NomeCampoFooter.ToLower().Trim()) colonnaCampoFooter = col;
                                        if (_bando.NomeCampoID.ToLower().Trim() != "" && foglioSezione.Cells[1, col].Text.ToLower().Trim() == _bando.NomeCampoID.ToLower().Trim()) colonnaID = col;
                                        if (_bando.NomeCampoSottotitoloGruppo.ToLower().Trim() !="" && foglioSezione.Cells[1, col].Text.ToLower().Trim() == _bando.NomeCampoSottotitoloGruppo.ToLower().Trim()) colonnaSottotitoloGruppo = col;
                                        if (_bando.NomeCampoSottotitolo.ToLower().Trim() !="" && foglioSezione.Cells[1, col].Text.ToLower().Trim() == _bando.NomeCampoSottotitolo.ToLower().Trim()) colonnaSottotitolo = col;
                                        if (_bando.NomeCampoPercorso.ToLower().Trim()!="" && foglioSezione.Cells[1, col].Text.ToLower().Trim() == _bando.NomeCampoPercorso.ToLower().Trim()) colonnaPercorso = col;
                                    }
                                    if (colonnaCF == 0 && colonnaID==0) throw new Exception("Non esiste la colonna specificata ID o codice fiscale/partita iva nella sezione indicata");
                                    if (_bando.NomeColonnaAmmessi != "" && colonnaAmmessi==0) throw new Exception("Non esiste la colonna specificata per selezionare gli ammessi nella sezione indicata");
                                    if (_bando.NomeCampoOrdinamentoGruppo != "" && colonnaOrdinamentoPerGruppo == 0) throw new Exception("Non esiste la colonna specificata per l'ordinamento nei gruppi");
                                    if (_bando.NomeCampoGruppo != "" && colonnaGruppo == 0) throw new Exception("Non esiste la colonna specificata per il raggruppamento");
                                    if (_bando.NomeCampoFooter != "" && colonnaCampoFooter== 0) throw new Exception("Non esiste la colonna specificata per il campo footer");
                                    if (_bando.NomeCampoPercorso != "" && colonnaPercorso == 0) throw new Exception("Non esiste la colonna specificata per il percorso");
                                };


                                if (colonnaAmmessi != 0)
                                {

                                    ammesso = false;
                                    ammissibilita = foglioSezione.Cells[riga, colonnaAmmessi].Text.ToLower().Trim();
                                    if (_bando.ValoreNegativo != "" && ammissibilita != _bando.ValoreNegativo)
                                        ammesso = true;
                                    else if (_bando.ValorePositivo != "" && ammissibilita == _bando.ValorePositivo)
                                        ammesso = true;
                                }
                                else 
                                    ammesso=true;


                                for (int colonna=1; colonna <= foglioSezione.Dimension.End.Column; colonna++)
                                {
                                    string valore;
                                    string intestazione;
 
                                   
                                    intestazione = WebUtility.HtmlDecode(foglioSezione.Cells[sezione.RigaIntestazione, colonna].Text).Replace("*","").Replace("<br>"," ").Trim();
                                    if(intestazione=="")
									{
                                        intestazione = "_" + Guid.NewGuid();

                                    }if(dato.Valori.ContainsKey(intestazione))
                                    {
                                        intestazione = intestazione + "|" + Guid.NewGuid();  //se esiste già una colonna con lo stesso nome aggiungo |
                                    }
                                   
                                    valore = WebUtility.HtmlDecode(foglioSezione.Cells[riga, colonna].Text).Trim();
                                    if (colonna==1)
                                    {
                                        if (valore != "")
                                        {
                                            dato.Id = valore;
                                            if (sezione.SezioneIndice) {
                                                datoIndice.Id = valore;
                                             }

                                        }
                                        else
                                            throw new Exception($"sezione {sezione.NomeFoglio} contiene un id nullo");
                                    }
                                    else
                                    {
                                        if(valore == "" && sezione.Tabella.Length != 0 )
                                        {
                                           
                                            if (!sezione.CampiFacoltativi.Contains(intestazione))
                                                dato.Valori.Add(intestazione, valore);

                                        } else 
                                        {
                                            //se non è nullo e non è nei facoltativi inserisco e non è da escludere
                                            if (valore !=""  && !sezione.EscludiCampi.Contains(intestazione))
                                            { 
                                                dato.Valori.Add(intestazione, valore); 
                                            }
                                            else
											{
                                                Debug.WriteLine($"escluso {intestazione}");
											}

                                            if (sezione.SezioneIndice && colonnaID ==0 && colonna==colonnaCF && (valore.Length == 11 || valore.Length == 16))
                                            { 
                                                if(valore.Length == 11 || valore.Length == 16)
                                                    datoIndice.Valori.Add("NomeFile", valore );
                                                else throw new Exception("Partita iva/codice fiscale in formato non valido");
                                            }

                                            if (sezione.SezioneIndice && colonna == colonnaID && colonnaCF==0)
                                            {
                                                datoIndice.Valori.Add("NomeFile", valore);
                                            }

                                            if (sezione.SezioneIndice && colonna == colonnaGruppo)
                                            {
                                                datoIndice.Valori.Add("NomeGruppo", valore);
                                                nomeGruppoTesto = valore;
                                            }
                                            if (sezione.SezioneIndice && colonna== colonnaOrdinamentoPerGruppo)
                                            {
                                                datoIndice.Valori.Add("CampoOrdinamento", valore);
                                            }
                                            if (sezione.SezioneIndice && colonna == colonnaCampoFooter)
                                            {
                                                datoIndice.Valori.Add("Footer", valore);
                                            }
                                            if (sezione.SezioneIndice && colonna == colonnaSottotitolo)
                                            {
                                                datoIndice.Valori.Add("Sottotitolo", valore);
                                            }
                                            if (sezione.SezioneIndice && colonna == colonnaSottotitoloGruppo)
                                            {
                                                sottotitoloGruppoTesto = valore;
                                            }
                                            if (sezione.SezioneIndice && colonna == colonnaPercorso)
                                            {
                                                datoIndice.Valori.Add("Percorso", valore);
                                            }

                                        }
                                    };
                                }

                                if(sezione.SezioneIndice)
                                {
                                    if(colonnaSottotitoloGruppo != 0 && colonnaGruppo !=0)
                                    {
                                        datoIndice.Valori.Add("sottotitolo"+ nomeGruppoTesto, sottotitoloGruppoTesto);
                                    }
                                };
                                sezione.Dati.Add(dato);
                                if (sezione.SezioneIndice && ammesso) _bando.Indice.Add(datoIndice); 
                            }; 
                        }
                        else
                        {
                            throw new Exception($"la sezione '{sezione.NomeFoglio}' è inesistente");
                        }

                    }

                }
            }
            else
            {
                throw new FileNotFoundException("File excel non esiste", _fileConfigurazione);
            };

            Console.WriteLine($"Letti '{_bando.Indice.Count}' nomi file");

        }


        public void Esporta()
		{
            EsportaSingoli();
		}
        public void EsportaSingoli()
        {
            int FilePrototti = 0;
            long MaxDocumenti = _bando.MaxDocumenti;
            var percorsoFonts = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            var PercorsoSistema = Environment.GetFolderPath(Environment.SpecialFolder.System);
            string percorsoLogo = "";
            string percorsoEseguibile = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            string etichettaCampo="";
            string valoreCampo = "";
            string collegamentoCampo="";

            if(_bando.Logo == "")
            {
                percorsoLogo = "";

            } else if(File.Exists(_cartellaDestinazione + $"\\{_bando.Logo}"))
            {
                percorsoLogo = _cartellaDestinazione + $"\\{_bando.Logo}";

            } else if(File.Exists(percorsoEseguibile + $"\\{_bando.Logo}"))
            {
                percorsoLogo = percorsoEseguibile + $"\\{_bando.Logo}";
            } else
            {
                throw new ArgumentNullException($"Manca il file {_bando.Logo}");
            }


            if (_cartellaDestinazione == null || _cartellaDestinazione.Length == 0)
            {
                throw new ArgumentNullException("Cartella di destinazione");
            };

            if (Directory.Exists(_cartellaDestinazione))
            { 

                foreach(DatoSezione voceIndice in _bando.Indice)
                {

                    if (FilePrototti < MaxDocumenti)
                    {
                        string percorso = "";

                    
                        if (voceIndice.Valori.ContainsKey("Percorso")) percorso = voceIndice.Valori["Percorso"];
                        string nomefiledestinazione = System.IO.Path.Combine(_cartellaDestinazione, percorso, voceIndice.Valori["NomeFile"] + ".pdf");


                        if (!Directory.Exists(System.IO.Path.Combine(_cartellaDestinazione, percorso))) Directory.CreateDirectory(System.IO.Path.Combine(_cartellaDestinazione, percorso));

                        using (var pdfWriter = new PdfWriter(nomefiledestinazione))
                        {
                            PdfDocument pdf;
                            Document document;
                            PdfFontFactory.EmbeddingStrategy strategiaFont;
                            TextFooterEventHandler textFooterHandler;

                            if (_pdfa)
                            {
                                pdf = new PdfADocument(pdfWriter, PdfAConformanceLevel.PDF_A_3A, new PdfOutputIntent("Custom", "", "https://www.color.org", "sRGB IEC61966-2.1", new FileStream(PercorsoSistema + "\\spool\\drivers\\color\\sRGB Color Space Profile.icm", FileMode.Open, FileAccess.Read)));
                                document = new Document(pdf, iText.Kernel.Geom.PageSize.A4);
                                document.SetMargins( 30 ,  20,  30, 20);
           
                                //Imposto alcuni parametri necessari per i PDF/A
                                pdf.SetTagged();
                                pdf.GetCatalog().SetLang(new PdfString("it-IT"));
                                pdf.GetCatalog().SetViewerPreferences(new PdfViewerPreferences().SetDisplayDocTitle(true));
                                PdfDocumentInfo info = pdf.GetDocumentInfo();
                                info.SetTitle(_bando.Titolo);

                                strategiaFont = PdfFontFactory.EmbeddingStrategy.FORCE_EMBEDDED;
                                textFooterHandler = new TextFooterEventHandler(document, strategiaFont, _bando.Numerazione);
                                pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, textFooterHandler);
                                pdf.AddEventHandler(PdfDocumentEvent.START_PAGE, new TextHeaderEventHandler(document, percorsoLogo, _bando.LogoAlt, strategiaFont));
                            }
                            else
                            {
                                pdf = new PdfDocument(pdfWriter);
                                document = new Document(pdf, iText.Kernel.Geom.PageSize.A4);

                                strategiaFont = PdfFontFactory.EmbeddingStrategy.FORCE_NOT_EMBEDDED;
                                textFooterHandler = new TextFooterEventHandler(document, strategiaFont, _bando.Numerazione);
                                pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, textFooterHandler);
                                pdf.AddEventHandler(PdfDocumentEvent.START_PAGE, new TextHeaderEventHandler(document, percorsoLogo, _bando.LogoAlt, strategiaFont));

                            }

                            Paragraph paragrafo;
                            Table tabella;
                               


                            var fontTesto = PdfFontFactory.CreateFont(percorsoFonts + "\\Tahoma.ttf", strategiaFont);
                              
                            var fontSezione = PdfFontFactory.CreateFont(percorsoFonts + "\\Tahomabd.ttf", strategiaFont); 

                        
                            //Metto il titolo
                            Paragraph header = new Paragraph(new Text(_bando.Titolo).SetFont(fontTesto)
                                   .SetFontSize(_bando.DimensioneFontTitolo))
                                   .SetTextAlignment(TextAlignment.CENTER);

                            document.Add(header);

                            if (voceIndice.Valori.ContainsKey("Sottotitolo")) 
                             {
                                 Paragraph subHeader = new Paragraph(new Text(voceIndice.Valori["Sottotitolo"]).SetFont(fontTesto)
                                 .SetFontSize(_bando.DimensioneFontTitolo))
                                 .SetTextAlignment(TextAlignment.CENTER);
                                 document.Add(subHeader);
                             }
                        


                            if (voceIndice.Valori.ContainsKey("Footer"))
                                textFooterHandler.setFooterFor(voceIndice.Valori["NomeFile"] + "-" + voceIndice.Valori["Footer"], pdf.GetLastPage());

                            //Aggiungo le sezioni
                            foreach (Sezione sezione in _bando.Sezioni)
                            {
                                string TitoloSezione = sezione.NomeSezione == "" ? sezione.NomeFoglio : sezione.NomeSezione;

                                paragrafo = new Paragraph(new Text(TitoloSezione).SetFont(fontSezione)
                                .SetFontSize(_bando.DimensioneFontSezione))
                                .SetTextAlignment(TextAlignment.LEFT);

                                document.Add(paragrafo);

                                var listaDatiSezione = from DatoSezione dati in sezione.Dati where dati.Id == voceIndice.Id select dati;

                                int numeroElementi = listaDatiSezione.Count();
                                if (sezione.Tipo == TipoSezione.multipla && numeroElementi > 1 && sezione.OrdinaPer != "")
                                {
                                    listaDatiSezione = listaDatiSezione.OrderBy(datosezione => datosezione.Valori[sezione.OrdinaPer]);
                                };

                                //se la sezione è vuota metto il testo TestoSeVuota
                                if (numeroElementi == 0)
                                {
                                    var testoSezioneVuota = new Paragraph(sezione.TestoSeVuota).SetFont(fontTesto).SetFontSize(_bando.DimensioneFontTesto);
                                    document.Add(testoSezioneVuota);
                                };


                                int contatoreElementi = 0;


                                if (sezione.Tabella.Count() == 0)
                                {

                                    if (sezione.Tipo == TipoSezione.singolovalore && numeroElementi==1)
                                    {
                                        var singoloValore = sezione.Dati.FirstOrDefault().Valori.FirstOrDefault().Value;

                                        var testoSezioneSingoloValore = new Paragraph(singoloValore).SetFont(fontTesto).SetFontSize(_bando.DimensioneFontTesto);
                                        document.Add(testoSezioneSingoloValore);
                                    }
                                    else
                                    { 
                                    //è una sezione tabella eventualmente in lista
                                        foreach (DatoSezione datiSezione in listaDatiSezione)
                                        {

                                            if (numeroElementi > 1)
                                            {

                                                contatoreElementi++;
                                                paragrafo = new Paragraph(new Text(" " + contatoreElementi.ToString() + " ").SetFont(fontSezione)
                                                .SetFontSize(8).SetFontColor(iText.Kernel.Colors.ColorConstants.WHITE))
                                                .SetTextAlignment(TextAlignment.CENTER);
                                                paragrafo.SetWidth(12).SetBackgroundColor(new iText.Kernel.Colors.DeviceRgb(119, 119, 119)).SetBorderRadius(new BorderRadius(5f));
                                                document.Add(paragrafo);

                                            }


                                            //metto i dati in lista di tabelle
                                            tabella = new Table(UnitValue.CreatePercentArray(new float[] { 30f, 70f })).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout().SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                                            foreach (var dato in datiSezione.Valori)
                                            {
                                                //rendo nuovamente anonime le celle
                                                if (dato.Key.StartsWith("_"))
                                                {
                                                    etichettaCampo = "";
                                                }
												else if (dato.Key.Contains("|")) //gestisco le colonne omonime
												{
													int pos = dato.Key.IndexOf("|");

													etichettaCampo = dato.Key.Substring(0, pos);
												} else
												{
                                                    etichettaCampo = dato.Key;
												};



												if (dato.Value!="")
                                                {

                                                    var nome = new Paragraph(etichettaCampo).SetFont(fontTesto).SetFontSize(_bando.DimensioneFontTesto).SetTextAlignment(TextAlignment.RIGHT).SetFontColor(new iText.Kernel.Colors.DeviceRgb(66, 66, 66));

                                                    var bordo = new iText.Layout.Borders.SolidBorder(new iText.Kernel.Colors.DeviceRgb(186, 186, 186), 1f);

                                                    var cellaIntestazione = new Cell().Add(nome);
                                                    cellaIntestazione.SetBackgroundColor(new iText.Kernel.Colors.DeviceRgb(233, 233, 233)).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetBorderBottom(bordo).SetBorderBottom(bordo).SetBorderTop(bordo).SetBorderLeft(bordo).SetBorderRight(bordo); ;

                                                    tabella.AddCell(cellaIntestazione);

                                                    var cellaValore = new Cell();

                                                    if (sezione.CampoMultiplo.Contains(etichettaCampo))
                                                    {
                                                        List unorderedelist = new List().SetFont(fontTesto).SetFontSize(_bando.DimensioneFontTesto);
                                                        unorderedelist.SetListSymbol("\u2022 ").SetListSymbolAlignment(ListSymbolAlignment.LEFT);
                                                        string[] valori = EstraiValoriMultipli(dato.Value);
                                                        foreach (string listItemValue in valori)
                                                        {
                                                            unorderedelist.Add(new ListItem(listItemValue));
                                                        };

                                                        cellaValore = cellaValore.Add(unorderedelist);
                                                    }
                                                    else
                                                    {
                                                        if(sezione.LinkElixTraParentesi.Contains(etichettaCampo))
                                                        {
                                                            valoreCampo = dato.Value.Split("(").ToList().FirstOrDefault().Trim();
															collegamentoCampo = dato.Value.Split("(").ToList().LastOrDefault().Replace(")","").Trim();
														}
														else if(sezione.Link.Contains(etichettaCampo))
                                                        {
                                                            valoreCampo = dato.Value;
                                                            collegamentoCampo = valoreCampo;

														}
                                                        else
                                                        {
                                                            valoreCampo = dato.Value;
															collegamentoCampo = "";
                                                        };

														var valore = new Paragraph(valoreCampo).SetFont(fontTesto).SetFontSize(_bando.DimensioneFontTesto);
                                                        cellaValore = cellaValore.Add(valore);

														if (collegamentoCampo != "") cellaValore.SetAction(iText.Kernel.Pdf.Action.PdfAction.CreateURI(collegamentoCampo));

													}

                                                    cellaValore.SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetBorderBottom(bordo).SetBorderTop(bordo).SetBorderLeft(bordo).SetBorderRight(bordo);
                                                    tabella.AddCell(cellaValore);

                                                }

                                            };

                                        document.Add(tabella);
                                        }

                                    };

                                }
                                else
                                {

                                    //è una sezione con tabella unica
                                    tabella = new Table(UnitValue.CreatePercentArray(sezione.Tabella)).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout().SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                                    bool creaIntestazione = true;
                                    //creo l'header
                                    foreach (DatoSezione datoSezione in listaDatiSezione)
                                    {
                                        //al primo giro creo la riga intestazione
                                        if (creaIntestazione)
                                            foreach (var dato in datoSezione.Valori)
                                            {
												//rendo nuovamente anonime le celle
												if (dato.Key.StartsWith("_"))
												{
													etichettaCampo = "";
												}
												else if (dato.Key.Contains("|"))
												{
													int pos = dato.Key.IndexOf("|");


													etichettaCampo = dato.Key.Substring(0, pos);
												}
												else
												{
													etichettaCampo = dato.Key;
												};
												var nome = new Paragraph(etichettaCampo).SetFont(fontTesto).SetFontSize(_bando.DimensioneFontTesto).SetTextAlignment(TextAlignment.LEFT).SetFontColor(new iText.Kernel.Colors.DeviceRgb(66, 66, 66));
                                                var bordo = new iText.Layout.Borders.SolidBorder(new iText.Kernel.Colors.DeviceRgb(186, 186, 186), 1f);
                                                var cellaIntestazione = new Cell().Add(nome);
                                                cellaIntestazione.SetBackgroundColor(new iText.Kernel.Colors.DeviceRgb(233, 233, 233)).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetBorderBottom(bordo).SetBorderTop(bordo).SetBorderLeft(bordo).SetBorderRight(bordo);
                                                tabella.AddHeaderCell(cellaIntestazione);
                                                creaIntestazione = false;
                                            };
                                        //aggiungo la riga dati
                                        foreach (var dato in datoSezione.Valori)
                                        {
											//rendo nuovamente anonime le celle
											if (dato.Key.StartsWith("_"))
											{
												etichettaCampo = "";
											}
											else if (dato.Key.Contains("|"))
											{
												int pos = dato.Key.IndexOf("|");


												etichettaCampo = dato.Key.Substring(0, pos);
											}
											else
											{
												etichettaCampo = dato.Key;
											};
											var cellaValore = new Cell();
                                            var bordo = new iText.Layout.Borders.SolidBorder(new iText.Kernel.Colors.DeviceRgb(186, 186, 186), 1f);
                                            if (sezione.CampoMultiplo.Contains(etichettaCampo) && dato.Value != "")
                                            {
                                                List unorderedelist = new List().SetFont(fontTesto).SetFontSize(_bando.DimensioneFontTesto);
                                                unorderedelist.SetListSymbol("\u2022 ").SetListSymbolAlignment(ListSymbolAlignment.LEFT);
                                                string[] valori = EstraiValoriMultipli(dato.Value);
                                                foreach (string listItemValue in valori)
                                                {
                                                    unorderedelist.Add(new ListItem(listItemValue));
                                                };

                                                cellaValore = cellaValore.Add(unorderedelist);
                                            }
                                            else
                                            {
												if (sezione.LinkElixTraParentesi.Contains(etichettaCampo))
												{
													valoreCampo = dato.Value.Split("(").ToList().FirstOrDefault().Trim();
													collegamentoCampo = dato.Value.Split("(").ToList().LastOrDefault().Replace(")", "").Trim();
												}
												else if (sezione.Link.Contains(etichettaCampo))
												{
													valoreCampo = dato.Value;
													collegamentoCampo = valoreCampo;

												}
												else
												{
													valoreCampo = dato.Value;
													collegamentoCampo = "";
												};


												var valore = new Paragraph(valoreCampo).SetFont(fontTesto).SetFontSize(_bando.DimensioneFontTesto);
												cellaValore = cellaValore.Add(valore);

												if (collegamentoCampo != "") cellaValore.SetAction(iText.Kernel.Pdf.Action.PdfAction.CreateURI(collegamentoCampo));


											}

                                            cellaValore.SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetBorderBottom(bordo).SetBorderTop(bordo).SetBorderLeft(bordo).SetBorderRight(bordo);
                                            tabella.AddCell(cellaValore);
                                        };

                                    }



                                    document.Add(tabella);
                                }



                            };

                            document.Close();
                            Console.WriteLine($"Creato '{nomefiledestinazione}' pdf");
                            FilePrototti++;
                        }
                    }
                }
            }
            else
            {
                throw new FileNotFoundException("Cartella di destionazione inesistente", _fileConfigurazione);
            };
            Console.WriteLine($"Prodotti '{FilePrototti}' pdf");
        }


        private class TextFooterEventHandler : IEventHandler
        {
            protected Document doc;
            protected int pageNumber = 0;
            protected PdfFontFactory.EmbeddingStrategy strategiaFont;
            protected bool numerazione;
            protected string testoFooter="";
            protected Dictionary<PdfPage, String> footers = new Dictionary<PdfPage, string>();
  

            public void setFooterFor(String footer, PdfPage page)
            {
                footers[page] = footer;
            }




            public TextFooterEventHandler(Document doc, PdfFontFactory.EmbeddingStrategy strategiaFont, bool numerazione)
            {
                this.doc = doc;
                this.strategiaFont = strategiaFont;
                this.numerazione = numerazione;
            }


            public void HandleEvent(Event currentEvent)
            {
                PdfDocumentEvent docEvent = (PdfDocumentEvent)currentEvent;
                PdfPage page = docEvent.GetPage();
                Rectangle pageSize = docEvent.GetPage().GetPageSize();

                var font = PdfFontFactory.CreateFont(Environment.GetFolderPath(Environment.SpecialFolder.Fonts) + "\\Tahoma.ttf", strategiaFont);

                float coordX = ((pageSize.GetLeft() + doc.GetLeftMargin())
                                 + (pageSize.GetRight() - doc.GetRightMargin())) / 2;
                float headerY = pageSize.GetTop() - doc.GetTopMargin() + 10;
                float footerY = doc.GetBottomMargin() - 10f;
                float coordXNumber = pageSize.GetRight() - doc.GetRightMargin() - 30.0f;
                pageNumber++;
                string numeroPagina = numerazione ? pageNumber.ToString() : "";
                Canvas canvas = new Canvas(docEvent.GetPage(), pageSize);

                if (footers.ContainsKey(page))
                {
                    testoFooter = footers[page];
                    footers.Remove(page);
                }
                 
   
                    canvas.SetFont(font)
                    .SetFontSize(6)
                    .ShowTextAligned(testoFooter, coordX, footerY, TextAlignment.CENTER)
                    .ShowTextAligned(numeroPagina, coordXNumber, footerY, TextAlignment.RIGHT)
                    .Close();
                //                    .ShowTextAligned("this is a header", coordX, headerY, TextAlignment.CENTER)
            }
        }

        private class TextHeaderEventHandler : IEventHandler
        {
            protected Document doc;
            protected string percorsoLogo;
            protected string descrizioneLogo;
            protected PdfFontFactory.EmbeddingStrategy strategiaFont;
            

            public TextHeaderEventHandler(Document doc, string percorsoLogo,string descrizioneLogo, PdfFontFactory.EmbeddingStrategy strategiaFont)
            {
                this.doc = doc;
                this.percorsoLogo = percorsoLogo;
                this.strategiaFont = strategiaFont;
                this.descrizioneLogo = descrizioneLogo;
            }


            public void HandleEvent(Event currentEvent)
            {
                PdfDocumentEvent docEvent = (PdfDocumentEvent)currentEvent;
                Rectangle pageSize = docEvent.GetPage().GetPageSize();
                Canvas canvas;

                var font = PdfFontFactory.CreateFont(Environment.GetFolderPath(Environment.SpecialFolder.Fonts) + "\\Tahoma.ttf", strategiaFont);

                float coordX = pageSize.GetLeft() + doc.GetLeftMargin();
                float headerY = pageSize.GetTop() - doc.GetTopMargin()+10; 
                float dimension = pageSize.GetWidth();

                if (this.percorsoLogo != "")
                {
                    ImageData imageData = ImageDataFactory.Create(percorsoLogo);

                    Image image = new Image(imageData).ScaleToFit(100f, 100f).SetFixedPosition(coordX, headerY);
                    image.GetAccessibilityProperties().SetAlternateDescription(this.descrizioneLogo);

                    canvas = new Canvas(docEvent.GetPage(), pageSize);
                    canvas.Add(image);
                }
                else canvas = new Canvas(docEvent.GetPage(), pageSize);

            }
        }

    }
}
