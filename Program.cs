
using GeneraPdfArticolo26.Servizi;
using System;

namespace GeneraPdfArticolo26
{
    class Program
    { 

        static void Main(string[] args)
        {    

            if (args.Length == 3 || args.Length == 4) {

                try {
                    bool pdfa = false;
                    if (args.Length == 4 && args[3].ToLower() == "pdfa") pdfa = true;

                    ExcelToPdf xlsx2pdf = new ExcelToPdf(args[0], args[1], args[2],pdfa);

                    xlsx2pdf.Importa();
                    xlsx2pdf.Esporta();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Environment.ExitCode = -1;
                }

        
            }
            else
            { 
                Console.WriteLine("Uso: GeneraPdfArticolo26 FileExcelSorgente(percorso completo) CartellaDestinazione FileConfigurazione pdfa (parametro facoltativo)");
            }
                        
        }

    }
}
