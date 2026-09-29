using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TehnickiPregled
{
    public class vozilo
    {
        public string IDBroj; public string RegistarskiBroj; public string VrstaVozila;public string PodvrstaVozila; public string klasa;public string oblikkaroserije;
        public string specijalno;public string Marka; public string Model;public string Tip;public string GodinaProizvodnje; public string BrojSasije;
        public string BrojMotora; public string Boja; public string DatumPrveRegostracije; public string Odnos; public string Snaga;public string Zapremina;
        public string NajvecaMasa; public string Tezina; public string Nosivost; public string Sedenje; public string Stajanje; public string BrojOsovina;
        public string Gorivo; public string Homologacija; public string ImeVL; public string PrezimeVL; public string MaticniBrojVL; public string BRLKVL; public string BrojPosteVL;
        public string MestoVL;public string UlicaVL; public string BrojVL; public string TelefonaVL; public string ZiroRacunVL; public string ImeDL; public string PrezimeDL;
        public string MaticniBrojDL; public string BRLKDL; public string BrojPosteDL;public string MestoDL; public string UlicaDL; public string BrojDL; public string TelefonaDL; 
        public string Napomena; public string Datum; public string VremeOd;public string VremeDo;
        public string Kontrolor1;public string Kontrolor2; public string Predjenikm;
        public string PoslednjeUverenjeBroj; public string PoslednjeUverenjeDatum;public string DecimalnaOznaka;public string OblikKaroserijeSkraceno; public string BojaL;
        public string BojaD; public string Ispravnost; public string BrSati; public string Kuka; public string OpterecenjeN1; public string OpterecenjeN2;
        public string OpterecenjeN3; public string OpterecenjeN4; public string OpterecenjeN5; public string PneumatikL1; public string PneumatikL2; public string PneumatikL3;
        public string PneumatikL4; public string PneumatikL5; public string PneumatikD1; public string PneumatikD2; public string PneumatikD3; public string PneumatikD4; public string PneumatikD5;
        public string zaustavljanje; public string terensko; public string komentar; public string rejting; public string decplusoznaka;public string duzina;public string sirina;public string visina;
        public string BrojDokumenta;public string SerijskiBroj;public string Poruka;public string VlaznostKT;public string DatumIzdavanjaSaobracajne;

        public string datumIsteka { get; internal set; }
        public string Prioritet { get; internal set; }
        public string DLZiroRacun { get; internal set; }
        public string Osiguranje { get; internal set; }

        public override string ToString()
        {
            return RegistarskiBroj; ;
        }
    }
}
