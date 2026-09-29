namespace TehnickiPregled
{
    public class Faktura
    {
        public string NazivFirme { get; set; }
        public string Ulica { get; set; }
        public string Grad { get; set; }
        public string PIB { get; set; }
        public string Telefon { get; set; }
        public string TR { get; set; }
        public string MaticniBroj { get; set; }
        public string Rabat { get; set; }
        public string NazivArtikla { get; set; }
        public string Cena { get; set; }
        public string Komada { get; set; }
        public override string ToString()
        {
            return NazivFirme;
        }
    }
}