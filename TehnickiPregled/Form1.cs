using System;
using System.Windows.Forms;

namespace TehnickiPregled
{
    public partial class Form1 : Form
    {
        public vozilo V { get; set; }
        public Form1()
        {
            InitializeComponent();
            V = new vozilo()
            {
                ImeVL = "Safet",
                PrezimeVL = "Purkovic",
                MaticniBrojVL = "3007990123123",
                UlicaVL = "Lenjinova bb",
                MestoVL = "Sjenica",
                BrojVL = "BB",
                RegistarskiBroj = "NP064-EX",
                BrojDokumenta = "123123",
                SerijskiBroj = "1231231230000123"
            };
            OpenSite();
        }
        private void OpenSite()
        {
            try
            {
                webBrowser1.ScriptErrorsSuppressed = true;
                webBrowser1.AllowNavigation = true;
                //webBrowser1.Navigate(@"https://plati.euprava.gov.rs/");
                
                webBrowser1.Navigate(@"https://google.com/");
                webBrowser1.ScriptErrorsSuppressed = true;
            }
            catch
            {
                MessageBox.Show("Doslo je do greske prilikom otvaranja sajta plati.euprava.gov.rs");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PopuniLicnePodatke();
        }

        private void PopuniLicnePodatke()
        {
            webBrowser1.Document.GetElementsByTagName("input")[0].InnerHtml = V.ImeVL + " " + V.PrezimeVL;
            webBrowser1.Document.GetElementsByTagName("input")[1].InnerHtml = V.MaticniBrojVL;
            webBrowser1.Document.GetElementsByTagName("input")[2].InnerHtml = V.UlicaVL + " " + V.BrojVL + "," + V.MestoVL;
        }
        private void PopuniPodatkeOVozilu()
        {
            webBrowser1.Document.GetElementsByTagName("input")[0].InnerHtml = V.RegistarskiBroj;
            webBrowser1.Document.GetElementsByTagName("input")[1].InnerHtml = V.BrojDokumenta;
            webBrowser1.Document.GetElementsByTagName("input")[2].InnerHtml = V.SerijskiBroj;
        }
    }
}
