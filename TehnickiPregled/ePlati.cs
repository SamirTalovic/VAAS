using System;
using System.Windows.Forms;

namespace TehnickiPregled
{
    public partial class ePlati : Form
    {
        public vozilo V { get; set; }
        public ePlati(vozilo v)
        {
            InitializeComponent();
            V = v;
            OpenSite();
        }
        private async void OpenSite()
        {
            try
            {
                await webView.EnsureCoreWebView2Async(null);
                //await webView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync("window.addEventListener(\'message\', event => alert(event.data));");
                webView.CoreWebView2.Navigate(@"https://plati.euprava.gov.rs/");
                //https://app.payspot.rs/
            }
            catch
            {
                MessageBox.Show("Doslo je do greske prilikom otvaranja sajta plati.euprava.gov.rs");
            }
        }
        private void PopuniLicnePodatke()
        {
            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[0].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            if(V.PrezimeDL != "")
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.ImeDL} {V.PrezimeDL}');");
            else
                webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.ImeDL}');");

            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[1].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.MaticniBrojDL}');");

            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[2].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.UlicaDL} {V.BrojDL} {V.MestoDL}');");
        }
        private void PopuniPodatkeOVozilu()
        {
            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[0].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.RegistarskiBroj}');");

            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[1].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.BrojDokumenta}');");

            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[2].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.SerijskiBroj}');");
        }
        private void bunifuThinButton24_Click(object sender, EventArgs e)
        {
            PopuniLicnePodatke();
        }
        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {
            PopuniPodatkeOVozilu();
        }

        private void bunifuThinButton22_Click(object sender, EventArgs e)
        {
            PopuniPodatkeOVoziluIzUvoza();
        }

        private void PopuniPodatkeOVoziluIzUvoza()
        {
            //txtGorivo.Text = "Gorivo : " + V.Gorivo + Environment.NewLine +" Vrsta vozila " + V.VrstaVozila;
            //txtGorivo.Visible = true;
            
            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[6].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.Zapremina}');");

            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[7].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.Sedenje}');");

            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[8].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.NajvecaMasa}');");

            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[9].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.Nosivost}');");

            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[10].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.GodinaProizvodnje}');");

            webView.ExecuteScriptAsync("document.getElementsByTagName('input')[11].focus();");
            webView.ExecuteScriptAsync("document.execCommand('selectAll', false);");
            webView.ExecuteScriptAsync($"document.execCommand('insertText', false, '{V.Snaga}');");
        }
    }
}
