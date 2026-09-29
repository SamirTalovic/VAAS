using BashChelik;
using SaobracajnaNET;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace TehnickiPregled
{
    //public class DownloadPageAsyncResult
    //{
    //    public string result { get; set; }
    //    public string reasonPhrase { get; set; }
    //    public HttpResponseHeaders headers { get; set; }
    //    public HttpStatusCode code { get; set; }
    //    public string errorMessage { get; set; }
    //}

    public partial class ucRegistracija : UserControl
    {
        Broker b = new Broker();
        int brOtpremnice = 0;
        DataTable dt = new DataTable();
        DataTable dt3 = new DataTable();
        int duzinaTehPregleda = 0;
        int i = 0, suma = 0;
        double kolicina = 1.0;
        double ukupno = 0, ukupno_bez_rabata = 0, ukupan_rabat = 0, ukupan_pdv = 0;
        string Brpregleda;
        int strana = 1;
        bool prvoUcitavanje = true;
        bool vrstaVozilaPrvo = true;
        DataTable dt2 = new DataTable();
        DialogResult ispunjenostUslova;
        string ispravnost = "Da";
        string terensko = "Ne";
        bool cbTabliceEnter = false;
        bool cbSasijeEnter = false;
        bool cbFaktureEnter = false;
        bool cbIdjeviEnter = false;
        int brNeispravnost = 0;
        public ucRegistracija()
        {

            //ucitavanje vremena sa interneta
            //try
            //{
            //var client = new TcpClient("time.nist.gov", 13);
            //var streamReader = new StreamReader(client.GetStream());
            //var response = streamReader.ReadToEnd();
            //var utcDateTimeString = response.Substring(7, 17);
            //var localDateTime = DateTime.ParseExact(utcDateTimeString, "yy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal);
            // if (localDateTime.Date < DateTime.Parse("31.12.2020"))
            //{

            InitializeComponent();
            dt2.Columns.Add("Ime i prezime vlasnika", typeof(string));
            dt2.Columns.Add("Registarski broj", typeof(string));
            dt2.Columns.Add("Marka", typeof(string));
            dt2.Columns.Add("Model", typeof(string));
            dt2.Columns.Add("Ime i prezime stranke", typeof(string));
            dt2.Columns.Add("Broj telefona", typeof(string));
            dt2.Columns.Add("Datum isteka", typeof(string));

            dt3.Columns.Add("cbPoruka", typeof(bool));
            dt3.Columns.Add("Ime i prezime vlasnika", typeof(string));
            dt3.Columns.Add("RegistarskiBroj", typeof(string));
            dt3.Columns.Add("Ime i prezime stranke", typeof(string));
            dt3.Columns.Add("Broj telefona", typeof(string));
            dt3.Columns.Add("Datum isteka", typeof(string));
            dt3.Columns.Add("IDBroj", typeof(string));

            cbOsiguranje.Items.Add("KOMPANIJA DUNAV OSIGURANJE");
            cbOsiguranje.Items.Add("DDOR NOVI SAD");
            cbOsiguranje.Items.Add("TRIGLAV OSIGURANJE");
            cbOsiguranje.Items.Add("MILENIJUM OSIGURANJE");
            cbOsiguranje.Items.Add("TAKOVO OSIGURANJE");
            cbOsiguranje.Items.Add("AMS OSIGURANJE");
            cbOsiguranje.Items.Add("UNIQA NEŽIVOTNO OSIGURANJE");
            cbOsiguranje.Items.Add("GENERALI OSIGURANJE SRBIJA");
            cbOsiguranje.Items.Add("SAVA OSIGURANJE");
            cbOsiguranje.Items.Add("WIENER STADTISCHE OSIGURANJE");
            cbOsiguranje.Items.Add("AS NEŽIVOTNO OSIGURANJE");
            cbOsiguranje.Items.Add("BASLER NEŽIVOTNO OSIGURANJE");
            cbOsiguranje.Items.Add("AXA OSIGURANJE");
            cbOsiguranje.Items.Add("GRAWE OSIGURANJE");
            cbOsiguranje.Items.Add("GLOBOS OSIGURANJE");
            cbOsiguranje.SelectedIndex = 0;


            dtpOd.Value = DateTime.Now;
            dtpDo.Value = DateTime.Now.AddDays(7);

            var vozilas2 = new List<vozilo>();
            var dolazilaVozila = b.ucitajTabliceDolazili(dtpOd.Value.AddDays(-21).ToString("dd/MM/yyyy"), dtpOd.Value.AddDays(1).ToString("dd/MM/yyyy"));
            if (DateTime.Now.DayOfWeek == DayOfWeek.Monday)
                vozilas2 = b.ucitajTabliceIsticu(dtpDo.Value.AddDays(-2).ToString("dd/MM/yyyy"), dtpDo.Value.ToString("dd/MM/yyyy"));
            else
                vozilas2 = b.ucitajTabliceIsticu(dtpDo.Value.ToString("dd/MM/yyyy"), dtpDo.Value.ToString("dd/MM/yyyy"));
            var vozilas = vozilas2.Where(x => x.IDBroj[8] == '1' && x.Ispravnost == "Da");
            dt3.Clear();
            //vozilas.Reverse();
            var vozila = vozilas.GroupBy(x => new { x.ImeVL, x.RegistarskiBroj }).Select(y => y.First());
            //TimeSpan razlika2 = new TimeSpan(365, 0, 0, 0);
            int j = 0;
            foreach (var auto in vozila)
            {
                //DateTime datum;
                try
                {
                    //datum = DateTime.ParseExact(auto.Datum, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                    //if ((datum + razlika2) >= dtpOd.Value && (datum+razlika2) <= dtpDo.Value)
                    var dolazio = dolazilaVozila.Where(a => a.BrojSasije == auto.BrojSasije).FirstOrDefault();
                    if (dolazio != null) auto.Poruka = "Da";
                    dgwPoruke.Rows.Add((auto.Poruka != "Da") ? true : false, auto.ImeVL + " " + auto.PrezimeVL, auto.RegistarskiBroj, auto.ImeDL + " " + auto.PrezimeDL, (auto.TelefonaDL != "") ? auto.TelefonaDL : auto.TelefonaVL, auto.datumIsteka, auto.IDBroj);
                    if (auto.Poruka == "Da")
                    {
                        dgwPoruke.Rows[j].Cells[0].Style.BackColor = Color.FromArgb(229, 226, 244);
                        dgwPoruke.Rows[j].Cells[1].Style.BackColor = Color.FromArgb(229, 226, 244);
                        dgwPoruke.Rows[j].Cells[2].Style.BackColor = Color.FromArgb(229, 226, 244);
                        dgwPoruke.Rows[j].Cells[3].Style.BackColor = Color.FromArgb(229, 226, 244);
                        dgwPoruke.Rows[j].Cells[4].Style.BackColor = Color.FromArgb(229, 226, 244);
                        dgwPoruke.Rows[j].Cells[5].Style.BackColor = Color.FromArgb(229, 226, 244);
                    }
                }
                catch { }
                j++;
            }
            //dgwPoruke.DataSource = null;
            //dgwPoruke.DataSource = dt3;

            txtTPDatum.Text = DateTime.Now.ToString("dd.MM.yyyy");
            datumIsteka.Text = (DateTime.Now).AddYears(1).ToString("dd.MM.yyyy");
            txtUplatniceDatum.Text = DateTime.Now.ToString("dd.MM.yyyy");
            textBox10.Text = b.nadjiMarginuLevo();
            textBox11.Text = b.nadjiMarginuDesno();
            brOtpremnice = b.nadjiOtpremnicu();
            txtOtp.Text = brOtpremnice.ToString();
            //podesavanje cb autocomplete
            cbTPKontr1.DropDownStyle = ComboBoxStyle.DropDown;
            cbTPKontr1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbTPKontr1.AutoCompleteSource = AutoCompleteSource.ListItems;

            cbTPKontr2.DropDownStyle = ComboBoxStyle.DropDown;
            cbTPKontr2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbTPKontr2.AutoCompleteSource = AutoCompleteSource.ListItems;

            cbTablice.DropDownStyle = ComboBoxStyle.DropDown;
            cbTablice.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbTablice.AutoCompleteSource = AutoCompleteSource.ListItems;

            cbFakture.DropDownStyle = ComboBoxStyle.DropDown;
            cbFakture.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbFakture.AutoCompleteSource = AutoCompleteSource.ListItems;

            cbSasije.DropDownStyle = ComboBoxStyle.DropDown;
            cbSasije.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbSasije.AutoCompleteSource = AutoCompleteSource.ListItems;

            cbIdjevi.DropDownStyle = ComboBoxStyle.DropDown;
            cbIdjevi.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbIdjevi.AutoCompleteSource = AutoCompleteSource.ListItems;
            //cbMusterije.DropDownStyle = ComboBoxStyle.DropDown;
            //cbMusterije.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            //cbMusterije.AutoCompleteSource = AutoCompleteSource.ListItems;

            //ucitavanje tablica, sasije i idjeva
            cbTablice.DataSource = b.ucitajTablice();
            cbFakture.DataSource = b.ucitajFakture();
            cbSasije.DataSource = b.ucitajSasije();
            cbIdjevi.DataSource = b.ucitajIdjeve();
            prvoUcitavanje = false;
            cbTablice.SelectedIndex = -1;
            cbSasije.SelectedIndex = -1;
            cbIdjevi.SelectedIndex = -1;
            cbIdjevi.Text = "141022" + DateTime.Now.Year % 100;
            //podesavanje za kuku da je Ne
            cbVKuka.SelectedIndex = 1;

            //ucitavanje kontrolora
            cbTPKontr1.DataSource = b.nadjiKotrolore();
            cbTPKontr1.SelectedIndex = 0;
            cbTPKontr2.DataSource = b.nadjiKotrolore();
            cbTPKontr2.SelectedIndex = 1;

            //podesavanje datetimepickera
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd.MM.yyyy";

            //podesavanje polja na fakturi
            txtNaziv.Text = "Tehnički pregled vozila";
            //txtSfr.Text = "3000";
            txtRabat.Text = "0";

            //podesavanje datatable za fakturu
            dt.Columns.Add("Šifra", typeof(int));
            dt.Columns.Add("Naziv", typeof(string));
            dt.Columns.Add("JM", typeof(string));
            dt.Columns.Add("Količina", typeof(double));
            dt.Columns.Add("Cena", typeof(double));
            dt.Columns.Add("Iznos", typeof(double));
            dt.Columns.Add("Rabat %", typeof(double));
            dt.Columns.Add("Rabat", typeof(double));
            dt.Columns.Add("PDV %", typeof(string));
            dt.Columns.Add("PDV", typeof(double));
            dt.Columns.Add("Cena sa PDV", typeof(double));
            dt.Columns.Add("Iznos sa PDV", typeof(double));
            dataGridView2.DataSource = dt;

            //podesavanje kontrolora
            cbTPKontr1.SelectedIndex = 0;
            cbTPKontr2.SelectedIndex = 1;

            //podesavanje praznog polja za vrijeme
            txtTPVremeOd.CustomFormat = " ";
            txtTPVremeDo.CustomFormat = " ";
            txtTPVremeDo.Format = DateTimePickerFormat.Custom;
            txtTPVremeOd.Format = DateTimePickerFormat.Custom;


            //pravljenje id broja
            try
            {
                string sifra = "141022";
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 6; i++)
                    sb.Append(sifra[i]);
                sifra = (DateTime.Now.Year % 100).ToString();
                for (int i = 0; i < 2; i++)
                    sb.Append(sifra[i]);
                sb.Append(1);
                string poslednji = b.nadjiPoslednjiPregled();
                int sledeci = int.Parse(poslednji);
                sledeci++;
                poslednji = sledeci.ToString().PadLeft(5, '0');
                Brpregleda = poslednji;
                for (int i = 0; i < 5; i++)
                    sb.Append(poslednji[i]);
                txtTPID.Text = sb.ToString();
            }
            catch { }

            //odabir prvog id broja
            comboBox2.SelectedIndex = 0;

            //dodavanje podataka u cb za neispravnosti
            List<string> lista = b.lista_grupa();
            cbGlPodela.DataSource = lista;
            cbGlPodela.Text = "";

            popuniUplatnice();

            //dodavanje marki vozila
            //List<string> marke = new List<string>();
            //foreach (var b in b.lista_marka())
            //{
            //    marke.Add(b.KaroserijaNaziv);
            //}
            //cbVMarka.DataSource = marke;
            //cbVMarka.Text = "";
            //}
            //else MessageBox.Show("Windows couldn’t complete the requested changes.The changes couldn’t be completed.\nPlease connect to the Internet and try again.Error code: 0x800F0954");

            //}
            //catch (Exception)
            //{

            //    MessageBox.Show("Windows couldn’t complete the requested changes.The changes couldn’t be completed.\nPlease connect to the Internet and try again.Error code: 0x800F0954");
            //}


            UkljuciIskljuciNeispravnosti();
            cbFakture.SelectedIndex = 0;
        }

        private void popuniUplatnice()
        {
            comboBox1.Items.Add("Zbirna uplatnica");
            comboBox1.Items.Add("Registraciona nalepnica (ZAVOD ZA IZRADU NOV. ...)");
            comboBox1.Items.Add("Registraciona nalepnica (RAT)");
            comboBox1.Items.Add("Saobraćajna dozvola (ZAVOD ZA IZRADU NOV. ...)");
            comboBox1.Items.Add("Saobraćajna dozvola (RAT)");
            comboBox1.Items.Add("Registarske tablice (ZAVOD ZA IZRADU NOV. ...)");
            comboBox1.Items.Add("Registarske tablice (RAT)");
            comboBox1.Items.Add("Komunalna taksa");
            comboBox1.Items.Add("Porez za upotrebu m/v ");
            comboBox1.Items.Add("Uplata premije za osiguranje");
            comboBox1.Items.Add("Zeleni karton");
            comboBox1.Items.Add("Tehnički pregled vozila");


            checkBox3.Text = ("Zbirna uplatnica");
            checkBox5.Text = ("Registraciona nalepnica (ZAVOD ZA IZRADU NOV. ...)");
            checkBox6.Text = ("Registraciona nalepnica (RAT)");
            checkBox7.Text = ("Saobraćajna dozvola (ZAVOD ZA IZRADU NOV. ...)");
            checkBox8.Text = ("Saobraćajna dozvola (RAT)");
            checkBox9.Text = ("Registarske tablice (ZAVOD ZA IZRADU NOV. ...)");
            checkBox10.Text = ("Registarske tablice (RAT)");
            checkBox11.Text = ("Komunalna taksa");
            checkBox13.Text = ("Uplata premije za osiguranje");
            checkBox14.Text = ("Porez za upotrebu m/v");
            checkBox16.Text = ("Tehnički pregled vozila");
            checkBox15.Text = ("Zeleni karton");

        }

        //stampanje zapisnika
        private void bunifuTileButton1_Click(object sender, EventArgs e)
        {
            if (txtVKM.Text != "" && cbVrstaVoz.Text != "" && (checkBox12.Checked && cbVOblikKar.Text == "" ? true : (cbVOblikKar.Text == "" ? false : true)) && cbVMarka.Text != "" && txtVModel.Text != "" && txtVGodinaProizvodnje.Text != ""
                && txtVBrojSasije.Text != "" && txtVBrojMotora.Text != "" && cbVBoja.Text != "" && cbBojaL.Text != "" && cbBojaD.Text != "" && txtVPrvaReg.Text != ""
                && txtVSnaga.Text != "" && txtVZapremina.Text != "" && txtVNajMasa.Text != "" && txtVTezina.Text != "" && txtVSedenje1.Text != "" && txtVStajanje.Text != ""
                && txtVBrojOsovina.Text != "" && cbVGorivo.Text != "" && txtTPID.Text != "" && ispravnost != "" && cbVPV.Text != "" && (checkBox12.Checked && cbVOKS.Text == "" ? true : (cbVOKS.Text == "" ? false : true))
                && txtTPVremeOd.Text != "  :" && txtTPVremeDo.Text != "  :")
            {
                int n;

                if (dataGridView1.RowCount <= 1 && rbTehNeis.Checked == true)
                {
                    tabControl1.SelectedTab = tabPage3;
                    MessageBox.Show("Tehnička ispranost vozila je obeležena sa \'Ne\', unesite neispravnosti!");
                }
                else if (dataGridView1.RowCount > 1 && rbTehIsp.Checked == true)
                {
                    tabControl1.SelectedTab = tabPage3;
                    MessageBox.Show("Tehnička ispranost vozila je obeležena sa \'Da\', izbrišite unete neispravnosti!");
                }

                else if (txtTPID.Text.Length < 14 || txtTPID.Text.Length > 14)
                {
                    MessageBox.Show("ID Broj je nepravilan!");
                }
                else if (int.TryParse(txtTPID.Text.Substring(9), out n) == false)
                {
                    MessageBox.Show("ID Broj je nepravilan!");
                }

                else
                {
                    lblVrstaVoz.ForeColor = SystemColors.ControlText; lblVOblikKar.ForeColor = SystemColors.ControlText; lblVMarka.ForeColor = SystemColors.ControlText;
                    lblVModel.ForeColor = SystemColors.ControlText; lblVGodinaProizvodnje.ForeColor = SystemColors.ControlText; lblVBrojSasije.ForeColor = SystemColors.ControlText;
                    lblVBrojMotora.ForeColor = SystemColors.ControlText; lblVBoja.ForeColor = SystemColors.ControlText; lblVPrvaReg.ForeColor = SystemColors.ControlText;
                    lblVSnaga.ForeColor = SystemColors.ControlText; lblVZapremina.ForeColor = SystemColors.ControlText; lblVNajMasa.ForeColor = SystemColors.ControlText;
                    lblVTezina.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText;
                    lblVBrojOsovina.ForeColor = SystemColors.ControlText; lblVGorivo.ForeColor = SystemColors.ControlText; lblTPID.ForeColor = SystemColors.ControlText;
                    lblIspravnost.ForeColor = SystemColors.ControlText; lblPredjenikm.ForeColor = SystemColors.ControlText; lblVremeOd.ForeColor = SystemColors.ControlText;
                    lblVremeDo.ForeColor = SystemColors.ControlText;
                    i = 0;
                    int brListova = brNeispravnost / 10 + 1;
                    if (brNeispravnost % 10 == 0)
                        brListova--;
                    if (rbTehIsp.Checked)
                        brListova = 1;
                    for (int k = 0; k < brListova; k++)
                    {
                        //ispunjenostUslova = MessageBox.Show("Da li su ispunjeni uslovi za vršenje tehničkog pregleda?", "Ispunjenost uslova", MessageBoxButtons.YesNo);

                        PrintDialog printDialog = new PrintDialog();

                        PrintDocument printDocument = new PrintDocument();

                        printDialog.Document = printDocument; //add the document to the dialog box...        

                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt); //add an event handler that will do the printing

                        //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

                        DialogResult result = printDialog.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            printDocument.Print();

                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Popunite sva obavezna polja na tabovima \'Podaci o vozilu\' i \'Tehnički pregled\'.");

                lblVrstaVoz.ForeColor = Color.Red; lblVOblikKar.ForeColor = Color.Red; lblVMarka.ForeColor = Color.Red; lblVModel.ForeColor = Color.Red;
                lblVGodinaProizvodnje.ForeColor = Color.Red; lblVBrojSasije.ForeColor = Color.Red; lblVBrojMotora.ForeColor = Color.Red; lblVBoja.ForeColor = Color.Red;
                lblVPrvaReg.ForeColor = Color.Red; lblVSnaga.ForeColor = Color.Red; lblVZapremina.ForeColor = Color.Red; lblVNajMasa.ForeColor = Color.Red;
                lblVTezina.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVBrojOsovina.ForeColor = Color.Red;
                lblVGorivo.ForeColor = Color.Red; lblTPID.ForeColor = Color.Red; lblPredjenikm.ForeColor = Color.Red; lblIspravnost.ForeColor = Color.Red;
                lblVremeOd.ForeColor = Color.Red; lblVremeDo.ForeColor = Color.Red;
            }


        }

        private void btnPrintReciept_Click(object sender, EventArgs e)
        {

        }

        //kreiranje zapisnika
        public void CreateReceipt(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int height = 0;
            int xpos = 17, ypos = 35;
            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.5f);
            Font FontNaznacenimanji = new Font("Times New Roman", 10, FontStyle.Bold);
            Font FontNaznaceni = new Font("Times New Roman", 11, FontStyle.Bold);
            Font FontNajjaci = new Font("Times New Roman", 14, FontStyle.Bold | FontStyle.Underline);
            Font Fontoznaka = new Font("Times New Roman", 10);
            Font Fontobicni = new Font("Times New Roman", 10);
            Font Fontobicni2 = new Font("Times New Roman", 10);
            Font fontUpisivanje = new Font("Times New Roman", 8, FontStyle.Italic);
            Font za_licnu_kartu = new Font("Times New Roman", 8);
            Font FontSvoji = new Font("Times New Roman", 14);
            Font fontmali = new Font("Times New Roman", 7);
            Font fontnajmanji = new Font("Times New Roman", 5);

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;

            StringFormat stringFormat2 = new StringFormat();
            stringFormat2.Alignment = StringAlignment.Near;
            stringFormat2.LineAlignment = StringAlignment.Center;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;

            //stampanje svojih podataka
            if (strana == 1)
            {
                e.Graphics.DrawString("VAAS DOO Sjenica", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 400, 25));
                e.Graphics.DrawString(txtTPRazlikaVremena.Text, fontnajmanji, Brushes.Black, new Rectangle(xpos + 5, ypos + 17, 400, 25));
                e.Graphics.DrawString("1 4 1 0 2 2", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 794, 20), stringFormat);
                e.Graphics.DrawString("Образац 3\nстрана 1", Fontoznaka, Brushes.Black, new Rectangle(xpos, ypos, 755, 30), stringFormatdesno);

                ypos += 40;
                e.Graphics.DrawString("ЗАПИСНИК О ВРШЕЊУ ТЕХНИЧКОГ ПРЕГЛЕДА ВОЗИЛА", FontNajjaci, Brushes.Black, new Rectangle(xpos, ypos, 794, 20), stringFormat);
                ypos += 50;
                int pom = ypos;
                //stampanje podataka o kupcu
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 120, 20));
                e.Graphics.DrawString("ИД број", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 120, ypos, 230, 20));
                e.Graphics.DrawString(txtTPID.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 120, ypos, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 120, 20));
                e.Graphics.DrawString("Датум", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 120, ypos, 230, 20));
                e.Graphics.DrawString(txtTPDatum.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 120, ypos, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 120, 20));
                e.Graphics.DrawString("Време почетка", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 120, ypos, 230, 20));
                e.Graphics.DrawString(txtTPVremeOd.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 120, ypos, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 120, 20));
                e.Graphics.DrawString("Време завршетка", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 120, ypos, 230, 20));
                e.Graphics.DrawString(txtTPVremeDo.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 120, ypos, 400, 20));
                ypos += 30;
                e.Graphics.DrawString("Испуњеност услова за вршење техничког прегледа возила", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 175, 20));
                e.Graphics.DrawString("ДА", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 20), stringFormat);
                //e.Graphics.DrawEllipse(pe,new Rectangle(80,ypos,20,20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString("НЕ", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20), stringFormat);
                //if (ispunjenostUslova == DialogResult.Yes)
                //{
                //    e.Graphics.DrawEllipse(pe, new Rectangle(xpos + 77, ypos+1,23,18));
                //}
                //else
                //{
                //    e.Graphics.DrawEllipse(pe, new Rectangle(xpos + 78 + 175, ypos+1, 23, 18));
                //}
                if (ispravnost == "Da" || ispravnost == "Ne")
                {
                    e.Graphics.DrawEllipse(pe, new Rectangle(xpos + 77, ypos + 1, 23, 18));
                }
                else if (ispravnost == "Prekid")
                {
                    e.Graphics.DrawEllipse(pe, new Rectangle(xpos + 78 + 175, ypos + 1, 23, 18));
                }
                ypos += 50;
                e.Graphics.DrawString("ПОДАЦИ О СТРАНЦИ", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 175, 20));
                e.Graphics.DrawString("Име и презиме", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                if (rbVL.Checked)
                    e.Graphics.DrawString(txtVIme.Text + " " + txtVPrezime.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                else
                    e.Graphics.DrawString(txtDLIme.Text + " " + txtDLPrezime.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 175, 70));
                e.Graphics.DrawString("Регистарски број личне карте странке, односно другог документа за странце", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 70));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 70));

                if (rbVL.Checked)
                    e.Graphics.DrawString(txtVBRLK.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 70), stringFormat2);
                else
                    e.Graphics.DrawString(txtDLBRLK.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 70), stringFormat2);
                ypos += 100;
                e.Graphics.DrawString("ПОДАЦИ О ВОЗИЛУ", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("А", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Регистарска ознака", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(txtRegBr.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
                e.Graphics.DrawString("E", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
                e.Graphics.DrawString("Идентификациона (VIN)\n ознака ", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
                e.Graphics.DrawString(txtVBrojSasije.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 35), stringFormat2);
                ypos += 35;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("J", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Врста и ознака врсте", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(cbVrstaVoz.Text + " " + cbVPV.Text + ((cbTerensko.Checked) ? "G" : ""), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("J.1", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Облик каросерије", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                //string s="";

                //if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
                //{
                //    if (!checkBox12.Checked)
                //        s = cbVKlasa.Text + "/" + cbVOKS.Text;
                //    else
                //        s = cbVKlasa.Text + "/" + (cbVSpec.Text.Split(' '))[0].ToString();
                //}
                //else
                //     if (checkBox12.Checked)
                //{
                //    if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
                //        s = cbVKlasa.Text + "/" + (cbVSpec.Text.Split(' '))[0].ToString();
                //    else
                //        if (cbDecOzn.Text != "")
                //        s = (cbVSpec.Text.Split(' '))[0].ToString() + (cbDecOzn.Text.Split(' '))[0].ToString();
                //    else
                //        s = (cbVSpec.Text.Split(' '))[0].ToString();
                //    txtVrstaPlusDecOznk.Text = s;
                //}
                //else
                //    if (cbDecOzn.Text != "")
                //    s = cbVOKS.Text + (cbDecOzn.Text.Split(' '))[0].ToString();

                //else
                //    s = cbVOKS.Text;
                e.Graphics.DrawString(txtVrstaPlusDecOznk.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("K", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Број хомологације", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawString("1", fontmali, Brushes.Black, new Rectangle(xpos + 16 + 122, ypos - 1, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(txtVHomologacija.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("D.1", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Марка", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(cbVMarka.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
                e.Graphics.DrawString("D.2", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
                e.Graphics.DrawString("Тип/варијанта/верзија", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 35));
                e.Graphics.DrawString("1", fontmali, Brushes.Black, new Rectangle(xpos + 160, ypos - 1, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
                e.Graphics.DrawString(checkBox1.Checked ? txtVTip.Text : "", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 35));
                ypos += 35;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
                e.Graphics.DrawString("D.3", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
                e.Graphics.DrawString("Комерцијална ознака\n(модел)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
                e.Graphics.DrawString(txtVModel.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 35), stringFormat2);
                ypos += 35;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("B", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Датум прве регистр.", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(txtVPrvaReg.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
                e.Graphics.DrawString("F.1", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
                e.Graphics.DrawString("Највећа дозвољена\nмаса (kg)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 35));
                e.Graphics.DrawString("2", fontmali, Brushes.Black, new Rectangle(xpos + 4 + 25 + 60, ypos + 14, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
                e.Graphics.DrawString(txtVNajMasa.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 35), stringFormat2);
                ypos += 35;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("G", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Маса (kg)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(txtVTezina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("G.1", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Носивост (kg)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(txtVNosivost.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("L", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Број осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(txtVBrojOsovina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 100));
                e.Graphics.DrawString("N", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 100), stringFormat2);
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 100));
                e.Graphics.DrawString("Дозвољена осовинска оптерећења", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 100), stringFormat2);
                e.Graphics.DrawString("3", fontmali, Brushes.Black, new Rectangle(xpos + 4 + 25 + 72, ypos + 47, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 87, 20));
                e.Graphics.DrawString("N.1", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 87, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 20));
                e.Graphics.DrawString(txtVOpterecenjeN1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 87 + 175 + 5, ypos, 87, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 87, 20));
                e.Graphics.DrawString("N.2", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 87, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 20));
                e.Graphics.DrawString(txtVOpterecenjeN2.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 87 + 175 + 5, ypos, 87, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 87, 20));
                e.Graphics.DrawString("N.3", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 87, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 20));
                e.Graphics.DrawString(txtVOpterecenjeN3.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 87 + 175 + 5, ypos, 87, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 87, 20));
                e.Graphics.DrawString("N.4", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 87, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 20));
                e.Graphics.DrawString(txtVOpterecenjeN4.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 87 + 175 + 5, ypos, 87, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 87, 20));
                e.Graphics.DrawString("N.5", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 87, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 20));
                e.Graphics.DrawString(txtVOpterecenjeN5.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 87 + 175 + 5, ypos, 87, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("P.5", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Ознака мотора", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(txtVBrojMotora.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
                e.Graphics.DrawString("P.1", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
                e.Graphics.DrawString("Радна запремина\nмотора (cm  )", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 35));
                e.Graphics.DrawString("3", fontmali, Brushes.Black, new Rectangle(xpos + 25 + 70, ypos + 15, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
                e.Graphics.DrawString(txtVZapremina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 35), stringFormat2);
                ypos += 35;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("P.2", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Снага мотора (kW)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(txtVSnaga.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("P.3", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Погонско гориво", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawString("4", fontmali, Brushes.Black, new Rectangle(xpos + 4 + 18 + 110, ypos - 1, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(cbVGorivo.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
                e.Graphics.DrawString("Q", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
                e.Graphics.DrawString("Однос снага/маса\n (kw/kg)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 35));
                e.Graphics.DrawString("5", fontmali, Brushes.Black, new Rectangle(xpos + 2 + 25 + 50, ypos + 15, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
                e.Graphics.DrawString(txtVOdnos.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25 + 150, ypos, 175, 35), stringFormat2);

                ypos += 35;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("R", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Боја возила", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(cbBojaL.Text.Substring(0, 1) + " " + cbVBoja.Text.ToString() + " " + cbBojaD.Text.Substring(0, 1), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos = pom;
                xpos += 2 + 175 + 175 + 50;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("S.1", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Број места за седење", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(txtVSedenje1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
                e.Graphics.DrawString("S.2", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Број места за стајање", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString(txtVStajanje.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
                e.Graphics.DrawString("B.1", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
                e.Graphics.DrawString("Година производње\nвозила", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
                e.Graphics.DrawString(txtVGodinaProizvodnje.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 35), stringFormat2);
                ypos += 35;

                //System.Drawing.StringFormat drawFormat = new System.Drawing.StringFormat();
                //drawFormat.FormatFlags = StringFormatFlags.DirectionVertical;
                e.Graphics.TranslateTransform(xpos, ypos);

                e.Graphics.RotateTransform(-90);
                e.Graphics.DrawString("Ознака пнеуматика", Fontobicni2, Brushes.Black, new Rectangle(-120 + 1, 7, 200, 200));
                e.Graphics.DrawString("6", fontmali, Brushes.Black, new Rectangle(-10, -1, 175, 20));
                e.Graphics.ResetTransform();
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 120));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("Лево", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawString("Десно", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("1. " + txtVPneumatikL1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString("1. " + txtVPneumatikD1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("2. " + txtVPneumatikL2.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString("2. " + txtVPneumatikD2.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("3. " + txtVPneumatikL3.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString("3. " + txtVPneumatikD3.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("4. " + txtVPneumatikL4.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString("4. " + txtVPneumatikD4.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
                e.Graphics.DrawString("5. " + txtVPneumatikL5.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString("5. " + txtVPneumatikD5.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 60));
                e.Graphics.DrawString("DV", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 60), stringFormat2);
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 60));
                e.Graphics.DrawString("Димензије(m)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175, 60), stringFormat2);
                e.Graphics.DrawString("7", fontmali, Brushes.Black, new Rectangle(xpos + 25 + 20 + 69, ypos + 20, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 87, 20));
                e.Graphics.DrawString("D", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 87, 20));
                e.Graphics.DrawString(txtVD.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175 + 87, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 87, 20));
                e.Graphics.DrawString("Š", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 87, 20));
                e.Graphics.DrawString(txtVS.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175 + 87, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 87, 20));
                e.Graphics.DrawString("V", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 87, 20));
                e.Graphics.DrawString(txtVV.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175 + 87, ypos, 175, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 175 + 87, 35));
                e.Graphics.DrawString("Постојање уређаја за спајање вучног и прикључног возила", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175 + 87, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 44, 35));
                e.Graphics.DrawString("ДА", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175 + 87, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87 + 44, ypos, 44, 35));
                e.Graphics.DrawString("НЕ", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175 + 87 + 44, ypos, 175, 35));
                if (cbVKuka.Text == "Da")
                {
                    e.Graphics.DrawEllipse(pe, new Rectangle(xpos + 77 + 187, ypos + 1, 23, 18));
                }
                else
                {
                    e.Graphics.DrawEllipse(pe, new Rectangle(xpos + 78 + 187 + 44, ypos - 1, 23, 18));
                }
                //
                //e.Graphics.DrawString("ДА", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 20), stringFormat);
                ////e.Graphics.DrawEllipse(pe,new Rectangle(80,ypos,20,20));
                //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                //e.Graphics.DrawString("НЕ", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20), stringFormat);

                //
                ypos += 35;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 175 + 87, 20));
                e.Graphics.DrawString("Пређени пут (km)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175 + 87, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 20));
                e.Graphics.DrawString(txtVKM.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175 + 87, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 175 + 87, 20));
                e.Graphics.DrawString("Број радних сати", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175 + 87, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 20));
                e.Graphics.DrawString(txtVBrSati.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175 + 87, ypos, 175, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
                e.Graphics.DrawString("UZ", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150 + 87, 35));
                e.Graphics.DrawString("Врста преносног механизма\nуређаја за заустављање", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 25, ypos, 175 + 87, 35));
                e.Graphics.DrawString("8", fontmali, Brushes.Black, new Rectangle(xpos + 57 + 20 + 100, ypos + 15, 175, 20));
                e.Graphics.DrawString(cbUredjajZaZaust.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175 + 87, ypos, 90, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175 + 87, ypos, 88, 35));
                //e.Graphics.DrawString(txtVZapremina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 88, 35), stringFormat2);
                ypos += 35;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 175, 54));
                e.Graphics.DrawString("Последње уверење о \nиспитивању возила\n(број/датум издавања)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175 + 87, 50));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 27));
                e.Graphics.DrawString(txtPoslednjeUverenjeBroj.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos + 2, 175, 35));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos + 27, 175, 27));
                e.Graphics.DrawString(txtPoslednjeUverenjeDatum.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos + 30, 175, 35));
                ypos += 75;
                e.Graphics.DrawString("ОЦЕНА ТЕХНИЧКЕ ИСПРАВНОСТИ ВОЗИЛА", FontNaznacenimanji, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 175, 20));
                e.Graphics.DrawString("ДА", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 20), stringFormat);
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
                e.Graphics.DrawString("НЕ", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 175, ypos, 175, 20), stringFormat);
                if (ispravnost == "Da")
                {
                    e.Graphics.DrawEllipse(pe, new Rectangle(xpos + 77, ypos + 1, 23, 18));
                }
                else if (ispravnost == "Ne")
                {
                    e.Graphics.DrawEllipse(pe, new Rectangle(xpos + 78 + 175, ypos + 1, 23, 18));
                }
                ypos += 40;
                e.Graphics.DrawString("СТЕПЕН НЕИСПРАВНОСТИ ВОЗИЛА", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 350, 20));
                e.Graphics.DrawString("", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 350, 20), stringFormat);
                ypos += 40;
                e.Graphics.DrawString("Уписује се:\n1 - за нова возила\n2 - за сва возила осим за мопеде, лаке трицикле, лаке четвороцикле,  мотоцикле,као и тракторе и радне машине код којих овај податак \nне постоји у произвођачкој таблици и документацији произвођача\n3 - за возила чије је осовинско оптерећење од сопствене масе веће од дозвољеног осовинског оптерећења\n4 - за моторна возила: бензин, дизел, етанол, ТНГ, КПГ, електро, \nбиодизел, водоник, мешавина(I)\n5 - за мотоцикле\n6 - на једном месту ако су пнеуматици исти на једној осовини, као и\n када су исти пнеуматици на свим осовинама\n7 - за возила чије су димензије веће од највећих дозвољених димензија\n8 - за теретна возила и аутобусе, чија највећа дозвољена маса\nпрелази 3.500kg и прикљична возила (пнеуматски, хидраулични,\nхидропнеуматски, инерциона команда, без уређаја за\nзаустављање остале врсте преносних механизама)", fontUpisivanje, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 400));
                ypos += 250;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 350, 80));
                e.Graphics.DrawString("Име и презиме контролора и потпис:", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 350, 20));
                ypos += 20;
                e.Graphics.DrawString(cbTPKontr1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 350, 20));
                ypos += 20;
                e.Graphics.DrawLine(pe, new Point(xpos + 5, ypos), new Point(xpos + 345, ypos));
                e.Graphics.DrawString(cbTPKontr2.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 350, 20));
                ypos += 20;
                e.Graphics.DrawLine(pe, new Point(xpos + 5, ypos), new Point(xpos + 345, ypos));
                strana++;
                e.HasMorePages = true;
                return;
            }
            else
            {
                xpos = 23; ypos = 35;
                //i = 0;
                p = new Pen(Brushes.Black, 1.0f);
                pe = new Pen(Brushes.Black, 0.5f);
                FontNaznacenimanji = new Font("Times New Roman", 10, FontStyle.Bold);
                FontNaznaceni = new Font("Times New Roman", 11, FontStyle.Bold);
                FontNajjaci = new Font("Times New Roman", 14, FontStyle.Bold | FontStyle.Underline);
                Fontoznaka = new Font("Times New Roman", 10);
                Fontobicni = new Font("Times New Roman", 9);
                Fontobicni2 = new Font("Times New Roman", 9, FontStyle.Bold);
                fontUpisivanje = new Font("Times New Roman", 6, FontStyle.Italic);
                za_licnu_kartu = new Font("Times New Roman", 8);
                FontSvoji = new Font("Times New Roman", 14);
                fontmali = new Font("Times New Roman", 6);

                stringFormat = new StringFormat();
                stringFormat.Alignment = StringAlignment.Center;
                stringFormat.LineAlignment = StringAlignment.Center;

                stringFormat2 = new StringFormat();
                stringFormat2.Alignment = StringAlignment.Near;
                stringFormat2.LineAlignment = StringAlignment.Center;

                stringFormatdesno = new StringFormat();
                stringFormatdesno.Alignment = StringAlignment.Far;
                stringFormatdesno.LineAlignment = StringAlignment.Near;


                //stampanje svojih podataka
                e.Graphics.DrawString("VAAS DOO Sjenica", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 400, 25));
                e.Graphics.DrawString("1 4 1 0 2 2", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 794, 20), stringFormat);
                e.Graphics.DrawString("Образац 3\nстрана 2", Fontoznaka, Brushes.Black, new Rectangle(xpos, ypos, 755, 30), stringFormatdesno);

                ypos += 40;
                e.Graphics.DrawString("ЗАПИСНИК О ВРШЕЊУ ТЕХНИЧКОГ ПРЕГЛЕДА ВОЗИЛА", FontNajjaci, Brushes.Black, new Rectangle(xpos, ypos, 794, 20), stringFormat);
                ypos += 50;
                int pom = ypos;
                e.Graphics.DrawString("ПОДАЦИ О ИЗВРШЕНИМ МЕРЕЊИМА ПОМОЋУ УРЕЂАЈА И ПРОРАЧУНИМА", Fontobicni2, Brushes.Black, new Rectangle(xpos + 10, ypos, 700, 20));
                ypos += 20;
                //prva tabela
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 124, 20));
                e.Graphics.DrawString("1. осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 134, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos, 124, 20));
                e.Graphics.DrawString("2. осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 258, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 20));
                e.Graphics.DrawString("3. осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 382, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawString("4. осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 506, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));
                e.Graphics.DrawString("5. осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 630, ypos, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 132, 40));
                e.Graphics.DrawString("Силе радног кочења", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 10, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 112, ypos, 20, 20));
                e.Graphics.DrawString("Л", Fontobicni, Brushes.Black, new Rectangle(xpos + 114, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 112, ypos + 20, 20, 20));
                e.Graphics.DrawString("Д", Fontobicni, Brushes.Black, new Rectangle(xpos + 114, ypos + 22, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos + 20, 124, 20));
                ypos += 40;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 132, 40));
                e.Graphics.DrawString("Овалност кочница", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 10, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 112, ypos, 20, 20));
                e.Graphics.DrawString("Л", Fontobicni, Brushes.Black, new Rectangle(xpos + 114, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 112, ypos + 20, 20, 20));
                e.Graphics.DrawString("Д", Fontobicni, Brushes.Black, new Rectangle(xpos + 114, ypos + 22, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos + 20, 124, 20));
                ypos += 40;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 132, 20));
                e.Graphics.DrawString("Разлике сила кочења", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 132, 20));
                e.Graphics.DrawString("Силе активирања", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 132, 20));
                e.Graphics.DrawString("Притисак ваздуха", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 132, 40));
                e.Graphics.DrawString("Силе паркирног", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 5, 400, 20));
                e.Graphics.DrawString("кочења", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 18, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 112, ypos, 20, 20));
                e.Graphics.DrawString("Л", Fontobicni, Brushes.Black, new Rectangle(xpos + 114, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 112, ypos + 20, 20, 20));
                e.Graphics.DrawString("Д", Fontobicni, Brushes.Black, new Rectangle(xpos + 114, ypos + 22, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos + 20, 124, 20));
                ypos += 40;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 132, 40));
                e.Graphics.DrawString("Силе помоћног", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 5, 400, 20));
                e.Graphics.DrawString("кочења", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 18, 400, 20));
                e.Graphics.DrawString("3", fontmali, Brushes.Black, new Rectangle(xpos + 42, ypos + 17, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 112, ypos, 20, 20));
                e.Graphics.DrawString("Л", Fontobicni, Brushes.Black, new Rectangle(xpos + 114, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 112, ypos + 20, 20, 20));
                e.Graphics.DrawString("Д", Fontobicni, Brushes.Black, new Rectangle(xpos + 114, ypos + 22, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos + 20, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos + 20, 124, 20));
                ypos += 40;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 132, 40));
                e.Graphics.DrawString("Кочни коефицијент", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 5, 400, 20));
                e.Graphics.DrawString("радног кочења", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 18, 400, 20));
                e.Graphics.DrawString("1", fontmali, Brushes.Black, new Rectangle(xpos + 80, ypos + 16, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 248, 40));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 40));
                e.Graphics.DrawString("Кочни коефицијент", Fontobicni, Brushes.Black, new Rectangle(xpos + 382, ypos + 5, 400, 20));
                e.Graphics.DrawString("паркирног кочења", Fontobicni, Brushes.Black, new Rectangle(xpos + 382, ypos + 18, 400, 20));
                e.Graphics.DrawString("2", fontmali, Brushes.Black, new Rectangle(xpos + 479, ypos + 16, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 248, 40));
                ypos += 40;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 132, 40));
                e.Graphics.DrawString("Влажност кочне", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 5, 400, 20));
                e.Graphics.DrawString("течности (%)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 18, 400, 20));
                e.Graphics.DrawString(txtVVlaznostKT.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 135, ypos + 11, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 248, 40));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 40));
                e.Graphics.DrawString("Брзина на почетку", Fontobicni, Brushes.Black, new Rectangle(xpos + 382, ypos + 3, 400, 20));
                e.Graphics.DrawString("мерења успорења", Fontobicni, Brushes.Black, new Rectangle(xpos + 382, ypos + 14, 400, 20));
                e.Graphics.DrawString("на полигону", Fontobicni, Brushes.Black, new Rectangle(xpos + 382, ypos + 24, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 248, 40));

                //fusnota
                ypos += 40;
                e.Graphics.DrawString("1- уписује   се   вредност   коефицијента,   осим   у   случају   контроле   прикључног   возила   на   полигон   када   се   уписује   ДА   када   је   систем   исправан   односно   НЕ   када   систем   није   исправан", fontUpisivanje, Brushes.Black, new Rectangle(xpos + 2, ypos + 5, 780, 20));
                e.Graphics.DrawString("2- уписује   се   вредност   коефицијента,   осим   у   случају   контроле   на   полигону   када   се   уписује   ДА   када   је   систем   исправан   односно   НЕ   када   систем   није исправан", fontUpisivanje, Brushes.Black, new Rectangle(xpos + 2, ypos + 15, 780, 20));
                e.Graphics.DrawString("3- утврђује   се   када   постоји   посебан   систем   помћног   кочења   или   када   су   помоћно   и   паркирно   кочење   изведени   као   један   систем", fontUpisivanje, Brushes.Black, new Rectangle(xpos + 2, ypos + 25, 780, 20));

                //druga tabela
                ypos += 45;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 124, 20));
                e.Graphics.DrawString("1. осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 134, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos, 124, 20));
                e.Graphics.DrawString("2. осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 258, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 20));
                e.Graphics.DrawString("3. осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 382, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawString("4. осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 506, ypos, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));
                e.Graphics.DrawString("5. осовина", Fontobicni, Brushes.Black, new Rectangle(xpos + 630, ypos, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 132, 20));
                e.Graphics.DrawString("Осовинско оптререћење", Fontobicni, Brushes.Black, new Rectangle(xpos, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 132, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 256, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 380, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));

                //treca tabela
                ypos += 30;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 260, 30));
                e.Graphics.DrawString("Брзина мопедаи возила која се производе\nу варијанти мопеда и моторцикла", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 1, 400, 30));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 492, 30));

                //cetvrta tabela
                ypos += 40;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 180, 80));
                e.Graphics.DrawString("Емисија издувних гасова мотора\nса активним паљењем", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 25, 400, 30));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 180, ypos, 150, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 330, ypos, 136, 20));
                e.Graphics.DrawString("Празан ход", Fontobicni, Brushes.Black, new Rectangle(xpos + 332, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 466, ypos, 136, 20));
                e.Graphics.DrawString("n≥2000 обртаја", Fontobicni, Brushes.Black, new Rectangle(xpos + 468, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 602, ypos, 150, 20));
                e.Graphics.DrawString("Радна темперетура мотора", Fontobicni, Brushes.Black, new Rectangle(xpos + 604, ypos + 2, 400, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 180, ypos, 150, 20));
                e.Graphics.DrawString("CO", Fontobicni, Brushes.Black, new Rectangle(xpos + 182, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 330, ypos, 136, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 466, ypos, 136, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 602, ypos, 150, 60));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 180, ypos, 150, 20));
                e.Graphics.DrawString("λ фактор", Fontobicni, Brushes.Black, new Rectangle(xpos + 182, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 330, ypos, 136, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 466, ypos, 136, 20));
                ypos += 20;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 180, ypos, 150, 20));
                e.Graphics.DrawString("Број обртаја мотора", Fontobicni, Brushes.Black, new Rectangle(xpos + 182, ypos + 2, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 330, ypos, 136, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 466, ypos, 136, 20));

                //peta tabela
                ypos += 30;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 141, 50));
                e.Graphics.DrawString("Димност издувних гасова\nдизел мотора", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 10, 400, 30));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 141, ypos, 80, 30));
                e.Graphics.DrawString("1. мерење", Fontobicni, Brushes.Black, new Rectangle(xpos + 143, ypos + 4, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 221, ypos, 80, 30));
                e.Graphics.DrawString("2. мерење", Fontobicni, Brushes.Black, new Rectangle(xpos + 223, ypos + 4, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 301, ypos, 80, 30));
                e.Graphics.DrawString("3. мерење", Fontobicni, Brushes.Black, new Rectangle(xpos + 303, ypos + 4, 400, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 381, ypos, 123, 30));
                e.Graphics.DrawString("Средња вредност\nдимности", Fontobicni, Brushes.Black, new Rectangle(xpos + 383, ypos + 2, 400, 30));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 30));
                e.Graphics.DrawString("Радна темпереатура\nмотора", Fontobicni, Brushes.Black, new Rectangle(xpos + 506, ypos + 2, 400, 30));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 30));
                e.Graphics.DrawString("Број обртаја мотора", Fontobicni, Brushes.Black, new Rectangle(xpos + 630, ypos + 4, 400, 30));
                ypos += 30;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 141, ypos, 80, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 221, ypos, 80, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 301, ypos, 80, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 381, ypos, 123, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 504, ypos, 124, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 628, ypos, 124, 20));

                //sesta tabela
                ypos += 30;
                //e.Graphics.DrawString("UTVRĐENE TEHNIČKE NEISPRAVNOSTI", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 400, 30));
                //ypos += 22;
                //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 780, 330));

                e.Graphics.DrawString("УТВРЂЕНЕ ТЕХНИЧКЕ НЕИСПРАВНОСТИ", Fontobicni2, Brushes.Black, new Rectangle(xpos, ypos, 794, 20));
                ypos += 18;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 210, 30));
                e.Graphics.DrawString("НАЗИВ ДЕЛА УРЕЂАЈА", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 200, 30), stringFormat);
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 210, ypos, 392, 30));
                e.Graphics.DrawString("НЕИСПРАВНОСТ/КОД", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 200, ypos, 420, 30), stringFormat);
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 602, ypos, 150, 30));
                e.Graphics.DrawString("СТЕПЕН", Fontobicni, Brushes.Black, new Rectangle(xpos + 602, ypos, 130, 30), stringFormat);
                ypos += 30;
                int tempYpos = ypos;
                for (int k = 0; k < 10; k++)
                {
                    e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 210, 20));
                    e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 210, ypos, 392, 20));
                    e.Graphics.FillRectangle(new SolidBrush(Color.LightGray), new Rectangle(xpos + 210, ypos, 392, 20));
                    e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 602, ypos, 150, 20));
                    e.Graphics.FillRectangle(new SolidBrush(Color.LightGray), new Rectangle(xpos + 602, ypos, 150, 20));
                    ypos += 20;
                }
                while (i < brNeispravnost)
                {
                    try
                    {
                        e.Graphics.DrawString(dataGridView1.Rows[i].Cells[1].FormattedValue.ToString().Length < 36 ? dataGridView1.Rows[i].Cells[1].FormattedValue.ToString() : dataGridView1.Rows[i].Cells[1].FormattedValue.ToString().Substring(0, 33) + "...", Fontobicni, Brushes.Black, new Rectangle(xpos, tempYpos, 200, 20), stringFormat2);
                        e.Graphics.DrawString(dataGridView1.Rows[i].Cells[2].FormattedValue.ToString().Length < 63 ? dataGridView1.Rows[i].Cells[2].FormattedValue.ToString() : dataGridView1.Rows[i].Cells[2].FormattedValue.ToString().Substring(0, 60) + "...", Fontobicni, Brushes.Black, new Rectangle(xpos + 212, tempYpos, 500, 20), stringFormat2);
                        e.Graphics.DrawString(dataGridView1.Rows[i].Cells[0].FormattedValue.ToString(), Fontobicni, Brushes.Black, new Rectangle(xpos + 550, tempYpos, 500, 20), stringFormat2);
                        e.Graphics.DrawString(dataGridView1.Rows[i].Cells[3].FormattedValue.ToString().Length < 63 ? dataGridView1.Rows[i].Cells[3].FormattedValue.ToString() : dataGridView1.Rows[i].Cells[3].FormattedValue.ToString().Substring(0, 60) + "...", Fontobicni, Brushes.Black, new Rectangle(xpos + 604, tempYpos, 140, 20), stringFormat2);
                    }
                    catch { }
                    tempYpos += 20;
                    i++;
                    if (i % 10 == 0)
                    {
                        break;
                    }
                }
                ypos += 10;
                e.Graphics.DrawString("НАПОМЕНА", Fontobicni2, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 730, 60));
                ypos += 18;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 752, 70));
                e.Graphics.DrawString(txtTPNap.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 730, 60));
                strana = 1;
                e.HasMorePages = false;
                return;
            }
        }

        //kreiranje registracionog
        public void CreateReceipt2(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int height = 0;
            int xpos = 30, ypos = 30;
            int i = 0;
            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.5f);
            Font FontNaznacenimanji = new Font("Times New Roman", 11, FontStyle.Bold);
            Font FontNaznaceni = new Font("Times New Roman", 11, FontStyle.Bold);
            Font FontNajjaci = new Font("Times New Roman", 14, FontStyle.Bold | FontStyle.Underline);
            Font Fontoznaka = new Font("Times New Roman", 10);
            Font Fontobicni = new Font("Times New Roman", 10);
            Font Fontobicni2 = new Font("Times New Roman", 10);
            Font fontUpisivanje = new Font("Times New Roman", 10, FontStyle.Italic);
            Font za_licnu_kartu = new Font("Times New Roman", 8);
            Font FontSvoji = new Font("Times New Roman", 14);
            Font fontmali = new Font("Times New Roman", 8);

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;

            StringFormat stringFormat2 = new StringFormat();
            stringFormat2.Alignment = StringAlignment.Near;
            stringFormat2.LineAlignment = StringAlignment.Center;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;

            //stampanje svojih podataka
            //e.Graphics.DrawString("ALSAZE 2016 DOO, Sjenica", FontSvoji, Brushes.Black, new Rectangle(xpos,ypos, 400, 20));
            //e.Graphics.DrawString("1 8 9 0 4 6", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 794, 15), stringFormat);
            e.Graphics.DrawString("Obrazac 1", Fontoznaka, Brushes.Black, new Rectangle(xpos, ypos, 750, 30), stringFormatdesno);

            ypos += 40;
            e.Graphics.DrawString("REGISTRACIONI LIST ZA MOTORNO-PRIKLJUČNO VOZILO", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 794, 20), stringFormat);
            ypos += 65;
            //stampanje podataka o kupcu

            e.Graphics.DrawString("PODACI O VOZILU", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            ypos += 50;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(J)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Vrsta vozila", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbVrstaVoz.Text + " " + cbVPV.Text + ((cbTerensko.Checked) ? "G" : ""), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(D.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Marka", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbVMarka.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(D.2)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Tip/varijanta/verzija", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            if (checkBox1.Checked)
                e.Graphics.DrawString(txtVTip.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 420, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(D.3)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Komercijalna oznaka (model)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVModel.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(E)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Broj šasije", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVBrojSasije.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(F.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Najveća dozvoljena masa (kg)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVNajMasa.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(G)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Masa", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVTezina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(K)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj homologacije", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            if (checkBox1.Checked)
                e.Graphics.DrawString(txtVHomologacija.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(P.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Radna zapremina motora", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVZapremina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(P.2)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Snaga motora u kW", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVSnaga.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(P.3)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Vrsta goriva ili pogona", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbVGorivo.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(Q)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Odnos snaga/masa u kw/kg (samo za motocikle)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVOdnos.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(S.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj mesta za sedenje uključujući i mesto vozača", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVSedenje1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(S.2)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj mesta za stajanje", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVStajanje.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(P.5)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj motora", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVBrojMotora.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(L)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVBrojOsovina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(R)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Boja vozila", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbBojaL.Text.Substring(0, 1) + " " + cbVBoja.Text.ToString() + " " + cbBojaD.Text.Substring(0, 1), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(B.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Godina proizvodnje vozila", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVGodinaProizvodnje.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(J.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Oblik karoserije", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            //string s="";
            //if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
            //{ 
            //    if(!checkBox12.Checked)
            //        s = cbVKlasa.Text + "/" + cbVOKS.Text;
            //    else
            //        s = cbVKlasa.Text + "/" + (cbVSpec.Text.Split(' '))[0].ToString();
            //}

            //else
            //     if (checkBox12.Checked)
            //{
            //    if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
            //        s = cbVKlasa.Text + "/" + (cbVSpec.Text.Split(' '))[0].ToString();
            //    else
            //        if (cbDecOzn.Text != "")
            //        s = (cbVSpec.Text.Split(' '))[0].ToString() + (cbDecOzn.Text.Split(' '))[0].ToString();
            //    else
            //        s = (cbVSpec.Text.Split(' '))[0].ToString();
            //    txtVrstaPlusDecOznk.Text = s;
            //}
            //else
            //    if (cbDecOzn.Text != "")
            //    s = cbVOKS.Text + (cbDecOzn.Text.Split(' '))[0].ToString();

            //else
            //    s = cbVOKS.Text;
            e.Graphics.DrawString(txtVrstaPlusDecOznk.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(G.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Nosivost", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVNosivost.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 30;

            e.Graphics.DrawString("PODACI O TEHNIČKOM PREGLEDU", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            ypos += 30;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Naziv i sedište ovlašćenog pravnog lica :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            e.Graphics.DrawString("VAAS d.o.o. Sjenica", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 20;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Šifra :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            e.Graphics.DrawString("141022", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 20;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("ID broj :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);

            e.Graphics.DrawString((int.Parse(txtTPID.Text.Substring(9))).ToString(), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 30;

            e.Graphics.DrawString("VOZILO JE TEHNIČKI ISPRAVNO", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            ypos += 30;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Datum vršenja tehničkog pregleda :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            e.Graphics.DrawString(txtTPDatum.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 30;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Kontrolor :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            e.Graphics.DrawString(cbTPKontr1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 50;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Napomena :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            e.Graphics.DrawString(txtTPNap.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 350, 100));
            ypos += 120;
            e.Graphics.DrawString("MP", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);


        }

        public void CreateReceipt21(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int height = 0;
            int xpos = 30, ypos = 30;
            int i = 0;
            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.5f);
            Font FontNaznacenimanji = new Font("Times New Roman", 11, FontStyle.Bold);
            Font FontNaznaceni = new Font("Times New Roman", 11, FontStyle.Bold);
            Font FontNajjaci = new Font("Times New Roman", 14, FontStyle.Bold | FontStyle.Underline);
            Font Fontoznaka = new Font("Times New Roman", 10);
            Font Fontobicni = new Font("Times New Roman", 10);
            Font Fontobicni2 = new Font("Times New Roman", 10);
            Font fontUpisivanje = new Font("Times New Roman", 10, FontStyle.Italic);
            Font za_licnu_kartu = new Font("Times New Roman", 8);
            Font FontSvoji = new Font("Times New Roman", 14);
            Font fontmali = new Font("Times New Roman", 8);

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;

            StringFormat stringFormat2 = new StringFormat();
            stringFormat2.Alignment = StringAlignment.Near;
            stringFormat2.LineAlignment = StringAlignment.Center;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;

            //stampanje svojih podataka
            //e.Graphics.DrawString("ALSAZE 2016 DOO, Sjenica", FontSvoji, Brushes.Black, new Rectangle(xpos,ypos, 400, 20));
            //e.Graphics.DrawString("1 8 9 0 4 6", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 794, 15), stringFormat);
            e.Graphics.DrawString("Obrazac 1", Fontoznaka, Brushes.Black, new Rectangle(xpos, ypos, 750, 30), stringFormatdesno);

            ypos += 40;
            e.Graphics.DrawString("REGISTRACIONI LIST ZA MOTORNO-PRIKLJUČNO VOZILO", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 794, 20), stringFormat);
            ypos += 65;
            //stampanje podataka o kupcu

            e.Graphics.DrawString("PODACI O VOZILU", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            ypos += 50;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(J)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Vrsta vozila", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbVrstaVoz.Text + " " + cbVPV.Text + ((cbTerensko.Checked) ? "G" : ""), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(D.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Marka", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbVMarka.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(D.2)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Tip/varijanta/verzija", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVTip.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 420, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(D.3)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Komercijalna oznaka (model)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVModel.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(E)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Broj šasije", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVBrojSasije.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(F.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Najveća dozvoljena masa (kg)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVNajMasa.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(G)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Masa", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVTezina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(K)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj homologacije", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVHomologacija.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(P.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Radna zapremina motora", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVZapremina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(P.2)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Snaga motora u kW", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVSnaga.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(P.3)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Vrsta goriva ili pogona", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbVGorivo.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(Q)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Odnos snaga/masa u kw/kg (samo za motocikle)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVOdnos.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(S.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj mesta za sedenje uključujući i mesto vozača", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVSedenje1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(S.2)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj mesta za stajanje", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVStajanje.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(P.5)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj motora", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVBrojMotora.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(L)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVBrojOsovina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(R)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Boja vozila", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbBojaL.Text.Substring(0, 1) + " " + cbVBoja.Text.ToString() + " " + cbBojaD.Text.Substring(0, 1), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(B.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Godina proizvodnje vozila", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVGodinaProizvodnje.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(J.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Oblik karoserije", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            //string s="";
            //if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
            //{ 
            //    if(!checkBox12.Checked)
            //        s = cbVKlasa.Text + "/" + cbVOKS.Text;
            //    else
            //        s = cbVKlasa.Text + "/" + (cbVSpec.Text.Split(' '))[0].ToString();
            //}

            //else
            //     if (checkBox12.Checked)
            //{
            //    if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
            //        s = cbVKlasa.Text + "/" + (cbVSpec.Text.Split(' '))[0].ToString();
            //    else
            //        if (cbDecOzn.Text != "")
            //        s = (cbVSpec.Text.Split(' '))[0].ToString() + (cbDecOzn.Text.Split(' '))[0].ToString();
            //    else
            //        s = (cbVSpec.Text.Split(' '))[0].ToString();
            //    txtVrstaPlusDecOznk.Text = s;
            //}
            //else
            //    if (cbDecOzn.Text != "")
            //    s = cbVOKS.Text + (cbDecOzn.Text.Split(' '))[0].ToString();

            //else
            //    s = cbVOKS.Text;
            e.Graphics.DrawString(txtVrstaPlusDecOznk.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(G.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Nosivost", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVNosivost.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 30;

            e.Graphics.DrawString("PODACI O TEHNIČKOM PREGLEDU", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            ypos += 30;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Naziv i sedište ovlašćenog pravnog lica :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            e.Graphics.DrawString("VAAS d.o.o. Sjenica", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 20;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Šifra :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawString("141022", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 20;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("ID broj :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);

            //e.Graphics.DrawString((int.Parse(txtTPID.Text.Substring(9))).ToString(), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 30;

            e.Graphics.DrawString("VOZILO JE TEHNIČKI ISPRAVNO", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            ypos += 30;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Datum vršenja tehničkog pregleda :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawString(txtTPDatum.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 30;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Kontrolor :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawString(cbTPKontr1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 50;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Napomena :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            e.Graphics.DrawString("ODJAVA VOZILA", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 350, 100));
            ypos += 120;
            e.Graphics.DrawString("MP", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);


        }


        //ocitavanje podataka sa saobracajne
        private void bunifuThinButton29_Click(object sender, EventArgs e)
        {
            List<ReaderDescriptor> readers = new List<ReaderDescriptor>();
            try
            {
                readers = SaobracajnaReader.GetAllReaders();
                ReaderDescriptor descriptor = readers[0];

                var data = SaobracajnaReader.ReadAll(descriptor);
                PrintData(data);
                string[] podaci = cbVMesto.Text.Split(',');

                //txtFLOps.Text = podaci[1];
                cbVMesto.Text = podaci[0] + ", " + podaci[1];
                txtVUlica.Text = podaci[2];
                txtVBroj.Text = podaci[3];
                richTextBox1.Text = txtVIme.Text + " " + txtVPrezime.Text + Environment.NewLine + podaci[2] + " " + podaci[3] + " " + podaci[1];
                popuniPrethodneNeispravnosti();
                popuniFakture();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Molim Vas ubacite saobraćajnu dozvolu");
            }

            //var readers = SaobracajnaReader.GetAllReaders();
            //if (readers.Count == 0)
            //{
            //    MessageBox.Show("Nema povezanih citaca.");

            //    return;
            //}

            //ReaderDescriptor descriptor = readers[0];
            //try
            //{
            //    var data = SaobracajnaReader.ReadAll(descriptor);
            //    PrintData(data);
            //    string[] podaci = cbFLMesto.Text.Split(',');

            //    txtFLOps.Text = podaci[1];
            //    cbFLMesto.Text = podaci[0];
            //    txtFLUlica.Text = podaci[2];
            //    txtFLBroj.Text = podaci[3];
            //    richTextBox1.Text = txtFLIme.Text + " " + txtFLPrez.Text + Environment.NewLine + podaci[2] + " " + podaci[3] + " " + podaci[1];

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

        }

        private void popuniPrethodneNeispravnosti()
        {
            vozilo v = b.reg(txtVBrojSasije.Text.Trim());
            //if (v.Ispravnost == "Ne" && v.Datum.Substring(6, 4) == DateTime.Now.Year.ToString())
            // DateTime dolazio = Convert.ToDateTime(v.Datum);
            if (v.Ispravnost == "Ne" && v.Datum.Substring(6, 4) == DateTime.Now.Year.ToString())
            {
                txtTPNap.Text = v.Datum + "\r\n" + v.IDBroj;
                txtVezaID.Text = v.IDBroj;
                txtVezaGod.Text = v.Datum;
                //rbTehNeis.Checked = true;
                //rbTehIsp.Checked = false;
                popuniTabeluNeispravnosti(txtVezaID.Text);
            }
        }

        //ispisivanje ocitanih podataka sa saobracajne
        private void PrintData(AllData data)
        {
            PrintStruct(data.VehicleData);
            PrintStruct(data.DocumentData);
            PrintStruct(data.PersonalData);
            //MessageBox.Show(data.RegistrationData.Count + " uspešno čitanje podataka sa saobraćajne dozvole");
            //txtTPImeiPrezime.Text = txtVIme.Text + " " + txtVPrezime.Text;
            //btnIzv.Visible = true;
            //btnPot2.Visible = true;
            //btnReg2.Visible = true;
            //btnZap1.Visible = true;
        }

        //ispisivanje ocitanih podataka sa saobracajne
        private void PrintStruct(object instance)
        {
            var type = instance.GetType();
            //int i = 0;
            var fields = type.GetFields();
            foreach (var field in fields)
            {
                var value = field.GetValue(instance);
                switch (field.Name)
                {
                    case "IssuingDate":
                        tbDatumIzdavanja.Text = value.ToString();                  //lblDatumIzdavanjaSaobracajne.Text = value.ToString();
                        break;
                    case "DateOfFirstRegistration":
                        txtVPrvaReg.Text = value.ToString();
                        break;
                    case "YearOfProduction":
                        txtVGodinaProizvodnje.Text = value.ToString();
                        break;
                    case "VehicleMake":
                        cbVMarka.Text = value.ToString();
                        break;
                    case "CommercialDescription":
                        txtVModel.Text = value.ToString();
                        break;
                    case "VehicleLoad":
                        txtVNosivost.Text = value.ToString();
                        break;
                    case "VehicleIDNumber":
                        txtVBrojSasije.Text = value.ToString();
                        break;
                    case "RegistrationNumberOfVehicle":
                        txtRegBr.Text = value.ToString();
                        break;
                    case "MaximumNetPower":
                        txtVSnaga.Text = value.ToString();
                        break;
                    case "EngineCapacity":
                        txtVZapremina.Text = value.ToString();
                        break;
                    case "TypeOfFuel":
                        if (value.ToString() == "BENZIN 98" || value.ToString() == "BENZIN 86")
                            cbVGorivo.Text = "BENZIN";
                        else if (value.ToString() == "BENZIN-GAS")
                            cbVGorivo.Text = "BENZIN/TNG";
                        else
                            cbVGorivo.Text = value.ToString();
                        break;
                    case "VehicleMass":
                        txtVTezina.Text = value.ToString();
                        break;
                    case "MaximumPermissibleLadenMass":
                        txtVNajMasa.Text = value.ToString();
                        break;
                    case "NumberOfSeats":
                        txtVSedenje1.Text = value.ToString();
                        break;
                    case "NumberOfStandingPlaces":
                        txtVStajanje.Text = value.ToString();
                        break;
                    case "EngineIDNumber":
                        txtVBrojMotora.Text = value.ToString();
                        break;
                    case "NumberOfAxles":
                        txtVBrojOsovina.Text = value.ToString();
                        break;
                    case "VehicleCategory":
                        cbVrstaVoz.Text = value.ToString();
                        break;
                    case "ColourOfVehicle":
                        if (value.ToString().Split(' ').Length > 2)
                        {
                            cbBojaL.Text = value.ToString().Split(' ')[0];
                            cbVBoja.Text = value.ToString().Split(' ')[1];
                            cbBojaD.Text = value.ToString().Split(' ')[2];
                        }
                        else
                            cbVBoja.Text = value.ToString();
                        break;
                    case "OwnerName":
                        txtVIme.Text = value.ToString();
                        break;
                    case "OwnersSurnameOrBusinessName":
                        //if (checkBox4.Checked)                        
                        txtVPrezime.Text = value.ToString();
                        break;
                    case "OwnerAddress":
                        cbVMesto.Text = value.ToString();
                        break;
                    case "OwnersPersonalNo":
                        //case "ownersPersonalNo":
                        txtVJMBG.Text = value.ToString();
                        break;
                    case "UnambiguousNumber":
                        txtVBrDokumenta.Text = value.ToString();
                        break;
                    case "SerialNumber":
                        txtVSerijskiBroj.Text = value.ToString().ToUpper();
                        break;
                    case "ExpiryDate":
                        //lblDatumIstekaSaobracajne.Text = value.ToString();
                        //var d2 = lblDatumIstekaSaobracajne.Text.Split('.');
                        //var d = new DateTime(int.Parse(d2[2].ToString()), int.Parse(d2[1].ToString()), int.Parse(d2[0].ToString()));
                        //if (DateTime.Now < d)
                        //{
                        //    lblDatumIstekaSaobracajne.ForeColor = Color.SeaGreen;

                        //}
                        //else lblDatumIstekaSaobracajne.ForeColor = Color.Red;

                        break;

                    default:
                        // MessageBox.Show(type.Name + " " + field.Name + ", " + value);
                        break;

                }

                //Console.WriteLine("{0}: {1}", type.Name + "." + field.Name, value);
                //i++;
            }

        }

        private void bunifuTileButton2_Click(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();

            PrintDocument printDocument = new PrintDocument();

            printDialog.Document = printDocument; //add the document to the dialog box...        

            printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceip5); //add an event handler that will do the printing

            //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

            DialogResult result = printDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                printDocument.Print();

            }
        }

        private void CreateReceip5(object sender, PrintPageEventArgs e)
        {
            int height = 0;
            int xpos = 30, ypos = 40;
            int i = 0;
            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.5f);
            Font FontNaznacenimanji = new Font("Times New Roman", 11, FontStyle.Bold);
            Font FontNaznaceni = new Font("Times New Roman", 12, FontStyle.Bold);
            Font FontNajjaci = new Font("Times New Roman", 14, FontStyle.Bold | FontStyle.Underline);
            Font Fontoznaka = new Font("Times New Roman", 10);
            Font Fontobicni = new Font("Times New Roman", 9);
            Font Fontobicni2 = new Font("Times New Roman", 10);
            Font fontUpisivanje = new Font("Times New Roman", 10, FontStyle.Italic);
            Font za_licnu_kartu = new Font("Times New Roman", 8);
            Font FontSvoji = new Font("Times New Roman", 14);
            Font fontmali = new Font("Times New Roman", 6);

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;

            StringFormat stringFormat2 = new StringFormat();
            stringFormat2.Alignment = StringAlignment.Near;
            stringFormat2.LineAlignment = StringAlignment.Center;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;


            //stampanje svojih podataka
            e.Graphics.DrawString("ALSAZE 2016 DOO, Sjenica", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 400, 20));
            e.Graphics.DrawString("1 8 9 0 4 6", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 794, 20), stringFormat);
            e.Graphics.DrawString("Obrazac 3\nstrana 2", Fontoznaka, Brushes.Black, new Rectangle(xpos, ypos, 780, 30), stringFormatdesno);

            ypos += 40;
            e.Graphics.DrawString("PODACI O IZVRŠENIM MERENJIMA POMOĆU UREĐAJA I PRORAČUNIMA", FontNajjaci, Brushes.Black, new Rectangle(xpos, ypos, 794, 20), stringFormat);
            ypos += 30;
            int pom = ypos;
            //prva tabela
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 130, 20));
            e.Graphics.DrawString("1. osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 132, ypos, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 130, 20));
            e.Graphics.DrawString("2. osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 262, ypos, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 20));
            e.Graphics.DrawString("3. osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 392, ypos, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 130, 20));
            e.Graphics.DrawString("4. osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 520, ypos, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos, 130, 20));
            e.Graphics.DrawString("5. osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 650, ypos, 400, 20));
            ypos += 20;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 130, 40));
            e.Graphics.DrawString("Sile radnog kočenja", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 10, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 110, ypos, 20, 20));
            e.Graphics.DrawString("L", Fontobicni, Brushes.Black, new Rectangle(xpos + 112, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 110, ypos + 20, 20, 20));
            e.Graphics.DrawString("D", Fontobicni, Brushes.Black, new Rectangle(xpos + 112, ypos + 22, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos + 20, 130, 20));
            ypos += 40;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 130, 40));
            e.Graphics.DrawString("Ovalnost kočnica", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 10, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 110, ypos, 20, 20));
            e.Graphics.DrawString("L", Fontobicni, Brushes.Black, new Rectangle(xpos + 112, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 110, ypos + 20, 20, 20));
            e.Graphics.DrawString("D", Fontobicni, Brushes.Black, new Rectangle(xpos + 112, ypos + 22, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos + 20, 130, 20));
            ypos += 40;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 130, 20));
            e.Graphics.DrawString("Razlike sila kočenja", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos, 130, 20));
            ypos += 20;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 130, 20));
            e.Graphics.DrawString("Sile aktiviranja", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos, 130, 20));
            ypos += 20;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 130, 20));
            e.Graphics.DrawString("Pritisak vazduha", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos, 130, 20));
            ypos += 20;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 130, 40));
            e.Graphics.DrawString("Sile parkirnog", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 5, 400, 20));
            e.Graphics.DrawString("kočenja", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 18, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 110, ypos, 20, 20));
            e.Graphics.DrawString("L", Fontobicni, Brushes.Black, new Rectangle(xpos + 112, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 110, ypos + 20, 20, 20));
            e.Graphics.DrawString("D", Fontobicni, Brushes.Black, new Rectangle(xpos + 112, ypos + 22, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos + 20, 130, 20));
            ypos += 40;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 130, 40));
            e.Graphics.DrawString("Sile pomoćnog", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 5, 400, 20));
            e.Graphics.DrawString("kočenja", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 18, 400, 20));
            e.Graphics.DrawString("3", fontmali, Brushes.Black, new Rectangle(xpos + 42, ypos + 18, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 110, ypos, 20, 20));
            e.Graphics.DrawString("L", Fontobicni, Brushes.Black, new Rectangle(xpos + 112, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 110, ypos + 20, 20, 20));
            e.Graphics.DrawString("D", Fontobicni, Brushes.Black, new Rectangle(xpos + 112, ypos + 22, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos + 20, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos + 20, 130, 20));
            ypos += 40;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 130, 40));
            e.Graphics.DrawString("Kočni koeficijent", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 5, 400, 20));
            e.Graphics.DrawString("radnog kočenja", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 18, 400, 20));
            e.Graphics.DrawString("1", fontmali, Brushes.Black, new Rectangle(xpos + 80, ypos + 18, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 260, 40));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 40));
            e.Graphics.DrawString("Kočni koeficijent", Fontobicni, Brushes.Black, new Rectangle(xpos + 392, ypos + 5, 400, 20));
            e.Graphics.DrawString("parkirnog kočenja", Fontobicni, Brushes.Black, new Rectangle(xpos + 392, ypos + 18, 400, 20));
            e.Graphics.DrawString("2", fontmali, Brushes.Black, new Rectangle(xpos + 484, ypos + 18, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 260, 40));
            ypos += 40;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 130, 40));
            e.Graphics.DrawString("Vlažnost kočne", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 5, 400, 20));
            e.Graphics.DrawString("tečnosti (%)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 18, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 260, 40));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 40));
            e.Graphics.DrawString("Brzina na početnu", Fontobicni, Brushes.Black, new Rectangle(xpos + 392, ypos + 3, 400, 20));
            e.Graphics.DrawString("merenja usporenja", Fontobicni, Brushes.Black, new Rectangle(xpos + 392, ypos + 14, 400, 20));
            e.Graphics.DrawString("na poligonu", Fontobicni, Brushes.Black, new Rectangle(xpos + 392, ypos + 24, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 260, 40));

            //fusnota
            ypos += 40;
            e.Graphics.DrawString("1- upisuje   se   vrednost   koeficijenta,   osim   u   slučaju   kontrole   priključnog   vozila   na   poligon   kada   se   upisuje   DA   kada   je   sistem   ispravan   odnosno   NE   kada   sistem   nije   ispravan", fontmali, Brushes.Black, new Rectangle(xpos + 2, ypos + 5, 780, 20));
            e.Graphics.DrawString("2- upisuje   se   vrednost   koeficijenta,   osim   u   slučaju   kontrole   na   poligonu   kada   se   upisuje   DA   kada   je   sistem   ispravan   odnosno   NE   kada   sistem   nije", fontmali, Brushes.Black, new Rectangle(xpos + 2, ypos + 15, 780, 20));
            e.Graphics.DrawString("3- utvrđuje   se   kada   postoji   poseban   sistem   pomoćnog   kočenja   ili   kada   su   pomoćno   i   parkirano   kočenje   izvedeni   kao   jedan   sistem", fontmali, Brushes.Black, new Rectangle(xpos + 2, ypos + 25, 780, 20));

            //druga tabela
            ypos += 45;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 130, 20));
            e.Graphics.DrawString("1. osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 132, ypos, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 130, 20));
            e.Graphics.DrawString("2. osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 262, ypos, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 20));
            e.Graphics.DrawString("3. osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 392, ypos, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 130, 20));
            e.Graphics.DrawString("4. osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 520, ypos, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos, 130, 20));
            e.Graphics.DrawString("5. osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 650, ypos, 400, 20));
            ypos += 20;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 130, 20));
            e.Graphics.DrawString("Osovinsko opterećenje", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 130, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 390, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 520, ypos, 130, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 650, ypos, 130, 20));

            //treca tabela
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 260, 30));
            e.Graphics.DrawString("Brzina mopeda i vozila koja se proizvode \nu vrarijanti mopeda i motorcikla", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 1, 400, 30));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 260, ypos, 520, 30));

            //cetvrta tabela
            ypos += 40;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 180, 80));
            e.Graphics.DrawString("Emisija izduvnih gasova motora \nsa aktivnim paljenjem", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 25, 400, 30));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 180, ypos, 150, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 330, ypos, 150, 20));
            e.Graphics.DrawString("Prazan hod", Fontobicni, Brushes.Black, new Rectangle(xpos + 332, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 480, ypos, 150, 20));
            e.Graphics.DrawString("n≥2000 obrtaja", Fontobicni, Brushes.Black, new Rectangle(xpos + 482, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 630, ypos, 150, 20));
            e.Graphics.DrawString("Radna temperatura motora", Fontobicni, Brushes.Black, new Rectangle(xpos + 632, ypos + 2, 400, 20));
            ypos += 20;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 180, ypos, 150, 20));
            e.Graphics.DrawString("CO", Fontobicni, Brushes.Black, new Rectangle(xpos + 182, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 330, ypos, 150, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 480, ypos, 150, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 630, ypos, 150, 60));
            ypos += 20;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 180, ypos, 150, 20));
            e.Graphics.DrawString("ƛ faktor", Fontobicni, Brushes.Black, new Rectangle(xpos + 182, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 330, ypos, 150, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 480, ypos, 150, 20));
            ypos += 20;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 180, ypos, 150, 20));
            e.Graphics.DrawString("Broj obrtaja motora", Fontobicni, Brushes.Black, new Rectangle(xpos + 182, ypos + 2, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 330, ypos, 150, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 480, ypos, 150, 20));

            //peta tabela
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 141, 50));
            e.Graphics.DrawString("Dimnost izduvnih gasova\ndizel motora", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 10, 400, 30));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 141, ypos, 85, 30));
            e.Graphics.DrawString("1. merenje", Fontobicni, Brushes.Black, new Rectangle(xpos + 143, ypos + 4, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 226, ypos, 85, 30));
            e.Graphics.DrawString("2. merenje", Fontobicni, Brushes.Black, new Rectangle(xpos + 228, ypos + 4, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 311, ypos, 85, 30));
            e.Graphics.DrawString("3. merenje", Fontobicni, Brushes.Black, new Rectangle(xpos + 313, ypos + 4, 400, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 396, ypos, 128, 30));
            e.Graphics.DrawString("Srednja vrednost\ndimnosti", Fontobicni, Brushes.Black, new Rectangle(xpos + 398, ypos + 2, 400, 30));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 524, ypos, 128, 30));
            e.Graphics.DrawString("Radna temperatura\nmotora", Fontobicni, Brushes.Black, new Rectangle(xpos + 526, ypos + 2, 400, 30));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 652, ypos, 128, 30));
            e.Graphics.DrawString("Broj obrtaja motora", Fontobicni, Brushes.Black, new Rectangle(xpos + 654, ypos + 4, 400, 30));
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 141, ypos, 85, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 226, ypos, 85, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 311, ypos, 85, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 396, ypos, 128, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 524, ypos, 128, 20));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 652, ypos, 128, 20));

            //sesta tabela
            ypos += 30;
            //e.Graphics.DrawString("UTVRĐENE TEHNIČKE NEISPRAVNOSTI", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 400, 30));
            //ypos += 22;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 780, 330));

            e.Graphics.DrawString("UTVRĐENE NEISPRAVNOSTI", FontNajjaci, Brushes.Black, new Rectangle(xpos, ypos, 794, 20));
            ypos += 22;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 210, 30));
            e.Graphics.DrawString("Naziv dela uređaja", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 200, 30), stringFormat);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 210, ypos, 420, 30));
            e.Graphics.DrawString("Neispravnost/kod", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 200, ypos, 420, 30), stringFormat);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 630, ypos, 150, 15));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 630, ypos + 15, 50, 15));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 680, ypos + 15, 50, 15));
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 730, ypos + 15, 50, 15));
            e.Graphics.DrawString("Stepen neispravnosti", Fontobicni, Brushes.Black, new Rectangle(xpos + 630, ypos, 130, 15), stringFormat);
            e.Graphics.DrawString("Male", Fontoznaka, Brushes.Black, new Rectangle(xpos + 630, ypos + 15, 50, 15), stringFormat);
            e.Graphics.DrawString("Velike", Fontoznaka, Brushes.Black, new Rectangle(xpos + 680, ypos + 15, 50, 15), stringFormat);
            e.Graphics.DrawString("Opasne", Fontoznaka, Brushes.Black, new Rectangle(xpos + 730, ypos + 15, 50, 15), stringFormat);
            for (i = 0; i < 7; i++)
            {
                ypos += 30;
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 210, 30));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 210, ypos, 420, 30));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 630, ypos, 50, 30));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 680, ypos, 50, 30));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 730, ypos, 50, 30));
                //try
                //{
                //    e.Graphics.DrawString(dataGridView1.Rows[i].Cells[0].FormattedValue.ToString(), Fontobicni, Brushes.Black, new Rectangle(xpos, ypos, 200, 40), stringFormat2);
                //    e.Graphics.DrawString(dataGridView1.Rows[i].Cells[1].FormattedValue.ToString(), Fontobicni, Brushes.Black, new Rectangle(xpos + 200, ypos, 420, 40), stringFormat2);
                //    e.Graphics.DrawString(dataGridView1.Rows[i].Cells[2].FormattedValue.ToString(), Fontobicni, Brushes.Black, new Rectangle(xpos + 620, ypos, 50, 40), stringFormat2);
                //    e.Graphics.DrawString(dataGridView1.Rows[i].Cells[3].FormattedValue.ToString(), Fontobicni, Brushes.Black, new Rectangle(xpos + 670, ypos, 50, 40), stringFormat2);
                //    e.Graphics.DrawString(dataGridView1.Rows[i].Cells[4].FormattedValue.ToString(), Fontobicni, Brushes.Black, new Rectangle(xpos + 720, ypos, 50, 40), stringFormat2);
                //}
                //catch { }
            }
            ypos += 35;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 780, 60));
            e.Graphics.DrawString("Napomena:", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 400, 20));

        }


        //stampanje registracionog
        private void bunifuThinButton211_Click(object sender, EventArgs e)
        {
            if (txtVKM.Text != "" && cbVrstaVoz.Text != "" && (checkBox12.Checked && cbVOblikKar.Text == "" ? true : (cbVOblikKar.Text == "" ? false : true)) && cbVMarka.Text != "" && txtVModel.Text != "" && txtVGodinaProizvodnje.Text != ""
                && txtVBrojSasije.Text != "" && txtVBrojMotora.Text != "" && cbVBoja.Text != "" && cbBojaL.Text != "" && cbBojaD.Text != "" && txtVPrvaReg.Text != ""
                && txtVSnaga.Text != "" && txtVZapremina.Text != "" && txtVNajMasa.Text != "" && txtVTezina.Text != "" && txtVSedenje1.Text != "" && txtVStajanje.Text != ""
                && txtVBrojOsovina.Text != "" && cbVGorivo.Text != "" && txtTPID.Text != "" && ispravnost != "" && cbVPV.Text != "" && (checkBox12.Checked && cbVOKS.Text == "" ? true : (cbVOKS.Text == "" ? false : true))
                && txtTPVremeOd.Text != "  :" && txtTPVremeDo.Text != "  :")
            {
                int n;
                if (txtTPID.Text.Length < 14 || txtTPID.Text.Length > 14)
                {
                    MessageBox.Show("ID Broj je nepravilan!");
                }
                else if (int.TryParse(txtTPID.Text.Substring(9), out n) == false)
                {
                    MessageBox.Show("ID Broj je nepravilan!");
                }

                else if (rbTehNeis.Checked == true)
                {
                    MessageBox.Show("Tehnička ispranost vozila je obeležena sa \'Ne\'!");
                }

                else
                {
                    lblVrstaVoz.ForeColor = SystemColors.ControlText; lblVOblikKar.ForeColor = SystemColors.ControlText; lblVMarka.ForeColor = SystemColors.ControlText;
                    lblVModel.ForeColor = SystemColors.ControlText; lblVGodinaProizvodnje.ForeColor = SystemColors.ControlText; lblVBrojSasije.ForeColor = SystemColors.ControlText;
                    lblVBrojMotora.ForeColor = SystemColors.ControlText; lblVBoja.ForeColor = SystemColors.ControlText; lblVPrvaReg.ForeColor = SystemColors.ControlText;
                    lblVSnaga.ForeColor = SystemColors.ControlText; lblVZapremina.ForeColor = SystemColors.ControlText; lblVNajMasa.ForeColor = SystemColors.ControlText;
                    lblVTezina.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText;
                    lblVBrojOsovina.ForeColor = SystemColors.ControlText; lblVGorivo.ForeColor = SystemColors.ControlText; lblTPID.ForeColor = SystemColors.ControlText;
                    lblIspravnost.ForeColor = SystemColors.ControlText; lblPredjenikm.ForeColor = SystemColors.ControlText; lblVremeOd.ForeColor = SystemColors.ControlText;
                    lblVremeDo.ForeColor = SystemColors.ControlText;

                    PrintDialog printDialog = new PrintDialog();

                    PrintDocument printDocument = new PrintDocument();

                    printDialog.Document = printDocument; //add the document to the dialog box...        

                    printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt2); //add an event handler that will do the printing

                    //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

                    DialogResult result = printDialog.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        printDocument.Print();

                    }
                }
            }
            else
            {
                MessageBox.Show("Popunite sva obavezna polja na tabovima \'Podaci o vozilu\' i \'Tehnički pregled\'.");

                lblVrstaVoz.ForeColor = Color.Red; lblVOblikKar.ForeColor = Color.Red; lblVMarka.ForeColor = Color.Red; lblVModel.ForeColor = Color.Red;
                lblVGodinaProizvodnje.ForeColor = Color.Red; lblVBrojSasije.ForeColor = Color.Red; lblVBrojMotora.ForeColor = Color.Red; lblVBoja.ForeColor = Color.Red;
                lblVPrvaReg.ForeColor = Color.Red; lblVSnaga.ForeColor = Color.Red; lblVZapremina.ForeColor = Color.Red; lblVNajMasa.ForeColor = Color.Red;
                lblVTezina.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVBrojOsovina.ForeColor = Color.Red;
                lblVGorivo.ForeColor = Color.Red; lblTPID.ForeColor = Color.Red; lblPredjenikm.ForeColor = Color.Red; lblIspravnost.ForeColor = Color.Red;
                lblVremeOd.ForeColor = Color.Red; lblVremeDo.ForeColor = Color.Red;
            }

        }

        //postavljanje podataka nakon loada ucregistracije
        private void ucRegistracija_Load(object sender, EventArgs e)
        {
            //dodavanje boja iz baze
            List<string> boje = new List<string>();
            foreach (var b in b.lista_boja())
            {
                boje.Add(b.BojaNaziv);
            }
            cbVBoja.DataSource = boje;
            cbVBoja.SelectedIndex = -1;
            cbVBoja.Text = "";

            //MestoVrsenja
            cbTPMestoVrsenja.SelectedIndex = 0;

            ////dodavanje marki iz baze
            List<string> marke = new List<string>();
            foreach (var b in b.lista_marka())
            {
                marke.Add(b.KaroserijaNaziv);
            }
            cbVMarka.DataSource = null;
            cbVMarka.DataSource = marke;
            cbVMarka.SelectedIndex = -1;
            cbVMarka.Text = "";

            //dodavanje specijalnih iz baze
            List<string> specijalna = new List<string>();
            foreach (var s in b.lista_spec())
            {
                specijalna.Add(s.Opis);
            }
            cbVSpec.DataSource = specijalna;
            cbVSpec.Text = "";

            //dodavanje vrste vozila
            List<string> str = new List<string>();
            foreach (var i in b.vrvozila())
            {
                str.Add(i.NazivVrste);
            }
            cbVrstaVoz.DataSource = str;
            cbVrstaVoz.Text = "";
            vrstaVozilaPrvo = false;

            cbFakture.SelectedIndex = 0;
            //cbTPKontr1.DataSource = b.nadjiKotrolore();
            //cbTPKontr2.DataSource = b.nadjiKotrolore();

        }

        //snimanje podataka u bazu
        private void bunifuThinButton210_Click(object sender, EventArgs e)
        {
            if (rbVL.Checked == true)
            {
                if (checkBox2.Checked == true)
                {
                    if (txtVZiroRacun.Text == "")
                    {
                        label130.ForeColor = Color.Red;
                        MessageBox.Show("Ukoliko je lice pravno potreban je Žiro Račun");
                        return;
                    }
                    else
                    {
                        label130.ForeColor = SystemColors.ControlText; // reset boje ako nije prazno
                    }
                }

                if (txtVIme.Text == "" || txtVPrezime.Text == "" || txtVJMBG.Text == "" || txtVBRLK.Text == "" || txtVBrPoste.Text == "" || cbVMesto.Text == ""
                    || txtVUlica.Text == "" )
                {
                    label47.ForeColor = Color.Red;
                    label50.ForeColor = Color.Red;
                    label51.ForeColor = Color.Red;
                    label45.ForeColor = Color.Red;
                    label41.ForeColor = Color.Red;
                    label37.ForeColor = Color.Red;
                    MessageBox.Show("Popunite sve obavezna polja u tabu Podaci o Vlasniku Vozila");
                    return;

                }
                else
                {
                    label47.ForeColor = SystemColors.ControlText;
                    label50.ForeColor = SystemColors.ControlText;
                    label51.ForeColor = SystemColors.ControlText;
                    label45.ForeColor = SystemColors.ControlText;
                    label41.ForeColor = SystemColors.ControlText;
                    label37.ForeColor = SystemColors.ControlText;
                    label55.ForeColor = SystemColors.ControlText;

                }
            }

            if (rbDL.Checked)
            {
                if (cbDLPravno.Checked)
                {
                    if (txtDLZiroRacun.Text == "")
                    {
                        label56.ForeColor = Color.Red;
                        MessageBox.Show("Ukoliko je lice Korisnik potreban je Žiro Račun");
                        return;
                    }
                    else
                    {
                        label56.ForeColor = SystemColors.ControlText; // reset boje ako nije prazno
                    }
                }
                if (txtDLIme.Text == "" || txtDLPrezime.Text == "" || txtDLJMBG.Text == "" || txtDLBRLK.Text == "" || txtDLBrPoste.Text == "" || cbDLMesto.Text == ""
                    || txtDLUlica.Text == "" )
                {
                    label35.ForeColor = Color.Red;
                    label34.ForeColor = Color.Red;
                    label33.ForeColor = Color.Red;
                    label36.ForeColor = Color.Red;
                    label38.ForeColor = Color.Red;
                    label42.ForeColor = Color.Red;
                    MessageBox.Show("Popunite sve obavezna polja u tabu Podaci o Drugom Licu");
                    return;

                }
                else
                {
                    label35.ForeColor = SystemColors.ControlText;
                    label34.ForeColor = SystemColors.ControlText;
                    label33.ForeColor = SystemColors.ControlText;
                    label36.ForeColor = SystemColors.ControlText;
                    label38.ForeColor = SystemColors.ControlText;
                    label42.ForeColor = SystemColors.ControlText;
                    label43.ForeColor = SystemColors.ControlText;
                }
            }


            if (txtVKM.Text != "" && cbVrstaVoz.Text != "" && (checkBox12.Checked && cbVOblikKar.Text == "" ? true : (cbVOblikKar.Text == "" ? false : true)) && cbVMarka.Text != "" && txtVModel.Text != "" && txtVGodinaProizvodnje.Text != ""
                && txtVBrojSasije.Text != "" && txtVBrojMotora.Text != "" && cbVBoja.Text != "" && cbBojaL.Text != "" && cbBojaD.Text != "" && txtVPrvaReg.Text != ""
                && txtVSnaga.Text != "" && txtVZapremina.Text != "" && txtVNajMasa.Text != "" && txtVTezina.Text != "" && txtVSedenje1.Text != "" && txtVStajanje.Text != ""
                && txtVBrojOsovina.Text != "" && cbVGorivo.Text != "" && txtVPneumatikL1.Text != "" && txtTPID.Text != "" && ispravnost != "" && cbVPV.Text != "" && cbUredjajZaZaust.Text != "" && txtVVlaznostKT.Text != "" && (checkBox12.Checked && cbVOKS.Text == "" ? true : (cbVOKS.Text == "" ? false : true))
                && txtVNosivost.Text != ""
                && txtTPVremeOd.Text.Trim() != "" && txtTPVremeDo.Text.Trim() != "")
            {
                float koef = 0;
                try
                {
                    koef = float.Parse(txtVVlaznostKT.Text);
                }
                catch
                {
                    koef = 0;
                }
                if (rbTehIsp.Checked == true && koef >= 4)
                {
                    MessageBox.Show("Vozilo ne moze biti ispravno ako je vlaznost 4 ili veca!");
                }
                else if (rbTehNeis.Checked == true && dataGridView1.RowCount <= 1)
                {
                    MessageBox.Show("Tehnička ispravnost vozila je označena sa \'Ne\', unesite neispravnosti!");
                }
                else if (rbTehIsp.Checked == true && dataGridView1.RowCount > 1)
                {
                    MessageBox.Show("Tehnička ispravnost vozila je označena sa \'Da\', izbrišite neispravnosti!");
                }
                //else if(rbTehIspPrekid.Checked && !(comboBox2.SelectedIndex == 0 || comboBox2.SelectedIndex == 1 || comboBox2.SelectedIndex == 3))
                //{
                //    MessageBox.Show("Prekid se može izvršiti samo za REDOVNE, VANREDNE i ŠESTOMESEČNE preglede");
                //}
                else
                {
                    vozilo v = b.id(txtTPID.Text.Trim());
                    if (rbVL.Checked)
                    {
                        v.Prioritet = "VLASNIK";
                    }
                    else
                    {
                        v.Prioritet = "DRUGO LICE";
                    }
                    //{
                    //    txtDLBRLK.Text = "";
                    //    txtDLBroj.Text = "";
                    //    txtDLBrPoste.Text = "";
                    //    txtDLBrTelefona.Text = "";
                    //    txtDLIme.Text = "";
                    //    txtDLJMBG.Text = "";
                    //    txtDLPrezime.Text = "";
                    //    txtDLUlica.Text = "";
                    //    cbDLMesto.Text = "";
                    //}
                    if (v.IDBroj == null)
                    {
                        try
                        {
                            //    b.unesiKorisnika(txtTPID.Text, txtRegBr.Text, cbVrstaVoz.Text, cbVKlasa.Text, cbVOblikKar.Text, cbVSpec.Text, cbVZemlja.Text,
                            //cbVMarka.Text, txtVModel.Text, txtVTip.Text, txtVGodinaProizvodnje.Text, txtVBrojSasije.Text, txtVBrojMotora.Text, cbVNamena.Text, cbVBoja.Text,
                            //txtVPrvaReg.Text, txtVOdnos.Text, txtVSnaga.Text, txtVZapremina.Text, txtVNajMasa.Text, txtVTezina.Text, txtVNosivost.Text, txtVSedenje1.Text,
                            //txtVStajanje.Text, txtVBrojOsovina.Text, txtVBrPogOsovina.Text, cbVGorivo.Text, txtVHomologacija.Text, txtVJMBG.Text, txtVPrezime.Text,
                            //txtVIme.Text, txtVBRLK.Text, txtFLOps.Text, txtVBrPoste.Text, cbVMesto.Text, txtVUlica.Text, txtVBroj.Text, txtVBrTelefona.Text, txtFLEmail.Text,
                            //txtPLNaziv.Text, txtPLMatBr.Text, txtPLPIB.Text, txtPLDel.Text, txtPLOps.Text, txtPLUl.Text, txtPLBroj.Text, txtPLBrojPoste.Text, txtPLMatBr.Text,
                            //txtPLBrTelefona.Text, txtVZRacun.Text, txtTPNap.Text, txtTPDatum.Text, txtTPVremeOd.Text, txtTPVremeDo.Text, txtTPImeiPrezime.Text, txtTPBrDokumentaStr.Text,
                            //cbTPKontr1.Text, cbTPKontr2.Text, txtBrTelStranke.Text, txtVKM.Text, txtPoslednjeUverenjeBroj.Text, txtPoslednjeUverenjeDatum.Text, cbDecOzn.Text,
                            //cbVOKS.Text, cbBojaL.Text, cbBojaD.Text, ispravnost, cbVPV.Text);
                            //    neispravnost n = new neispravnost();

                            b.unesiKorisnika(txtTPID.Text, txtRegBr.Text, cbVrstaVoz.Text, cbVPV.Text, cbVKlasa.Text, cbVOblikKar.Text, cbVSpec.Text, cbVMarka.Text, txtVModel.Text,
                                txtVTip.Text, txtVGodinaProizvodnje.Text, txtVBrojSasije.Text, txtVBrojMotora.Text, cbVBoja.Text, txtVPrvaReg.Text, txtVOdnos.Text,
                                txtVSnaga.Text, txtVZapremina.Text, txtVNajMasa.Text, txtVTezina.Text, txtVNosivost.Text, txtVSedenje1.Text, txtVStajanje.Text,
                                txtVBrojOsovina.Text, cbVGorivo.Text, txtVHomologacija.Text, txtVIme.Text, txtVPrezime.Text, txtVJMBG.Text, txtVBRLK.Text, txtVBrPoste.Text,
                                cbVMesto.Text, txtVUlica.Text, txtVBroj.Text, txtVBrTelefona.Text, txtVZiroRacun.Text, txtDLIme.Text, txtDLPrezime.Text, txtDLJMBG.Text,
                                txtDLBRLK.Text, txtDLBrPoste.Text, cbDLMesto.Text, txtDLUlica.Text, txtDLBroj.Text, txtDLBrTelefona.Text, txtTPNap.Text, txtTPDatum.Text,
                                txtTPVremeOd.Text, txtTPVremeDo.Text, cbTPKontr1.Text, cbTPKontr2.Text, txtVKM.Text, txtPoslednjeUverenjeBroj.Text, txtPoslednjeUverenjeDatum.Text,
                                cbDecOzn.Text, cbVOKS.Text, cbBojaL.Text, cbBojaD.Text, ispravnost, txtVBrSati.Text, cbVKuka.Text, txtVOpterecenjeN1.Text, txtVOpterecenjeN2.Text,
                                txtVOpterecenjeN3.Text, txtVOpterecenjeN4.Text, txtVOpterecenjeN5.Text, txtVPneumatikL1.Text, txtVPneumatikL2.Text, txtVPneumatikL3.Text, txtVPneumatikL4.Text,
                                txtVPneumatikL5.Text, txtVPneumatikD1.Text, txtVPneumatikD2.Text, txtVPneumatikD3.Text, txtVPneumatikD4.Text, txtVPneumatikD5.Text, cbUredjajZaZaust.Text, terensko
                                , txtVKomentar.Text, txtVRejting.Value.ToString(), txtVrstaPlusDecOznk.Text, datumIsteka.Text, txtVD.Text, txtVS.Text, txtVV.Text, txtVBrDokumenta.Text, txtVSerijskiBroj.Text, "Ne", txtVVlaznostKT.Text, cbOsiguranje.Text, v.Prioritet, txtDLZiroRacun.Text, tbDatumIzdavanja.Text);
                            neispravnost n = new neispravnost();
                            int i = 0;
                            //b.brisiNeispravnostiPregleda(txtTPID.Text);
                            while (i < dataGridView1.Rows.Count - 1)
                            {
                                n.ID = dataGridView1.Rows[i].Cells[0].FormattedValue.ToString();
                                n.Stepen = dataGridView1.Rows[i].Cells[3].FormattedValue.ToString();

                                b.unesiNeispravnostiPregleda(n, txtTPID.Text);

                                i++;
                            }

                            btnIzv.Visible = true;
                            btnPot2.Visible = true;
                            btnReg2.Visible = true;
                            btnZap1.Visible = true;

                            b.uvecajPreglede(Brpregleda);
                            //ucitavanje tablica
                            cbTablice.DataSource = b.ucitajTablice();
                            prvoUcitavanje = false;
                            cbTablice.SelectedIndex = -1;

                            lblVrstaVoz.ForeColor = SystemColors.ControlText; lblVOblikKar.ForeColor = SystemColors.ControlText; lblVMarka.ForeColor = SystemColors.ControlText;
                            lblVModel.ForeColor = SystemColors.ControlText; lblVGodinaProizvodnje.ForeColor = SystemColors.ControlText; lblVBrojSasije.ForeColor = SystemColors.ControlText;
                            lblVBrojMotora.ForeColor = SystemColors.ControlText; lblVBoja.ForeColor = SystemColors.ControlText; lblVPrvaReg.ForeColor = SystemColors.ControlText;
                            lblVSnaga.ForeColor = SystemColors.ControlText; lblVZapremina.ForeColor = SystemColors.ControlText; lblVNajMasa.ForeColor = SystemColors.ControlText;
                            lblVTezina.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText;
                            lblVBrojOsovina.ForeColor = SystemColors.ControlText; lblVGorivo.ForeColor = SystemColors.ControlText; lblTPID.ForeColor = SystemColors.ControlText;
                            lblIspravnost.ForeColor = SystemColors.ControlText; lblPredjenikm.ForeColor = SystemColors.ControlText; lblVremeOd.ForeColor = SystemColors.ControlText;
                            lblVremeDo.ForeColor = SystemColors.ControlText;
                            lblUZZ.ForeColor = SystemColors.ControlText;
                            lblPneumatici.ForeColor = SystemColors.ControlText;
                            lblVlaznost.ForeColor = SystemColors.ControlText;
                            label62.ForeColor = SystemColors.ControlText;

                        }
                        catch (OleDbException ex)
                        {
                            MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));
                        }
                    }

                    else if (MessageBox.Show("Podaci za zadati ID broj već postoje u bazi, da li želite da ih apdejtujete sa novim vrednostima?\nSve prethodne vrednosti biće zamenjene novim!", "Upozorenje", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            b.brisiRegistraciju(txtTPID.Text);
                        }
                        catch { }
                        try
                        {
                            //    b.unesiKorisnika(txtTPID.Text, txtRegBr.Text, cbVrstaVoz.Text, cbVKlasa.Text, cbVOblikKar.Text, cbVSpec.Text, cbVZemlja.Text,
                            //cbVMarka.Text, txtVModel.Text, txtVTip.Text, txtVGodinaProizvodnje.Text, txtVBrojSasije.Text, txtVBrojMotora.Text, cbVNamena.Text, cbVBoja.Text,
                            //txtVPrvaReg.Text, txtVOdnos.Text, txtVSnaga.Text, txtVZapremina.Text, txtVNajMasa.Text, txtVTezina.Text, txtVNosivost.Text, txtVSedenje1.Text,
                            //txtVStajanje.Text, txtVBrojOsovina.Text, txtVBrPogOsovina.Text, cbVGorivo.Text, txtVHomologacija.Text, txtVJMBG.Text, txtVPrezime.Text,
                            //txtVIme.Text, txtVBRLK.Text, txtFLOps.Text, txtVBrPoste.Text, cbVMesto.Text, txtVUlica.Text, txtVBroj.Text, txtVBrTelefona.Text, txtFLEmail.Text,
                            //txtPLNaziv.Text, txtPLMatBr.Text, txtPLPIB.Text, txtPLDel.Text, txtPLOps.Text, txtPLUl.Text, txtPLBroj.Text, txtPLBrojPoste.Text, txtPLMatBr.Text,
                            //txtPLBrTelefona.Text, txtVZRacun.Text, txtTPNap.Text, txtTPDatum.Text, txtTPVremeOd.Text, txtTPVremeDo.Text, txtTPImeiPrezime.Text, txtTPBrDokumentaStr.Text,
                            //cbTPKontr1.Text, cbTPKontr2.Text, txtBrTelStranke.Text, txtVKM.Text, txtPoslednjeUverenjeBroj.Text, txtPoslednjeUverenjeDatum.Text, cbDecOzn.Text,
                            //cbVOKS.Text, cbBojaL.Text, cbBojaD.Text, ispravnost, cbVPV.Text);
                            b.unesiKorisnika(txtTPID.Text, txtRegBr.Text, cbVrstaVoz.Text, cbVPV.Text, cbVKlasa.Text, cbVOblikKar.Text, cbVSpec.Text, cbVMarka.Text, txtVModel.Text,
                                txtVTip.Text, txtVGodinaProizvodnje.Text, txtVBrojSasije.Text, txtVBrojMotora.Text, cbVBoja.Text, txtVPrvaReg.Text, txtVOdnos.Text,
                                txtVSnaga.Text, txtVZapremina.Text, txtVNajMasa.Text, txtVTezina.Text, txtVNosivost.Text, txtVSedenje1.Text, txtVStajanje.Text,
                                txtVBrojOsovina.Text, cbVGorivo.Text, txtVHomologacija.Text, txtVIme.Text, txtVPrezime.Text, txtVJMBG.Text, txtVBRLK.Text, txtVBrPoste.Text,
                                cbVMesto.Text, txtVUlica.Text, txtVBroj.Text, txtVBrTelefona.Text, txtVZiroRacun.Text, txtDLIme.Text, txtDLPrezime.Text, txtDLJMBG.Text,
                                txtDLBRLK.Text, txtDLBrPoste.Text, cbDLMesto.Text, txtDLUlica.Text, txtDLBroj.Text, txtDLBrTelefona.Text, txtTPNap.Text, txtTPDatum.Text,
                                txtTPVremeOd.Text, txtTPVremeDo.Text, cbTPKontr1.Text, cbTPKontr2.Text, txtVKM.Text, txtPoslednjeUverenjeBroj.Text, txtPoslednjeUverenjeDatum.Text,
                                cbDecOzn.Text, cbVOKS.Text, cbBojaL.Text, cbBojaD.Text, ispravnost, txtVBrSati.Text, cbVKuka.Text, txtVOpterecenjeN1.Text, txtVOpterecenjeN2.Text,
                                txtVOpterecenjeN3.Text, txtVOpterecenjeN4.Text, txtVOpterecenjeN5.Text, txtVPneumatikL1.Text, txtVPneumatikL2.Text, txtVPneumatikL3.Text, txtVPneumatikL4.Text,
                                txtVPneumatikL5.Text, txtVPneumatikD1.Text, txtVPneumatikD2.Text, txtVPneumatikD3.Text, txtVPneumatikD4.Text, txtVPneumatikD5.Text, cbUredjajZaZaust.Text, terensko
                                , txtVKomentar.Text, txtVRejting.Value.ToString(), txtVrstaPlusDecOznk.Text, datumIsteka.Text, txtVD.Text, txtVS.Text, txtVV.Text, txtVBrDokumenta.Text, txtVSerijskiBroj.Text, "Ne", txtVVlaznostKT.Text, cbOsiguranje.Text, v.Prioritet, txtDLZiroRacun.Text, tbDatumIzdavanja.Text);
                            neispravnost n = new neispravnost();
                            int i = 0;

                            b.brisiNeispravnostiPregleda(txtTPID.Text);
                            while (i < dataGridView1.Rows.Count - 1)
                            {
                                n.ID = dataGridView1.Rows[i].Cells[0].FormattedValue.ToString();
                                n.Stepen = dataGridView1.Rows[i].Cells[3].FormattedValue.ToString();

                                b.unesiNeispravnostiPregleda(n, txtTPID.Text);

                                i++;
                            }

                            btnIzv.Visible = true;
                            btnPot2.Visible = true;
                            btnReg2.Visible = true;
                            btnZap1.Visible = true;

                            //b.uvecajPreglede(Brpregleda);

                            //ucitavanje tablica
                            cbTablice.DataSource = b.ucitajTablice();
                            prvoUcitavanje = false;
                            cbTablice.SelectedIndex = -1;

                            lblVrstaVoz.ForeColor = SystemColors.ControlText; lblVOblikKar.ForeColor = SystemColors.ControlText; lblVMarka.ForeColor = SystemColors.ControlText;
                            lblVModel.ForeColor = SystemColors.ControlText; lblVGodinaProizvodnje.ForeColor = SystemColors.ControlText; lblVBrojSasije.ForeColor = SystemColors.ControlText;
                            lblVBrojMotora.ForeColor = SystemColors.ControlText; lblVBoja.ForeColor = SystemColors.ControlText; lblVPrvaReg.ForeColor = SystemColors.ControlText;
                            lblVSnaga.ForeColor = SystemColors.ControlText; lblVZapremina.ForeColor = SystemColors.ControlText; lblVNajMasa.ForeColor = SystemColors.ControlText;
                            lblVTezina.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText;
                            lblVBrojOsovina.ForeColor = SystemColors.ControlText; lblVGorivo.ForeColor = SystemColors.ControlText; lblTPID.ForeColor = SystemColors.ControlText;
                            lblIspravnost.ForeColor = SystemColors.ControlText; lblPredjenikm.ForeColor = SystemColors.ControlText; lblVremeOd.ForeColor = SystemColors.ControlText;
                            lblVremeDo.ForeColor = SystemColors.ControlText;
                            lblUZZ.ForeColor = SystemColors.ControlText;
                            lblPneumatici.ForeColor = SystemColors.ControlText;
                            lblVlaznost.ForeColor = SystemColors.ControlText;
                            label62.ForeColor = SystemColors.ControlText;

                        }
                        catch (OleDbException ex)
                        {
                            MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));
                        }
                    }

                }


            }
            else
            {
                MessageBox.Show("Popunite sva obavezna polja na tabovima \'Podaci o vozilu\' i \'Tehnički pregled\'.");

                lblVrstaVoz.ForeColor = Color.Red; lblVOblikKar.ForeColor = Color.Red; lblVMarka.ForeColor = Color.Red; lblVModel.ForeColor = Color.Red;
                lblVGodinaProizvodnje.ForeColor = Color.Red; lblVBrojSasije.ForeColor = Color.Red; lblVBrojMotora.ForeColor = Color.Red; lblVBoja.ForeColor = Color.Red;
                lblVPrvaReg.ForeColor = Color.Red; lblVSnaga.ForeColor = Color.Red; lblVZapremina.ForeColor = Color.Red; lblVNajMasa.ForeColor = Color.Red;
                lblVTezina.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVBrojOsovina.ForeColor = Color.Red;
                lblVGorivo.ForeColor = Color.Red; lblTPID.ForeColor = Color.Red; lblPredjenikm.ForeColor = Color.Red; lblIspravnost.ForeColor = Color.Red;
                lblVremeOd.ForeColor = Color.Red; lblVremeDo.ForeColor = Color.Red;
                lblUZZ.ForeColor = Color.Red;
                lblPneumatici.ForeColor = Color.Red; label74.ForeColor = Color.Red;
                lblVlaznost.ForeColor = Color.Red; label62.ForeColor = Color.Red; lblVSedenje.ForeColor = Color.Red;
            }
        }

        //nalazenje podataka u bazi za dati broj sasije
        private void txtsasija_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    vozilo v = b.reg(txtsasija.Text.Trim());

                    txtTPID.Text = v.IDBroj;
                    txtRegBr.Text = v.RegistarskiBroj;
                    cbVrstaVoz.Text = v.VrstaVozila;
                    cbVPV.Text = v.PodvrstaVozila;
                    cbVKlasa.Text = v.klasa;
                    cbVOblikKar.Text = v.oblikkaroserije;
                    cbUredjajZaZaust.Text = v.zaustavljanje;
                    cbVOKS.Text = v.OblikKaroserijeSkraceno;
                    cbVSpec.Text = v.specijalno;
                    if (cbVSpec.Text != "")
                        checkBox12.Checked = true;
                    else
                        checkBox12.Checked = false;
                    cbVMarka.Text = v.Marka;
                    txtVModel.Text = v.Model;
                    txtVTip.Text = v.Tip;
                    txtVGodinaProizvodnje.Text = v.GodinaProizvodnje;
                    txtVBrojSasije.Text = v.BrojSasije;
                    txtVBrojMotora.Text = v.BrojMotora;
                    cbVBoja.Text = v.Boja;
                    txtVPrvaReg.Text = v.DatumPrveRegostracije;
                    txtVOdnos.Text = v.Odnos;
                    txtVSnaga.Text = v.Snaga;
                    txtVZapremina.Text = v.Zapremina;
                    txtVNajMasa.Text = v.NajvecaMasa;
                    txtVTezina.Text = v.Tezina;
                    txtVNosivost.Text = v.Nosivost;
                    txtVSedenje1.Text = v.Sedenje;
                    txtVStajanje.Text = v.Stajanje;
                    txtVBrojOsovina.Text = v.BrojOsovina;
                    cbVGorivo.Text = v.Gorivo;
                    txtVHomologacija.Text = v.Homologacija;
                    txtVIme.Text = v.ImeVL;
                    txtVPrezime.Text = v.PrezimeVL;
                    txtVJMBG.Text = v.MaticniBrojVL;
                    txtVBRLK.Text = v.BRLKVL;
                    txtVBrPoste.Text = v.BrojPosteVL;
                    cbVMesto.Text = v.MestoVL;
                    txtVUlica.Text = v.UlicaVL;
                    txtVBroj.Text = v.BrojVL;
                    txtVBrTelefona.Text = v.TelefonaVL;
                    txtVZiroRacun.Text = v.ZiroRacunVL;
                    txtDLIme.Text = v.ImeDL;
                    txtDLPrezime.Text = v.PrezimeDL;
                    txtDLJMBG.Text = v.MaticniBrojDL;
                    txtDLBRLK.Text = v.BRLKDL;
                    txtDLBrPoste.Text = v.BrojPosteDL;
                    cbDLMesto.Text = v.MestoDL;
                    txtDLUlica.Text = v.UlicaDL;
                    txtDLBroj.Text = v.BrojDL;
                    txtDLBrTelefona.Text = v.TelefonaDL;
                    txtTPNap.Text = v.Napomena;
                    txtTPDatum.Text = v.Datum;
                    txtTPVremeOd.CustomFormat = "HH:mm";
                    txtTPVremeDo.CustomFormat = "HH:mm";
                    txtTPVremeOd.Text = v.VremeOd;
                    txtTPVremeDo.Text = v.VremeDo;
                    cbTPKontr1.Text = v.Kontrolor1;
                    cbTPKontr2.Text = v.Kontrolor2;
                    txtVKM.Text = v.Predjenikm;
                    txtPoslednjeUverenjeBroj.Text = v.PoslednjeUverenjeBroj;
                    txtPoslednjeUverenjeDatum.Text = v.PoslednjeUverenjeDatum;
                    cbDecOzn.Text = v.DecimalnaOznaka;
                    cbBojaL.Text = v.BojaL;
                    cbBojaD.Text = v.BojaD;

                    txtVBrSati.Text = v.BrSati;
                    cbVKuka.Text = v.Kuka;
                    txtVOpterecenjeN1.Text = v.OpterecenjeN1; txtVOpterecenjeN2.Text = v.OpterecenjeN2;
                    txtVOpterecenjeN3.Text = v.OpterecenjeN3;
                    txtVOpterecenjeN4.Text = v.OpterecenjeN4;
                    txtVOpterecenjeN5.Text = v.OpterecenjeN5;
                    txtVPneumatikL1.Text = v.PneumatikL1;
                    txtVPneumatikL2.Text = v.PneumatikL2;
                    txtVPneumatikL3.Text = v.PneumatikL3;
                    txtVPneumatikL4.Text = v.PneumatikL4;
                    txtVPneumatikL5.Text = v.PneumatikL5;
                    txtVPneumatikD1.Text = v.PneumatikD1;
                    txtVPneumatikD2.Text = v.PneumatikD2;
                    txtVPneumatikD3.Text = v.PneumatikD3;
                    txtVPneumatikD4.Text = v.PneumatikD4;
                    txtVPneumatikD5.Text = v.PneumatikD5;
                    cbUredjajZaZaust.Text = v.zaustavljanje;
                    txtVKomentar.Text = v.komentar;
                    txtVrstaPlusDecOznk.Text = v.decplusoznaka;
                    datumIsteka.Text = v.datumIsteka;

                    if (txtDLIme.Text != "")
                        rbDL.Checked = true;
                    else
                        rbVL.Checked = true;
                    radioButton2_CheckedChanged(sender, e);
                    int a = 1;
                    try
                    {
                        a = int.Parse(v.rejting);
                    }
                    catch
                    {
                        a = 1;
                    }
                    txtVRejting.Value = a;

                    if (v.Ispravnost == "Da")
                    {
                        rbTehIsp.Checked = true;
                        ispravnost = "Da";
                    }
                    else if (v.Ispravnost == "Ne")
                    {
                        rbTehNeis.Checked = true;
                        ispravnost = "Ne";
                    }
                    else if (v.Ispravnost == "Prekid")
                    {
                        rbTehIspPrekid.Checked = true;
                        ispravnost = "Prekid";
                    }
                    if (v.terensko == "Da")
                    {
                        cbTerensko.Checked = true;
                        terensko = "Da";
                    }
                    else if (v.terensko == "Ne")
                    {
                        cbTerensko.Checked = false;
                        terensko = "Ne";
                    }
                    popuniTabeluNeispravnosti(txtTPID.Text);
                    cbVPV.Text = v.PodvrstaVozila;
                    btnIzv.Visible = true;
                    btnPot2.Visible = true;
                    btnReg2.Visible = true;
                    btnZap1.Visible = true;
                    txtImePrezimePretraga.Text = "";
                    txtsasija.Clear();
                    if (v.ZiroRacunVL != "")
                        checkBox2.Checked = true;
                    else
                        checkBox4.Checked = true;
                    popuniFakture();
                    oznaciVrstuPregleda(txtTPID.Text[8]);
                    txtTPID.Text = v.IDBroj;
                }
                catch
                {
                    MessageBox.Show("Nije pronadjen broj sasije");
                    bunifuThinButton219_Click(sender, e);
                }
            }
        }

        //stampanje potvrde
        private void bunifuThinButton213_Click(object sender, EventArgs e)
        {
            if (txtVKM.Text != "" && cbVrstaVoz.Text != "" && (checkBox12.Checked && cbVOblikKar.Text == "" ? true : (cbVOblikKar.Text == "" ? false : true)) && cbVMarka.Text != "" && txtVModel.Text != "" && txtVGodinaProizvodnje.Text != ""
                && txtVBrojSasije.Text != "" && txtVBrojMotora.Text != "" && cbVBoja.Text != "" && cbBojaL.Text != "" && cbBojaD.Text != "" && txtVPrvaReg.Text != ""
                && txtVSnaga.Text != "" && txtVZapremina.Text != "" && txtVNajMasa.Text != "" && txtVTezina.Text != "" && txtVSedenje1.Text != "" && txtVStajanje.Text != ""
                && txtVBrojOsovina.Text != "" && cbVGorivo.Text != "" && txtTPID.Text != "" && ispravnost != "" && cbVPV.Text != "" && (checkBox12.Checked && cbVOKS.Text == "" ? true : (cbVOKS.Text == "" ? false : true))
                && txtTPVremeOd.Text != "  :" && txtTPVremeDo.Text != "  :")
            {
                int n;
                if (txtTPID.Text.Length < 14 || txtTPID.Text.Length > 14)
                {
                    MessageBox.Show("ID Broj je nepravilan!");
                }
                else if (int.TryParse(txtTPID.Text.Substring(9), out n) == false)
                {
                    MessageBox.Show("ID Broj je nepravilan!");
                }

                else if (rbTehNeis.Checked == true)
                {
                    MessageBox.Show("Tehnička ispranost vozila je obeležena sa \'Ne\'!");
                }

                else
                {
                    lblVrstaVoz.ForeColor = SystemColors.ControlText; lblVOblikKar.ForeColor = SystemColors.ControlText; lblVMarka.ForeColor = SystemColors.ControlText;
                    lblVModel.ForeColor = SystemColors.ControlText; lblVGodinaProizvodnje.ForeColor = SystemColors.ControlText; lblVBrojSasije.ForeColor = SystemColors.ControlText;
                    lblVBrojMotora.ForeColor = SystemColors.ControlText; lblVBoja.ForeColor = SystemColors.ControlText; lblVPrvaReg.ForeColor = SystemColors.ControlText;
                    lblVSnaga.ForeColor = SystemColors.ControlText; lblVZapremina.ForeColor = SystemColors.ControlText; lblVNajMasa.ForeColor = SystemColors.ControlText;
                    lblVTezina.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText;
                    lblVBrojOsovina.ForeColor = SystemColors.ControlText; lblVGorivo.ForeColor = SystemColors.ControlText; lblTPID.ForeColor = SystemColors.ControlText;
                    lblIspravnost.ForeColor = SystemColors.ControlText; lblPredjenikm.ForeColor = SystemColors.ControlText; lblVremeOd.ForeColor = SystemColors.ControlText;
                    lblVremeDo.ForeColor = SystemColors.ControlText;

                    PrintDialog printDialog = new PrintDialog();

                    PrintDocument printDocument = new PrintDocument();

                    printDialog.Document = printDocument; //add the document to the dialog box...        

                    printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt3); //add an event handler that will do the printing

                    //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

                    DialogResult result = printDialog.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        printDocument.Print();

                    }
                }

            }

            else
            {
                MessageBox.Show("Popunite sva obavezna polja na tabovima \'Podaci o vozilu\' i \'Tehnički pregled\'.");

                lblVrstaVoz.ForeColor = Color.Red; lblVOblikKar.ForeColor = Color.Red; lblVMarka.ForeColor = Color.Red; lblVModel.ForeColor = Color.Red;
                lblVGodinaProizvodnje.ForeColor = Color.Red; lblVBrojSasije.ForeColor = Color.Red; lblVBrojMotora.ForeColor = Color.Red; lblVBoja.ForeColor = Color.Red;
                lblVPrvaReg.ForeColor = Color.Red; lblVSnaga.ForeColor = Color.Red; lblVZapremina.ForeColor = Color.Red; lblVNajMasa.ForeColor = Color.Red;
                lblVTezina.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVBrojOsovina.ForeColor = Color.Red;
                lblVGorivo.ForeColor = Color.Red; lblTPID.ForeColor = Color.Red; lblPredjenikm.ForeColor = Color.Red; lblIspravnost.ForeColor = Color.Red;
                lblVremeOd.ForeColor = Color.Red; lblVremeDo.ForeColor = Color.Red;
            }
        }

        //kreiranje potvrde
        public void CreateReceipt3(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int height = 0;
            int xpos = 18, ypos = 40;
            int i = 0;
            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.5f);
            Font FontNaznacenimanji = new Font("Times New Roman", 11, FontStyle.Bold);
            Font FontNaznaceni = new Font("Times New Roman", 12, FontStyle.Bold);
            Font FontNajjaci = new Font("Times New Roman", 14, FontStyle.Bold /*| FontStyle.Underline*/);
            Font Fontoznaka = new Font("Times New Roman", 10);
            Font Fontobicni = new Font("Times New Roman", 10);
            Font Fontobicni2 = new Font("Times New Roman", 10);
            Font fontUpisivanje = new Font("Times New Roman", 10, FontStyle.Italic);
            Font za_licnu_kartu = new Font("Times New Roman", 8);
            Font FontSvoji = new Font("Times New Roman", 14);
            Font fontmali = new Font("Times New Roman", 8);

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;

            StringFormat stringFormat2 = new StringFormat();
            stringFormat2.Alignment = StringAlignment.Near;
            stringFormat2.LineAlignment = StringAlignment.Center;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;

            //stampanje svojih podataka
            e.Graphics.DrawString("VAAS DOO Sjenica", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 400, 25));
            e.Graphics.DrawString("1 4 1 0 2 2", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 794, 15), stringFormat);
            e.Graphics.DrawString("Образац 1", Fontoznaka, Brushes.Black, new Rectangle(xpos, ypos, 770, 30), stringFormatdesno);

            ypos += 40;
            e.Graphics.DrawString("ПОТВРДА О  ТЕХНИЧКОЈ ИСПРАВНОСТИ ВОЗИЛА", FontNajjaci, Brushes.Black, new Rectangle(xpos, ypos, 780, 20), stringFormat);
            ypos += 190;
            int pom = ypos;
            //stampanje podataka o kupcu
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 150, 30));
            e.Graphics.DrawString("ИД број", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 150, ypos, 600, 30));
            e.Graphics.DrawString(txtTPID.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 150, ypos, 400, 30), stringFormat2);
            ypos += 60;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Регистарска ознака", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(txtRegBr.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Идентификациона (VIN) ознака  ", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(txtVBrojSasije.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Врста (ознака врсте)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(cbVrstaVoz.Text + " " + cbVPV.Text + ((cbTerensko.Checked) ? "G" : ""), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Облик каросерије - подврста\n(ознака подврсте)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 250, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(txtVrstaPlusDecOznk.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Марка", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(cbVMarka.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Тип/варијанта/верзија", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(txtVTip.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Комерцијална ознака (модел)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 250, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(txtVModel.Text + " ", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 750, 100));
            e.Graphics.DrawString("Напомена: " + txtTPNap.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 2, 730, 100));
            ypos += 150;
            e.Graphics.DrawString("ВОЗИЛО ЈЕ ТЕХНИЧКИ ИСПРАВНО", FontNajjaci, Brushes.Black, new Rectangle(xpos, ypos, 750, 20), stringFormat);
            ypos += 60;
            e.Graphics.DrawString("Датум вршењна техничког прегледа: " + txtTPDatum.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 600, 20));
            ypos += 30;
            e.Graphics.DrawString("Име и презие контролора и потпис: " + cbTPKontr1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 750, 20));
            e.Graphics.DrawString("_______________________________________", Fontobicni, Brushes.Black, new Rectangle(xpos + 215, ypos + 1, 770, 20));
            ypos += 50;
            e.Graphics.DrawString("МП", Fontobicni, Brushes.Black, new Rectangle(xpos + 400, ypos, 300, 20));

        }

        //stampanje izvestaja
        private void bunifuThinButton214_Click(object sender, EventArgs e)
        {
            i = 0;
            int brListova = brNeispravnost / 8 + 1;
            if (brNeispravnost % 8 == 0)
                brListova--;
            if (rbTehIsp.Checked)
                brListova = 1;
            for (int k = 0; k < brListova; k++)
            {
                if (dataGridView1.RowCount <= 1)
                {
                    MessageBox.Show("Unesite neispravnosti!");
                }
                else
                {
                    PrintDialog printDialog = new PrintDialog();

                    PrintDocument printDocument = new PrintDocument();

                    printDialog.Document = printDocument; //add the document to the dialog box...        

                    printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt4); //add an event handler that will do the printing

                    //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

                    DialogResult result = printDialog.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        printDocument.Print();

                    }

                }
            }

        }

        //kreiranje izvestaja
        public void CreateReceipt4(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int height = 0;
            int xpos = 18, ypos = 40;
            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.5f);
            Font FontNaznacenimanji = new Font("Times New Roman", 11, FontStyle.Bold);
            Font FontNaznaceni = new Font("Times New Roman", 12, FontStyle.Bold);
            Font FontNajjaci = new Font("Times New Roman", 14, FontStyle.Bold);
            Font Fontoznaka = new Font("Times New Roman", 10);
            Font Fontobicni = new Font("Times New Roman", 10);
            Font Fontobicni2 = new Font("Times New Roman", 10, FontStyle.Bold);
            Font fontUpisivanje = new Font("Times New Roman", 10, FontStyle.Italic);
            Font za_licnu_kartu = new Font("Times New Roman", 8);
            Font FontSvoji = new Font("Times New Roman", 14);
            Font fontmali = new Font("Times New Roman", 8);
            Font fontnajmanji = new Font("Times New Roman", 6);

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;

            StringFormat stringFormat2 = new StringFormat();
            stringFormat2.Alignment = StringAlignment.Near;
            stringFormat2.LineAlignment = StringAlignment.Center;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;

            //stampanje svojih podataka
            e.Graphics.DrawString("VAAS DOO Sjenica", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 400, 25));
            e.Graphics.DrawString("1 4 1 0 2 2", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 770, 15), stringFormat);
            e.Graphics.DrawString("Образац 2", Fontoznaka, Brushes.Black, new Rectangle(xpos, ypos, 770, 30), stringFormatdesno);

            ypos += 40;
            e.Graphics.DrawString("ИЗВЕШТАЈ О УТВРЂЕНОМ СТАЊУ ВОЗИЛА", FontNajjaci, Brushes.Black, new Rectangle(xpos, ypos, 770, 20), stringFormat);
            ypos += 50;
            int pom = ypos;
            //stampanje podataka o kupcu
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 150, 30));
            e.Graphics.DrawString("ИД број", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 150, ypos, 400, 30));
            e.Graphics.DrawString(txtTPID.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 150, ypos, 380, 30), stringFormat2);
            ypos += 50;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Врста техничког прегледа", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 300, 30));
            e.Graphics.DrawString(comboBox2.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 290, 30), stringFormat2);
            ypos += 50;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Регистарска ознака", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(txtRegBr.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Идентификациона ознака (VIN) ", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(txtVBrojSasije.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Врста (ознака врсте)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(cbVrstaVoz.Text + " " + cbVPV.Text + ((cbTerensko.Checked) ? "G" : ""), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 40));
            e.Graphics.DrawString("Облик каросерије - подврста\n(ознака подврсте)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 230, 40), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 40));
            e.Graphics.DrawString(txtVrstaPlusDecOznk.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 40), stringFormat2);
            ypos += 40;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Марка", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(cbVMarka.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Тип/варијанта/верзија", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(txtVTip.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 250, 30));
            e.Graphics.DrawString("Комерцијална ознака (модел)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 250, 30), stringFormat2);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 250, ypos, 500, 30));
            e.Graphics.DrawString(txtVModel.Text + " ", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 250, ypos, 400, 30), stringFormat2);
            ypos += 40;
            e.Graphics.DrawString("УТВРЂЕНЕ НЕИСПРАВНОСТИ", Fontobicni2, Brushes.Black, new Rectangle(xpos, ypos, 794, 20));
            ypos += 22;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 210, 30));
            e.Graphics.DrawString("Назив дела уређаја", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 200, 30), stringFormat);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 210, ypos, 392, 30));
            e.Graphics.DrawString("Неисправност/код", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 200, ypos, 420, 30), stringFormat);
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 602, ypos, 150, 30));
            e.Graphics.DrawString("Степен", Fontobicni, Brushes.Black, new Rectangle(xpos + 602, ypos, 130, 30), stringFormat);
            ypos += 30;
            int tempYpos = ypos;
            while (i < brNeispravnost)
            {
                try
                {
                    e.Graphics.DrawString(dataGridView1.Rows[i].Cells[1].FormattedValue.ToString().Length < 36 ? dataGridView1.Rows[i].Cells[1].FormattedValue.ToString() : dataGridView1.Rows[i].Cells[1].FormattedValue.ToString().Substring(0, 33) + "...", Fontobicni, Brushes.Black, new Rectangle(xpos, ypos, 220, 20), stringFormat2);
                    e.Graphics.DrawString(dataGridView1.Rows[i].Cells[2].FormattedValue.ToString().Length < 57 ? dataGridView1.Rows[i].Cells[2].FormattedValue.ToString() : dataGridView1.Rows[i].Cells[2].FormattedValue.ToString().Substring(0, 54) + "...", Fontobicni, Brushes.Black, new Rectangle(xpos + 212, ypos, 500, 20), stringFormat2);
                    e.Graphics.DrawString(dataGridView1.Rows[i].Cells[0].FormattedValue.ToString(), Fontobicni, Brushes.Black, new Rectangle(xpos + 550, ypos, 500, 20), stringFormat2);
                    e.Graphics.DrawString(dataGridView1.Rows[i].Cells[3].FormattedValue.ToString().Length < 57 ? dataGridView1.Rows[i].Cells[3].FormattedValue.ToString() : dataGridView1.Rows[i].Cells[3].FormattedValue.ToString().Substring(0, 54) + "...", Fontobicni, Brushes.Black, new Rectangle(xpos + 604, ypos, 140, 20), stringFormat2);
                }
                catch { }
                ypos += 20;
                i++;
                if (i % 8 == 0)
                {
                    break;
                }
            }
            for (int k = 0; k < 8; k++)
            {
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos, tempYpos, 210, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 210, tempYpos, 392, 20));
                e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 602, tempYpos, 150, 20));
                tempYpos += 20;
            }
            ypos = tempYpos;
            ypos += 15;
            e.Graphics.DrawString("ВОЗИЛО ЈЕ ТЕХНИЧКИ НЕИСПРАВНО", FontNajjaci, Brushes.Black, new Rectangle(xpos, ypos, 780, 20), stringFormat);
            ypos += 30;
            e.Graphics.DrawString("СТЕПЕН НЕИСПРАВНОСТИ ВОЗИЛА:", FontNaznaceni, Brushes.Black, new Rectangle(xpos, ypos, 780, 20), stringFormat);
            ypos += 30;
            e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 750, 100));
            e.Graphics.DrawString("Напомена: " + txtTPNap.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 730, 90));
            ypos += 130;
            e.Graphics.DrawString("Датум вршења техничког прегледа: " + txtTPDatum.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 600, 20));
            ypos += 30;
            //e.Graphics.DrawString(cbTPKontr1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 780, 20));
            //e.Graphics.DrawString("_______________________________________", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos + 1, 780, 20));
            //ypos += 20;
            //e.Graphics.DrawString("Име и презиме контролора и потпис", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 300, 20), stringFormat);
            //ypos += 40;
            //e.Graphics.DrawString("МП", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 780, 20), stringFormat);

            e.Graphics.DrawString("Име и презие контролора и потпис: " + cbTPKontr1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 750, 20));
            e.Graphics.DrawString("_______________________________________", Fontobicni, Brushes.Black, new Rectangle(xpos + 215, ypos + 1, 770, 20));
            ypos += 50;
            e.Graphics.DrawString("МП", Fontobicni, Brushes.Black, new Rectangle(xpos + 400, ypos, 300, 20));
        }

        //odabir fizickog lica
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                checkBox2.Checked = false;


                txtVZiroRacun.Enabled = false;
                groupBox3.Visible = true;
                groupBox2.Visible = false;
                //txtPLNaziv.Enabled = false;
                //txtPLMatBr.Enabled = false;
                //txtPLPIB.Enabled = false;
                //txtPLDel.Enabled = false;
                //txtPLOps.Enabled = false;
                //txtPLUl.Enabled = false;
                //txtPLBroj.Enabled = false;
                //txtPLBrojPoste.Enabled = false;
                //cbPLMesto.Enabled = false;
                //txtPLBrTelefona.Enabled = false;
                //txtVZRacun.Enabled = false;
            }
        }

        //odabir pravnog lica
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                checkBox4.Checked = false;
                groupBox3.Visible = false;
                groupBox2.Visible = true;
                //gbDL.Enabled = true;
                //txtPLNaziv.Enabled = true;
                //txtPLMatBr.Enabled = true;
                //txtPLPIB.Enabled = true;
                //txtPLDel.Enabled = true;
                //txtPLOps.Enabled = true;
                //txtPLUl.Enabled = true;
                //txtPLBroj.Enabled = true;
                //txtPLBrojPoste.Enabled = true;
                //cbPLMesto.Enabled = true;
                //txtPLBrTelefona.Enabled = true;
                //txtVZRacun.Enabled = true;


                txtVZiroRacun.Enabled = true;
            }
        }


        //usnos nove boje vozila u bazu
        //private void bunifuThinButton26_Click(object sender, EventArgs e)
        //{
        //    string sifra;
        //    string naziv;
        //    string vrsta = promptBox.ShowDialog("Boja vozila", "Unesite novu boju vozila");
        //    if (vrsta != "")
        //    {
        //        try
        //        {
        //            sifra = vrsta.Substring(0, 2);
        //            naziv = vrsta.Substring(3);
        //            b.unesiBoju(sifra, naziv);
        //            cbVBoja.DataSource = null;
        //            List<string> boje = new List<string>();
        //            foreach (var b in b.lista_boja())
        //            {
        //                boje.Add(b.BojaSifra + " " + b.BojaNaziv);
        //            }
        //            cbVBoja.DataSource = boje;
        //            cbVBoja.Text = "";

        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message);
        //        }

        //    }

        //}

        //omogucuje specijalna cb za specijalna vozila
        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox12.Checked)
            {
                cbVSpec.Enabled = true;
                cbVOblikKar.Enabled = false;
                cbVOKS.Enabled = false;
                cbDecOzn.DataSource = null;
                cbDecOzn.Items.Add("/01 Platforma");
                cbDecOzn.Items.Add("/02 Sa stranicama");
                cbDecOzn.Items.Add("/03 Zatvorena nadradnja (furgon)");
                cbDecOzn.Items.Add("/04 Klimatizovana nadgradnja sa izolacionim zidovima I opremom za održavanje unutrašnje temperature");
                cbDecOzn.Items.Add("/05 Klimatizovana nadgradnja sa izolacionim zidovima ali bez opreme za održavanje temperature");
                cbDecOzn.Items.Add("/06 Fleksibilne stranice (cerada)");
                cbDecOzn.Items.Add("/07 Izmenjiva nadgradnja (transportni sud)");
                cbDecOzn.Items.Add("/08 Nosač kontenera");
                cbDecOzn.Items.Add("/09 Vozilo opremljeno sa huk-liftom");
                cbDecOzn.Items.Add("/10 Kiper");
                cbDecOzn.Items.Add("/11 Cisterna");
                cbDecOzn.Items.Add("/12 Cisterna namenjena za prevoz opasnih tereta");
                cbDecOzn.Items.Add("/13 Prevoz stoke");
                cbDecOzn.Items.Add("/14 Prevoz vozila");
                cbDecOzn.Items.Add("/15 Mešalica za beton");
                cbDecOzn.Items.Add("/16 Pumpa za beton");
                cbDecOzn.Items.Add("/17 Prevoz drva");
                cbDecOzn.Items.Add("/18 Odvožnje smeća");
                cbDecOzn.Items.Add("/19 Čišćenje ulica, čišćenje odvoda, pranje ulica");
                cbDecOzn.Items.Add("/20 Kompresor");
                cbDecOzn.Items.Add("/21 Prevoz čamaca");
                cbDecOzn.Items.Add("/22 Prevoz jedrilica");
                cbDecOzn.Items.Add("/23 Vozilo prodavnica ili izlog");
                cbDecOzn.Items.Add("/24 Vozilo za pomoć na putu");
                cbDecOzn.Items.Add("/25 Vozilo sa merdevinama");
                cbDecOzn.Items.Add("/26 Dizalica (različita od pokretne dizalice SF)");
                cbDecOzn.Items.Add("/27 Platformska dizalica");
                cbDecOzn.Items.Add("/28 Vozilo sa uređajem za bušenje i kopanje");
                cbDecOzn.Items.Add("/29 Niskopodno priključno vozilo");
                cbDecOzn.Items.Add("/30 Prevoz stakala");
                cbDecOzn.Items.Add("/31 Vatrogasno vozilo");
                cbDecOzn.Items.Add("/99 Oblik karoserije ne postoji u važećoj listi");
            }
            else
            {
                cbVSpec.Enabled = false;
                cbVSpec.Text = "";
                cbVOblikKar.Enabled = true;
                cbVOKS.Enabled = true;
            }
        }

        //promena indeksa kod vrste vozila
        private void cbVrstaVoz_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (vrstaVozilaPrvo == false)
            {
                List<string> lista = new List<string>();
                foreach (var k in b.lista_ka(cbVrstaVoz.Text))
                {
                    lista.Add(k.KaroserijaNaziv);
                }
                cbVOblikKar.DataSource = lista;
                cbVOblikKar.Text = "";
                cbVOKS.Text = "";
                txtVrstaPlusDecOznk.Text = "";

                List<string> listaPodVrsta = new List<string>();
                foreach (var k in b.lista_podVrs(cbVrstaVoz.Text))
                {
                    listaPodVrsta.Add(k.KaroserijaNaziv);
                }
                cbVPV.DataSource = listaPodVrsta;
                cbVPV.Text = "";


                if (cbVrstaVoz.Text == "AUTOBUS")
                {
                    cbVKlasa.Enabled = true;
                    label76.Enabled = true;
                }
                else
                {

                    cbVKlasa.Enabled = false;
                    label76.Enabled = false;
                    cbVKlasa.Text = "";


                }
            }

        }

        //unos nove vrste vozila u bazu
        private void bunifuThinButton212_Click(object sender, EventArgs e)
        {
            string vrsta = promptBox.ShowDialog("Vrsta vozila", "Unesite novu vrstu vozila");
            int sifra = cbVrstaVoz.Items.Count;
            if (vrsta != "")
            {
                try
                {
                    b.unesiVrstu(sifra, vrsta);
                    cbVrstaVoz.DataSource = null;
                    List<string> str = new List<string>();
                    foreach (var i in b.vrvozila())
                    {
                        str.Add(i.NazivVrste);
                    }
                    cbVrstaVoz.DataSource = str;
                    cbVrstaVoz.Text = "";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        //unos nove karoserije vozila u bazu
        private void bunifuThinButton24_Click(object sender, EventArgs e)
        {
            string vrsta = promptBox.ShowDialog("Karoserija vozila", "Unesite novu karoseriju vozila");
            int sifra = cbVrstaVoz.SelectedIndex;
            if (vrsta != "")
            {
                try
                {
                    b.unesiKaroseriju(sifra, vrsta);
                    cbVOblikKar.DataSource = null;
                    List<string> lista = new List<string>();
                    foreach (var k in b.lista_ka(cbVrstaVoz.Text))
                    {
                        lista.Add(k.KaroserijaNaziv);
                    }
                    cbVOblikKar.DataSource = lista;
                    cbVOblikKar.Text = "";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        //unos nove podvrste u bazu
        private void bunifuThinButton22_Click(object sender, EventArgs e)
        {
            string vrsta = promptBox.ShowDialog("Podvrsta vozila", "Unesite novu podvrstu vozila");
            int sifra = cbVrstaVoz.SelectedIndex;

        }

        //unos nove marke vozila u bazu
        private void bunifuThinButton25_Click(object sender, EventArgs e)
        {
            string vrsta = promptBox.ShowDialog("Marka vozila", "Unesite novu marku vozila");
            if (vrsta != "")
            {
                try
                {
                    b.unesiMarku(vrsta);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        //pretraga datog id u bazi
        private void txtPretragaID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    vozilo v = b.id(txtPretragaID.Text.Trim());
                    txtTPID.Text = v.IDBroj;
                    txtRegBr.Text = v.RegistarskiBroj;
                    cbVrstaVoz.Text = v.VrstaVozila;
                    cbVPV.Text = v.PodvrstaVozila;
                    cbVKlasa.Text = v.klasa;
                    cbVOblikKar.Text = v.oblikkaroserije;
                    cbUredjajZaZaust.Text = v.zaustavljanje;
                    cbVOKS.Text = v.OblikKaroserijeSkraceno;
                    cbVSpec.Text = v.specijalno;
                    if (cbVSpec.Text != "")
                        checkBox12.Checked = true;
                    else
                        checkBox12.Checked = false;
                    cbVMarka.Text = v.Marka;
                    txtVModel.Text = v.Model;
                    txtVTip.Text = v.Tip;
                    txtVGodinaProizvodnje.Text = v.GodinaProizvodnje;
                    txtVBrojSasije.Text = v.BrojSasije;
                    txtVBrojMotora.Text = v.BrojMotora;
                    cbVBoja.Text = v.Boja;
                    txtVPrvaReg.Text = v.DatumPrveRegostracije;
                    txtVOdnos.Text = v.Odnos;
                    txtVSnaga.Text = v.Snaga;
                    txtVZapremina.Text = v.Zapremina;
                    txtVNajMasa.Text = v.NajvecaMasa;
                    txtVTezina.Text = v.Tezina;
                    txtVNosivost.Text = v.Nosivost;
                    txtVSedenje1.Text = v.Sedenje;
                    txtVStajanje.Text = v.Stajanje;
                    txtVBrojOsovina.Text = v.BrojOsovina;
                    cbVGorivo.Text = v.Gorivo;
                    txtVHomologacija.Text = v.Homologacija;
                    txtVIme.Text = v.ImeVL;
                    txtVPrezime.Text = v.PrezimeVL;
                    txtVJMBG.Text = v.MaticniBrojVL;
                    txtVBRLK.Text = v.BRLKVL;
                    txtVBrPoste.Text = v.BrojPosteVL;
                    cbVMesto.Text = v.MestoVL;
                    txtVUlica.Text = v.UlicaVL;
                    txtVBroj.Text = v.BrojVL;
                    txtVBrTelefona.Text = v.TelefonaVL;
                    txtVZiroRacun.Text = v.ZiroRacunVL;
                    txtDLIme.Text = v.ImeDL;
                    txtDLPrezime.Text = v.PrezimeDL;
                    txtDLJMBG.Text = v.MaticniBrojDL;
                    txtDLBRLK.Text = v.BRLKDL;
                    txtDLBrPoste.Text = v.BrojPosteDL;
                    cbDLMesto.Text = v.MestoDL;
                    txtDLUlica.Text = v.UlicaDL;
                    txtDLBroj.Text = v.BrojDL;
                    txtDLBrTelefona.Text = v.TelefonaDL;
                    txtTPNap.Text = v.Napomena;
                    txtTPDatum.Text = v.Datum;
                    txtTPVremeOd.CustomFormat = "HH:mm";
                    txtTPVremeDo.CustomFormat = "HH:mm";
                    txtTPVremeOd.Text = v.VremeOd;
                    txtTPVremeDo.Text = v.VremeDo;

                    cbTPKontr1.Text = v.Kontrolor1;
                    cbTPKontr2.Text = v.Kontrolor2;
                    txtVKM.Text = v.Predjenikm;
                    txtPoslednjeUverenjeBroj.Text = v.PoslednjeUverenjeBroj;
                    txtPoslednjeUverenjeDatum.Text = v.PoslednjeUverenjeDatum;
                    cbDecOzn.Text = v.DecimalnaOznaka;
                    cbBojaL.Text = v.BojaL;
                    cbBojaD.Text = v.BojaD;

                    txtVBrSati.Text = v.BrSati;
                    cbVKuka.Text = v.Kuka;
                    txtVOpterecenjeN1.Text = v.OpterecenjeN1; txtVOpterecenjeN2.Text = v.OpterecenjeN2;
                    txtVOpterecenjeN3.Text = v.OpterecenjeN3;
                    txtVOpterecenjeN4.Text = v.OpterecenjeN4;
                    txtVOpterecenjeN5.Text = v.OpterecenjeN5;
                    txtVPneumatikL1.Text = v.PneumatikL1;
                    txtVPneumatikL2.Text = v.PneumatikL2;
                    txtVPneumatikL3.Text = v.PneumatikL3;
                    txtVPneumatikL4.Text = v.PneumatikL4;
                    txtVPneumatikL5.Text = v.PneumatikL5;
                    txtVPneumatikD1.Text = v.PneumatikD1;
                    txtVPneumatikD2.Text = v.PneumatikD2;
                    txtVPneumatikD3.Text = v.PneumatikD3;
                    txtVPneumatikD4.Text = v.PneumatikD4;
                    txtVPneumatikD5.Text = v.PneumatikD5;
                    cbUredjajZaZaust.Text = v.zaustavljanje;
                    txtVKomentar.Text = v.komentar;
                    txtVrstaPlusDecOznk.Text = v.decplusoznaka;
                    datumIsteka.Text = v.datumIsteka;
                    if (txtDLIme.Text != "")
                        rbDL.Checked = true;
                    else
                        rbVL.Checked = true;
                    radioButton2_CheckedChanged(sender, e);
                    int a = 1;
                    try
                    {
                        a = int.Parse(v.rejting);
                    }
                    catch
                    {
                        a = 1;
                    }
                    txtVRejting.Value = a;

                    if (v.Ispravnost == "Da")
                    {
                        rbTehIsp.Checked = true;
                        ispravnost = "Da";
                    }
                    else if (v.Ispravnost == "Ne")
                    {
                        rbTehNeis.Checked = true;
                        ispravnost = "Ne";
                    }
                    else if (v.Ispravnost == "Prekid")
                    {
                        rbTehIspPrekid.Checked = true;
                        ispravnost = "Prekid";
                    }
                    if (v.terensko == "Da")
                    {
                        cbTerensko.Checked = true;
                        terensko = "Da";
                    }
                    else if (v.terensko == "Ne")
                    {
                        cbTerensko.Checked = false;
                        terensko = "Ne";
                    }
                    popuniTabeluNeispravnosti(txtTPID.Text);
                    cbVPV.Text = v.PodvrstaVozila;
                    btnIzv.Visible = true;
                    btnPot2.Visible = true;
                    btnReg2.Visible = true;
                    btnZap1.Visible = true;
                    txtImePrezimePretraga.Text = "";
                    txtPretragaID.Clear();
                    if (v.ZiroRacunVL != "")
                        checkBox2.Checked = true;
                    else
                        checkBox4.Checked = true;
                    if (v.DLZiroRacun != "")
                        cbDLPravno.Checked = true;
                    else
                        cbDLFizicko.Checked = true;
                    popuniFakture();
                    oznaciVrstuPregleda(txtTPID.Text[8]);
                    txtTPID.Text = v.IDBroj;
                }
                catch
                {
                    MessageBox.Show("Trazeni ID broj ne postoji");
                    bunifuThinButton219_Click(sender, e);
                }
            }
        }

        //popunjavanje tabele za neispravnosti kod izvestaja
        private void popuniTabeluNeispravnosti(string text)
        {
            var lista = b.popuniNeispravnost(text);
            dataGridView1.Rows.Clear();
            foreach (var v in lista)
            {
                dataGridView1.Rows.Add(v.ID, v.Deo, v.Neispravnost, v.Stepen);
            }
            brNeispravnost = dataGridView1.Rows.Count - 1;
            //dataGridView1.Columns[0].Visible = false;
        }

        //stampanje fakture
        private void bunifuThinButton22_Click_1(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();

            PrintDocument printDocument = new PrintDocument();

            printDialog.Document = printDocument; //add the document to the dialog box...        

            printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt5); //add an event handler that will do the printing

            //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

            DialogResult result = printDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                printDocument.Print();

            }
        }

        //kreiranje fakture
        public void CreateReceipt5(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            i = 0;
            strana++;
            int pocetak = 5;
            int sredina = 395;
            int height = 10;
            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.1f);
            Font svojipodaci_font = new Font("Arial", 14);
            Font svojipodaci_font_bold = new Font("Arial", 15, FontStyle.Bold);
            Font tekst = new Font("Arial", 10);
            Font obicni = new Font("Arial", 12);
            Font obicni_bold = new Font("Arial", 12, FontStyle.Bold);
            Font za_ukupno = new Font("Arial", 10, FontStyle.Bold);
            Font tabela_font = new Font("Arial", 8);
            Font otpremnica_font = new Font("Arial", 14, FontStyle.Bold);
            Font otpremnica_font_opdvuceno = new Font("Arial", 14, FontStyle.Bold | FontStyle.Underline);
            dataGridView2.Font = tekst;

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Near;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;

            //e.Graphics.DrawString("Strana " + strana, dataGridView2.Font, Brushes.Black, new Rectangle(sredina, 1100, 500, 20));

            //stampanje svojih podataka
            if (i < 20)
            {
                e.Graphics.DrawString("\"VAAS\" DOO SJENICA", svojipodaci_font_bold, Brushes.Black, new Rectangle(300, height, 500, 500), stringFormat);
                height += svojipodaci_font_bold.Height + 2;
                e.Graphics.DrawString("Ul. Nikole Tesle 1, 36310 Sjenica \nPib 104436359, matični broj 20159626\nŽiro račun 200-3410880101944-68\nTelefon 063/8875657", obicni, Brushes.Black, new Rectangle(300, height, 500, 650), stringFormat);
                height += 3 * obicni.Height + 30;
                Point p1 = new Point(0, height);
                Point p2 = new Point(900, height);
                e.Graphics.DrawLine(p, p1, p2);
                height += 5;
                e.Graphics.DrawString("Matični broj: " + textBox20.Text, za_ukupno, Brushes.Black, new Rectangle(sredina, height, 900, 600));
                //e.Graphics.DrawString(textBox4.Text, tekst, Brushes.Black, new Rectangle(sredina + 100, height, 900, 600));
                height += za_ukupno.Height + 10;

                using (GraphicsPath path = RoundedRect(new Rectangle(sredina - 5, height - 3, 396, 170), 20))
                {
                    e.Graphics.DrawPath(pe, path);
                }

                e.Graphics.DrawString("Objekat:", obicni_bold, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString("TEHNIČKI PREGLED VOZILA", za_ukupno, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));
                height += obicni_bold.Height;
                e.Graphics.DrawString("Mesto promenta:", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString("Sjenica", tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));

                e.Graphics.DrawString("Kupac", za_ukupno, Brushes.Black, new Rectangle(sredina, height - 5, 330, 600));
                e.Graphics.DrawString(textBox1.Text, otpremnica_font, Brushes.Black, new Rectangle(sredina + 100, height - 10, 290, 600));

                height += obicni_bold.Height;
                e.Graphics.DrawString("Datum izdavanja računa:", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString(dateTimePicker1.Value.ToString("dd.MM.yyyy"), tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));
                height += obicni_bold.Height;
                e.Graphics.DrawString("Datum prometa:", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString(dateTimePicker1.Value.ToString("dd.MM.yyyy"), tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));

                e.Graphics.DrawString("Adresa", za_ukupno, Brushes.Black, new Rectangle(sredina, height + 3, 330, 600));
                e.Graphics.DrawString(textBox2.Text + ", " + textBox3.Text, tekst, Brushes.Black, new Rectangle(sredina + 100, height + 3, 300, 600));

                height += obicni_bold.Height;
                e.Graphics.DrawString("Datum plaćanja:", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString(dateTimePicker2.Value.ToString("dd.MM.yyyy"), tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));
                height += obicni_bold.Height;
                e.Graphics.DrawString("Način plaćanja", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString("VIRMANOM", tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));

                e.Graphics.DrawString("PIB", za_ukupno, Brushes.Black, new Rectangle(sredina, height, 900, 600));
                e.Graphics.DrawString(textBox4.Text, tekst, Brushes.Black, new Rectangle(sredina + 100, height, 900, 600));
                height += obicni_bold.Height;
                e.Graphics.DrawString("Broj fiskalnog isečka:", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString(txtBI.Text, tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));

                e.Graphics.DrawString("TELEFON", za_ukupno, Brushes.Black, new Rectangle(sredina, height, 900, 600));
                e.Graphics.DrawString(textBox6.Text, tekst, Brushes.Black, new Rectangle(sredina + 100, height, 900, 600));
                height += obicni_bold.Height;
                e.Graphics.DrawString("Otpremnica:", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString(txtOtp.Text + "/" + (DateTime.Now.Year % 100), tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));

                e.Graphics.DrawString("T.R.", za_ukupno, Brushes.Black, new Rectangle(sredina, height, 900, 600));

                e.Graphics.DrawString(textBox7.Text, tekst, Brushes.Black, new Rectangle(sredina + 100, height, 900, 600));

                height += obicni_bold.Height;

                height += 2 * svojipodaci_font.Height;
                e.Graphics.DrawString("MP Račun-Otpremnica  " + txtOtp.Text + "/" + (DateTime.Now.Year % 100), otpremnica_font_opdvuceno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                height += 1;
            }
            else height = 50;

            height += otpremnica_font.Height + 1;


            //ceo okvir
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak, height, 430, za_ukupno.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 430, height, 110, za_ukupno.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 540, height, 110, za_ukupno.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 650, height, 130, za_ukupno.Height + tekst.Height + 4));
            //ZAGLAVLJEEEEE
            #region data1
            e.Graphics.DrawString("PROIZVOD / USLUGA", tabela_font, Brushes.Black, new Rectangle(pocetak, height + 1, 430, dataGridView2.Rows[0].Height), stringFormat);
            #endregion

            #region data2
            e.Graphics.DrawString("RABAT", tabela_font, Brushes.Black, new Rectangle(pocetak + 430, height + 1, 110, dataGridView2.Rows[0].Height), stringFormat);
            #endregion

            #region data3
            e.Graphics.DrawString("PDV", tabela_font, Brushes.Black, new Rectangle(pocetak + 540, height + 1, 110, dataGridView2.Rows[0].Height), stringFormat);
            #endregion

            #region data4
            e.Graphics.DrawString("Cena sa PDV-om", tabela_font, Brushes.Black, new Rectangle(pocetak + 650, height + 8, 60, 2 * dataGridView2.Rows[0].Height), stringFormat);
            #endregion

            #region data5
            e.Graphics.DrawString("Vrednost sa PDV-om", tabela_font, Brushes.Black, new Rectangle(pocetak + 710, height + 8, 70, 2 * dataGridView2.Rows[0].Height), stringFormat);
            #endregion


            height += za_ukupno.Height + 2;
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak, height, 20, tekst.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 20, height, 50, tekst.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 70, height, 150, tekst.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 220, height, 30, tekst.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 250, height, 40, tekst.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 290, height, 70, tekst.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 360, height, 70, tekst.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 430, height, 40, tekst.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 470, height, 70, tekst.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 540, height, 40, tekst.Height + 2));
            e.Graphics.DrawRectangle(p, new Rectangle(pocetak + 580, height, 70, tekst.Height + 2));
            //ZAGLAVLJEEEEE
            #region data1
            e.Graphics.DrawString("RB", tabela_font, Brushes.Black, new Rectangle(pocetak, height + 1, 20, dataGridView2.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("Šifra", tabela_font, Brushes.Black, new Rectangle(pocetak + 20, height + 1, 50, dataGridView2.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("Naziv", tabela_font, Brushes.Black, new Rectangle(pocetak + 70, height + 1, 150, dataGridView2.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("JM", tabela_font, Brushes.Black, new Rectangle(pocetak + 220, height + 1, 30, dataGridView2.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("Kol.", tabela_font, Brushes.Black, new Rectangle(pocetak + 250, height + 1, 40, dataGridView2.Rows[0].Height), stringFormatdesno);
            #endregion

            #region data1
            e.Graphics.DrawString("Cena", tabela_font, Brushes.Black, new Rectangle(pocetak + 290, height + 1, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
            #endregion

            #region data1
            e.Graphics.DrawString("Iznos", tabela_font, Brushes.Black, new Rectangle(pocetak + 360, height + 1, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
            #endregion

            #region data1
            e.Graphics.DrawString("%", tabela_font, Brushes.Black, new Rectangle(pocetak + 430, height + 1, 40, dataGridView2.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("Iznos", tabela_font, Brushes.Black, new Rectangle(pocetak + 470, height + 1, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
            #endregion

            #region data1
            e.Graphics.DrawString("%", tabela_font, Brushes.Black, new Rectangle(pocetak + 540, height + 1, 40, dataGridView2.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("Porez", tabela_font, Brushes.Black, new Rectangle(pocetak + 580, height + 1, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
            #endregion

            height += dataGridView2.Rows[0].Height;
            Point prva = new Point(pocetak, height);
            Point zadnja = new Point(790, height);
            e.Graphics.DrawLine(pe, prva, zadnja);

            // Pocetna visina na kojoj se nalaze podaci iz tabele
            height += 2;


            while (i < dataGridView2.Rows.Count - 1)
            {
                string txt = dataGridView2.Rows[i].Cells[1].FormattedValue.ToString();
                if (txt.Length > 28)
                {
                    e.Graphics.DrawString(txt, tabela_font, Brushes.Black, new Rectangle(pocetak + 70, height, 150, 2 * dataGridView2.Rows[0].Height));
                    height += tabela_font.Height;
                }
                else
                    e.Graphics.DrawString(txt, tabela_font, Brushes.Black, new Rectangle(pocetak + 70, height, 150, dataGridView2.Rows[0].Height));

                e.Graphics.DrawString((i + 1).ToString(), tabela_font, Brushes.Black, new Rectangle(pocetak, height, 20, dataGridView2.Rows[0].Height), stringFormatdesno);
                e.Graphics.DrawString(dataGridView2.Rows[i].Cells[0].FormattedValue.ToString(), tabela_font, Brushes.Black, new Rectangle(pocetak + 22, height, 50, dataGridView2.Rows[0].Height));
                //if (txt.Length > 28)
                //    e.Graphics.DrawString(txt.Substring(0, 28), tabela_font, Brushes.Black, new Rectangle(pocetak + 70, height, 150, dataGridView2.Rows[0].Height));
                //else
                e.Graphics.DrawString(dataGridView2.Rows[i].Cells[2].FormattedValue.ToString(), tabela_font, Brushes.Black, new Rectangle(pocetak + 220, height, 30, dataGridView2.Rows[0].Height));

                e.Graphics.DrawString((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 250, height, 40, dataGridView2.Rows[0].Height), stringFormatdesno);

                e.Graphics.DrawString((Convert.ToDouble(dataGridView2.Rows[i].Cells[4].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 290, height, 70, dataGridView2.Rows[0].Height), stringFormatdesno);

                e.Graphics.DrawString((Convert.ToDouble(dataGridView2.Rows[i].Cells[5].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 360, height, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
                ukupno_bez_rabata += Convert.ToDouble(dataGridView2.Rows[i].Cells[5].FormattedValue);
                e.Graphics.DrawString((Convert.ToDouble(dataGridView2.Rows[i].Cells[6].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 430, height, 40, dataGridView2.Rows[0].Height), stringFormatdesno);

                e.Graphics.DrawString((Convert.ToDouble(dataGridView2.Rows[i].Cells[7].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 470, height, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
                ukupan_rabat += Convert.ToDouble(dataGridView2.Rows[i].Cells[7].FormattedValue);
                e.Graphics.DrawString((Convert.ToInt32(dataGridView2.Rows[i].Cells[8].FormattedValue)).ToString(), tabela_font, Brushes.Black, new Rectangle(pocetak + 540, height, 40, dataGridView2.Rows[0].Height), stringFormatdesno);

                e.Graphics.DrawString((Convert.ToDouble(dataGridView2.Rows[i].Cells[9].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 580, height, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
                ukupan_pdv += Convert.ToDouble(dataGridView2.Rows[i].Cells[9].FormattedValue);
                e.Graphics.DrawString((Convert.ToDouble(dataGridView2.Rows[i].Cells[10].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 650, height, 60, dataGridView2.Rows[0].Height), stringFormatdesno);

                e.Graphics.DrawString((Convert.ToDouble(dataGridView2.Rows[i].Cells[11].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 710, height, 70, dataGridView2.Rows[0].Height), stringFormatdesno);

                height += tabela_font.Height + 1;
                prva = new Point(pocetak, height);
                zadnja = new Point(790, height);
                e.Graphics.DrawLine(pe, prva, zadnja);
                if (i == dataGridView2.Rows.Count - 2)
                    e.Graphics.DrawLine(p, prva, zadnja);
                height += 1;
                //e.Graphics.DrawString((Convert.ToDouble(dataGridView2.Rows[i].Cells[7].FormattedValue)).ToString("N", new CultureInfo("is-IS")), dataGridView2.Font, Brushes.Black, new Rectangle(600, height, 100, dataGridView2.Rows[0].Height), stringFormatdesno);
                //potrazuje += Convert.ToDouble(dataGridView2.Rows[i].Cells[7].FormattedValue);
                //duguje += Convert.ToDouble(dataGridView2.Rows[i].Cells[6].FormattedValue);
                //e.Graphics.DrawString((Convert.ToDouble(dataGridView2.Rows[i].Cells[8].FormattedValue)).ToString("N", new CultureInfo("is-IS")), dataGridView2.Font, Brushes.Black, new Rectangle(700, height, 100, dataGridView2.Rows[0].Height), stringFormatdesno);
                //saldo = Convert.ToDouble(dataGridView2.Rows[i].Cells[8].FormattedValue);
                //if (i == 40 || (i > 20 && i+40 % 60 == 0) || (i == dataGridView2.Rows.Count - 2 && dataGridView2.Rows.Count - 1 > 40))
                if (dataGridView2.Rows.Count - 1 > 25 && i == dataGridView2.Rows.Count - 2)
                {
                    i++;

                    e.HasMorePages = true;
                    return;
                }
                i++;
            }

            height += dataGridView2.Rows[0].Height;

            e.Graphics.DrawString("Ukupno bez PDV-a", tekst, Brushes.Black, new Rectangle(pocetak + 470, height, 150, dataGridView2.Rows[0].Height));
            e.Graphics.DrawString(ukupno_bez_rabata.ToString("N", new CultureInfo("is-IS")), tekst, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView2.Rows[0].Height), stringFormatdesno);
            height += tekst.Height;
            e.Graphics.DrawString("PDV ukupno", tekst, Brushes.Black, new Rectangle(pocetak + 470, height, 150, dataGridView2.Rows[0].Height));
            e.Graphics.DrawString(ukupan_pdv.ToString("N", new CultureInfo("is-IS")), tekst, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView2.Rows[0].Height), stringFormatdesno);

            e.Graphics.DrawString("IZNOSI SU ISKAZANI U VALUTI PLAĆANJA - Dinar", tekst, Brushes.Black, new Rectangle(pocetak, height - 5, 500, dataGridView2.Rows[0].Height));

            height += tekst.Height;
            e.Graphics.DrawString("Rabat", tekst, Brushes.Black, new Rectangle(pocetak + 470, height, 150, dataGridView2.Rows[0].Height));
            e.Graphics.DrawString((ukupno_bez_rabata - ukupan_rabat).ToString("N", new CultureInfo("is-IS")), tekst, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView2.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString("NAPOMENA: Obveznik nije u sistemu PDV-a", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 500, dataGridView2.Rows[0].Height));

            height += tekst.Height;
            prva = new Point(pocetak + 470, height);
            zadnja = new Point(790, height);
            e.Graphics.DrawLine(pe, prva, zadnja);
            ukupno += ukupno_bez_rabata - ukupan_rabat + ukupan_pdv;
            height += 2;
            e.Graphics.DrawString("Svega za plaćanje", za_ukupno, Brushes.Black, new Rectangle(pocetak + 470, height, 150, dataGridView2.Rows[0].Height));
            e.Graphics.DrawString(ukupan_rabat.ToString("N", new CultureInfo("is-IS")), za_ukupno, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView2.Rows[0].Height), stringFormatdesno);
            height += za_ukupno.Height;
            e.Graphics.DrawString("Prilikom uplate pozvati se na broj ", tekst, Brushes.Black, new Rectangle(pocetak, height, 500, dataGridView2.Rows[0].Height));
            height += za_ukupno.Height;

            e.Graphics.DrawString("Tarifa", tabela_font, Brushes.Black, new Rectangle(pocetak + 290, height, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString("Osnovica", tabela_font, Brushes.Black, new Rectangle(pocetak + 450, height, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString("PDV", tabela_font, Brushes.Black, new Rectangle(pocetak + 600, height, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString("Iznos", tabela_font, Brushes.Black, new Rectangle(pocetak + 710, height, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
            height += tabela_font.Height;
            prva = new Point(pocetak + 290, height);
            zadnja = new Point(790, height);
            e.Graphics.DrawLine(pe, prva, zadnja);
            height += 2;
            e.Graphics.DrawString("1", tabela_font, Brushes.Black, new Rectangle(pocetak + 290, height, 70, dataGridView2.Rows[0].Height));
            e.Graphics.DrawString("0,00", tabela_font, Brushes.Black, new Rectangle(pocetak + 290, height, 70, dataGridView2.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString((ukupan_rabat).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 390, height, 130, dataGridView2.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString(ukupan_pdv.ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 540, height, 130, dataGridView2.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString(ukupan_rabat.ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView2.Rows[0].Height), stringFormatdesno);
            height += tabela_font.Height;

            prva = new Point(pocetak + 290, height);
            zadnja = new Point(790, height);
            e.Graphics.DrawLine(pe, prva, zadnja);
            height += 2;

            e.Graphics.DrawString((ukupan_rabat).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 390, height, 130, dataGridView2.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString(ukupan_pdv.ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 540, height, 130, dataGridView2.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString(ukupan_rabat.ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView2.Rows[0].Height), stringFormatdesno);
            height += tabela_font.Height + 20;

            e.Graphics.DrawString("Za plaćanje posle gore navedenog datuma zaračunava se zakonska kamata", tekst, Brushes.Black, new Rectangle(pocetak, height - 5, 500, dataGridView2.Rows[0].Height));
            height += 50;
            e.Graphics.DrawString("Fakturisao", dataGridView2.Font, Brushes.Black, new Rectangle(20, height, 230, 50), stringFormat);
            e.Graphics.DrawString("Ovlašćeno lice", dataGridView2.Font, Brushes.Black, new Rectangle(270, height, 230, 50), stringFormat);
            e.Graphics.DrawString("Račun preuzeo", dataGridView2.Font, Brushes.Black, new Rectangle(550, height, 250, 50), stringFormat);

            height += 40;
            e.Graphics.DrawString("________________________", dataGridView2.Font, Brushes.Black, new Rectangle(20, height, 200, 50), stringFormat);
            e.Graphics.DrawString("M.P.", dataGridView2.Font, Brushes.Black, new Rectangle(230, height, 50, 50), stringFormat);
            e.Graphics.DrawString("________________________", dataGridView2.Font, Brushes.Black, new Rectangle(270, height, 230, 50), stringFormat);
            e.Graphics.DrawString("M.P.", dataGridView2.Font, Brushes.Black, new Rectangle(500, height, 50, 50), stringFormat);
            e.Graphics.DrawString("________________________", dataGridView2.Font, Brushes.Black, new Rectangle(550, height, 250, 50), stringFormat);
            height += dataGridView2.Font.Height;
            e.Graphics.DrawString("Asmir Kolašinac", dataGridView2.Font, Brushes.Black, new Rectangle(270, height, 230, 50), stringFormat);
            height += dataGridView2.Font.Height;

            e.Graphics.DrawString("Broj identifikacionog dokumenta", dataGridView2.Font, Brushes.Black, new Rectangle(550, height, 250, 50), stringFormat);
            height += 40;
            e.Graphics.DrawString("Ovaj dokument je urađen na računaru i punovažan je bez pečata i potpisa", tabela_font, Brushes.Black, new Rectangle(150, height + 3, 380, 50), stringFormat);
            e.Graphics.DrawString("________________________", dataGridView2.Font, Brushes.Black, new Rectangle(550, height, 250, 50), stringFormat);

            i = 0;
            strana = 0;
            ukupno = 0;
            ukupno_bez_rabata = 0;
            ukupan_rabat = 0;
            ukupan_pdv = 0;
            if (MessageBox.Show("Da li želite da povećate broj otpremnice?", "Potvrda", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                brOtpremnice = int.Parse(txtOtp.Text);
                brOtpremnice++;
                b.promeniBrojac(brOtpremnice);
                txtOtp.Text = brOtpremnice.ToString();
            }
        }

        public static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int diameter = radius * 2;
            Size size = new Size(diameter, diameter);
            Rectangle arc = new Rectangle(bounds.Location, size);
            GraphicsPath path = new GraphicsPath();

            if (radius == 0)
            {
                path.AddRectangle(bounds);
                return path;
            }

            // top left arc  
            path.AddArc(arc, 180, 90);

            // top right arc  
            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90);

            // bottom right arc  
            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90);

            // bottom left arc 
            arc.X = bounds.Left;
            path.AddArc(arc, 90, 90);

            path.CloseFigure();
            return path;
        }

        //dodavanje artika u datagrid na fakturi
        private void bunifuThinButton27_Click_1(object sender, EventArgs e)
        {
            i++;
            double rabat = Convert.ToDouble(txtRabat.Text);
            double cena = double.Parse(txtSfr.Text);

            kolicina = Convert.ToDouble(Kom.Value);
            double iznos_rabat = cena * kolicina - ((cena * rabat) / 100) * kolicina;
            double pdv = 0;
            double sa_pdvom = cena;
            double iznos = cena * kolicina;
            dt.Rows.Add(i, txtNaziv.Text, "kom", kolicina, cena, iznos, rabat, iznos_rabat, "0", pdv, iznos_rabat / kolicina, iznos_rabat);
            dataGridView2.DataSource = dt;
            double u = 0;
            int k = 0;
            while (k < dataGridView2.Rows.Count - 1)
            {
                u += Convert.ToDouble(dataGridView2.Rows[k].Cells[11].FormattedValue);

                k++;
            }
            textBox5.Text = u.ToString();
            Kom.Value = 1;
            txtRabat.Text = "0";

        }

        //brisanje izabranog artikla sa datagrida na fakturi
        private void bunifuThinButton23_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView2.Rows.RemoveAt(dataGridView2.CurrentCell.RowIndex);
                i--;
                int k = 0;
                double u = 0;
                while (k < dataGridView2.Rows.Count - 1)
                {
                    u += Convert.ToDouble(dataGridView2.Rows[k].Cells[11].FormattedValue);
                    k++;
                }
                textBox5.Text = u.ToString();
            }
            catch (Exception)
            {


            }

        }

        //stampanje uplatnice
        private void bunifuThinButton211_Click_1(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();

            PrintDocument printDocument = new PrintDocument();

            PaperSize ps = new PaperSize();
            ps.RawKind = (int)PaperKind.A4;
            printDocument.DefaultPageSettings.PaperSize = ps;
            //printDocument.DefaultPageSettings.Landscape = true;
            printDialog.Document = printDocument; //add the document to the dialog box...        

            printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6); //add an event handler that will do the printing

            //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

            DialogResult result = printDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                printDocument.Print();

            }
        }
        public void CreateReceipt16(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int height = 47;
            int px = 26;
            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.5f);
            Font svojipodaci_font = new Font("Arial", 10);
            Font naslov_font = new Font("Arial", 8, FontStyle.Bold);
            Font tekst = new Font("Arial", 8);
            Font ob = new Font("Arial", 5, FontStyle.Bold);
            Font obicni = new Font("Arial", 5);
            Font za_ukupno = new Font("Arial", 8, FontStyle.Bold);
            Font za_licnu_kartu = new Font("Arial", 8);
            Font otpremnica_font = new Font("Arial", 10, FontStyle.Bold);
            Font naslov_font_nebold = new Font("Arial", 8);



            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Near;
            stringFormat.LineAlignment = StringAlignment.Near;

            StringFormat stringFormatCenter = new StringFormat();
            stringFormatCenter.Alignment = StringAlignment.Center;
            stringFormatCenter.LineAlignment = StringAlignment.Center;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;
            //e.Graphics.DrawString("НАЛОГ ЗА ПРЕНОС", naslov_font, Brushes.Black, new Rectangle(407 - 19, height, 400, 50), stringFormatdesno);
            height = int.Parse(textBox11.Text) + 12;
            px = int.Parse(textBox10.Text);
            //stampanje svojih podataka
            //e.Graphics.DrawString("платилац", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            e.Graphics.DrawString(richTextBox4.Text, naslov_font, Brushes.Black, new Rectangle(px, height + 3, 354, 59)/*, stringFormatCenter*/);
            height += 80;
            //e.Graphics.DrawString("сврха плаћања", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            e.Graphics.DrawString(richTextBox6.Text, naslov_font_nebold, Brushes.Black, new Rectangle(px, height, 354, 59)/*, stringFormatCenter*/);
            height += 80 + 6;
            //e.Graphics.DrawString("прималац", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            e.Graphics.DrawString(richTextBox5.Text, naslov_font, Brushes.Black, new Rectangle(px, height, 354, 59)/*, stringFormatCenter*/);

            //Point po = new Point(397, 62);
            //Point po2 = new Point(397, 223 + 59);
            //e.Graphics.DrawLine(pe, po, po2);
            px += 423;
            height = int.Parse(textBox11.Text) + 80 - 62;
            //stampanje svojih podataka
            //e.Graphics.DrawString("шифра плаћања", za_licnu_kartu, Brushes.Black, new Rectangle(436, 51, 52, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 77, 59, 22));
            e.Graphics.DrawString(textBox27.Text, tekst, Brushes.Black, new Rectangle(px, height, 51, 50));
            px += 70;

            //e.Graphics.DrawString("\nвалута", za_licnu_kartu, Brushes.Black, new Rectangle(503, 51, 52, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(502, 77, 59, 22));
            e.Graphics.DrawString(textBox26.Text, tekst, Brushes.Black, new Rectangle(px, height, 51, 50));
            px += 80;
            //e.Graphics.DrawString("\nизнос", za_licnu_kartu, Brushes.Black, new Rectangle(583, 51, 52, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(582, 77, 192 + 26 - 19, 22));
            string izs = "";
            double izsd = 0;
            try
            {
                izsd = double.Parse(textBox25.Text);
                izs = izsd.ToString("N", new CultureInfo("is-IS"));
            }
            catch
            {
                izs = "";
            }
            e.Graphics.DrawString("   =" + izs, za_ukupno, Brushes.Black, new Rectangle(px, height, 196, 50));
            px -= 150;
            height += 58 - 15;
            //e.Graphics.DrawString("\nрачун платиоца", za_licnu_kartu, Brushes.Black, new Rectangle(436, 100, 520, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 126, 339 + 26 - 19, 22));
            e.Graphics.DrawString(textBox23.Text, naslov_font, Brushes.Black, new Rectangle(px, height + 7, 339 + 26 - 19, 20), stringFormatCenter);
            px += 3;
            height += 78 - 18;
            //px = int.Parse(textBox10.Text) + 436;
            //e.Graphics.DrawString("\nмодел и позив на број (задужење)", za_licnu_kartu, Brushes.Black, new Rectangle(436, 174 - 26, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 174, 59, 22));
            e.Graphics.DrawString(textBox21.Text, tekst, Brushes.Black, new Rectangle(px - 5, height, 39, 50));
            px += 80;
            //e.Graphics.DrawRectangle(pe, new Rectangle(502, 174, 272 + 26 - 19, 22));
            e.Graphics.DrawString(textBox19.Text, tekst, Brushes.Black, new Rectangle(px, height, 275, 50));
            height += 68 - 18;
            px -= 80;
            //e.Graphics.DrawString("\nрачун примаоца", za_licnu_kartu, Brushes.Black, new Rectangle(436, 219 - 26, 339, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 219, 339 + 26 - 19, 22));
            //e.Graphics.DrawString(textBox8.Text, naslov_font, Brushes.Black, new Rectangle(px, height, 339 + 26 - 19, 20), stringFormatCenter);
            height += 68 - 18;

            //e.Graphics.DrawString("\nмодел и позив на број (одобрење)", za_licnu_kartu, Brushes.Black, new Rectangle(436, 272 - 26, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 272, 59, 22));
            //e.Graphics.DrawString(textBox9.Text, tekst, Brushes.Black, new Rectangle(px, height, 39, 50));
            px += 70;
            //e.Graphics.DrawRectangle(pe, new Rectangle(502, 272, 272 + 26 - 19, 22));
            //e.Graphics.DrawString(textBox12.Text, tekst, Brushes.Black, new Rectangle(px, height, 275, 50));
            px = int.Parse(textBox10.Text) + 185;
            height += 88 - 18;
            e.Graphics.DrawString(txtUplatniceMesto.Text + "   " + txtUplatniceDatum.Text, tekst, Brushes.Black, new Rectangle(px, height - 15, 500, 50));

        }

        //kreiranje nalog ya prenos
        public void CreateReceipt6(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int height = 47;
            int px = 26;
            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.5f);
            Font svojipodaci_font = new Font("Arial", 10);
            Font naslov_font = new Font("Arial", 8, FontStyle.Bold);
            Font tekst = new Font("Arial", 8);
            Font ob = new Font("Arial", 5, FontStyle.Bold);
            Font obicni = new Font("Arial", 5);
            Font za_ukupno = new Font("Arial", 8, FontStyle.Bold);
            Font za_licnu_kartu = new Font("Arial", 8);
            Font otpremnica_font = new Font("Arial", 10, FontStyle.Bold);
            Font naslov_font_nebold = new Font("Arial", 8);



            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Near;
            stringFormat.LineAlignment = StringAlignment.Near;

            StringFormat stringFormatCenter = new StringFormat();
            stringFormatCenter.Alignment = StringAlignment.Center;
            stringFormatCenter.LineAlignment = StringAlignment.Center;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;
            //e.Graphics.DrawString("НАЛОГ ЗА ПРЕНОС", naslov_font, Brushes.Black, new Rectangle(407 - 19, height, 400, 50), stringFormatdesno);
            height = int.Parse(textBox11.Text) + 12;
            px = int.Parse(textBox10.Text);
            //stampanje svojih podataka
            //e.Graphics.DrawString("платилац", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            e.Graphics.DrawString(richTextBox1.Text, naslov_font, Brushes.Black, new Rectangle(px, height + 3, 354, 59)/*, stringFormatCenter*/);
            height += 80 + 5;
            //e.Graphics.DrawString("сврха плаћања", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            e.Graphics.DrawString(richTextBox2.Text, naslov_font_nebold, Brushes.Black, new Rectangle(px, height, 354, 59)/*, stringFormatCenter*/);
            height += 80 + 8;
            //e.Graphics.DrawString("прималац", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            e.Graphics.DrawString(richTextBox3.Text, naslov_font, Brushes.Black, new Rectangle(px, height, 354, 59)/*, stringFormatCenter*/);

            //Point po = new Point(397, 62);
            //Point po2 = new Point(397, 223 + 59);
            //e.Graphics.DrawLine(pe, po, po2);
            px += 423;
            height = int.Parse(textBox11.Text) + 80 - 62 + 7;
            //stampanje svojih podataka
            //e.Graphics.DrawString("шифра плаћања", za_licnu_kartu, Brushes.Black, new Rectangle(436, 51, 52, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 77, 59, 22));
            e.Graphics.DrawString(textBox18.Text, tekst, Brushes.Black, new Rectangle(px, height, 51, 50));
            px += 70;

            //e.Graphics.DrawString("\nвалута", za_licnu_kartu, Brushes.Black, new Rectangle(503, 51, 52, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(502, 77, 59, 22));
            e.Graphics.DrawString(textBox17.Text, tekst, Brushes.Black, new Rectangle(px, height, 51, 50));
            px += 80;
            //e.Graphics.DrawString("\nизнос", za_licnu_kartu, Brushes.Black, new Rectangle(583, 51, 52, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(582, 77, 192 + 26 - 19, 22));
            string izs = "";
            double izsd = 0;
            try
            {
                izsd = double.Parse(textBox16.Text);
                izs = izsd.ToString("N", new CultureInfo("is-IS"));
            }
            catch
            {
                izs = "";
            }
            e.Graphics.DrawString("   =" + izs, za_ukupno, Brushes.Black, new Rectangle(px, height, 196, 50));
            px -= 150;
            height += 58 - 15;
            //e.Graphics.DrawString("\nрачун платиоца", za_licnu_kartu, Brushes.Black, new Rectangle(436, 100, 520, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 126, 339 + 26 - 19, 22));
            e.Graphics.DrawString(textBox15.Text, naslov_font, Brushes.Black, new Rectangle(px, height + 3, 339 + 26 - 19, 20), stringFormatCenter);
            px += 3;
            height += 78 - 18;
            //px = int.Parse(textBox10.Text) + 436;
            //e.Graphics.DrawString("\nмодел и позив на број (задужење)", za_licnu_kartu, Brushes.Black, new Rectangle(436, 174 - 26, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 174, 59, 22));
            e.Graphics.DrawString(textBox14.Text, tekst, Brushes.Black, new Rectangle(px, height, 39, 50));
            px += 80;
            //e.Graphics.DrawRectangle(pe, new Rectangle(502, 174, 272 + 26 - 19, 22));
            e.Graphics.DrawString(textBox13.Text, tekst, Brushes.Black, new Rectangle(px, height, 275, 50));
            height += 68 - 18 + 2;
            px -= 80;
            //e.Graphics.DrawString("\nрачун примаоца", za_licnu_kartu, Brushes.Black, new Rectangle(436, 219 - 26, 339, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 219, 339 + 26 - 19, 22));
            e.Graphics.DrawString(textBox8.Text, naslov_font, Brushes.Black, new Rectangle(px, height - 9, 339 + 26 - 19, 20), stringFormatCenter);
            height += 68 - 18;

            //e.Graphics.DrawString("\nмодел и позив на број (одобрење)", za_licnu_kartu, Brushes.Black, new Rectangle(436, 272 - 26, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 272, 59, 22));
            e.Graphics.DrawString(textBox9.Text, tekst, Brushes.Black, new Rectangle(px - 5, height, 39, 50));
            px += 70;
            //e.Graphics.DrawRectangle(pe, new Rectangle(502, 272, 272 + 26 - 19, 22));
            e.Graphics.DrawString(textBox12.Text, tekst, Brushes.Black, new Rectangle(px, height, 275, 50));
            px = int.Parse(textBox10.Text) + 185;
            height += 88 - 18 - 3;
            e.Graphics.DrawString(txtUplatniceMesto.Text + "   " + txtUplatniceDatum.Text, tekst, Brushes.Black, new Rectangle(px, height - 15, 500, 50));

        }

        //popunjavanje podataka na uplatnici iz cba
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            richTextBox2.Text = "";
            richTextBox3.Text = "";
            textBox18.Text = "";
            textBox16.Text = "";
            textBox15.Text = "";
            textBox14.Text = "";
            textBox8.Text = "";
            textBox12.Text = "";

            richTextBox5.Text = "";
            richTextBox6.Text = "";
            textBox27.Text = "";
            textBox26.Text = "";
            textBox25.Text = "";
            textBox23.Text = "";
            textBox21.Text = "";
            textBox19.Text = "";
            textBox17.Text = "";
            textBox9.Text = "";

            if (checkBox2.Checked)
            {
                textBox15.Text = txtVZiroRacun.Text;
                //if(rbDL.Checked)
                //    richTextBox1.Text = txtDLIme.Text + " " + txtDLPrezime.Text + "\n" + txtDLUlica.Text + " " + txtDLBroj.Text + "\n" + cbDLMesto.Text;
                //else
                richTextBox1.Text = txtVPrezime.Text + "\n" + txtVUlica.Text + " " + txtVBroj.Text + "\n" + cbVMesto.Text;
            }
            else
            {
                textBox15.Text = txtDLZiroRacun.Text;
                //if (rbDL.Checked)
                //    richTextBox4.Text = txtDLIme.Text + " " + txtDLPrezime.Text + "\n" + txtDLUlica.Text + " " + txtDLBroj.Text + "\n" + cbDLMesto.Text;
                //else
                richTextBox4.Text = txtVIme.Text + " " + txtVPrezime.Text + "\n" + txtVUlica.Text + " " + txtVBroj.Text + "\n" + cbVMesto.Text;
            }

            if (comboBox1.SelectedIndex == 0)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Naknada za registracionu nalepnicu " + txtRegBr.Text;
                    richTextBox3.Text = "MUP";
                    textBox18.Text = "253";
                    textBox17.Text = "RSD";
                    textBox16.Text = "";
                    textBox8.Text = "840-1848-16";
                    textBox12.Text = "";
                    textBox9.Text = "97";
                }
                else
                {
                    richTextBox6.Text = "Naknada za registracionu nalepnicu " + txtRegBr.Text;
                    richTextBox5.Text = "MUP";
                    textBox27.Text = "153";
                    textBox26.Text = "RSD";
                    textBox25.Text = "";
                    textBox23.Text = "840-1848-16";
                    textBox21.Text = "97";
                    textBox19.Text = "";
                }
            }

            if (comboBox1.SelectedIndex == 1)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Registraciona nalepnica, " + txtVBrojSasije.Text;
                    richTextBox3.Text = "ZAVOD ZA IZRADU NOV. I KOV. NOVCA\nBeograd\nBeograd";
                    textBox18.Text = "253";
                    textBox17.Text = "RSD";
                    textBox16.Text = "236";
                    textBox8.Text = "980-333-07";
                    textBox12.Text = "081218";
                }
                else
                {
                    richTextBox6.Text = "Registraciona nalepnica, " + txtVBrojSasije.Text;
                    richTextBox5.Text = "ZAVOD ZA IZRADU NOV. I KOV. NOVCA\nBeograd\nBeograd";
                    textBox27.Text = "153";
                    textBox26.Text = "RSD";
                    textBox25.Text = "236";
                    textBox23.Text = "980-333-07";
                    textBox19.Text = "081218";
                }

            }

            if (comboBox1.SelectedIndex == 2)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Registraciona nalepnica, " + txtVBrojSasije.Text;
                    richTextBox3.Text = "REPUBLIČKA ADMINISTRATIVNA TAKSA\nBeograd\nBeograd";
                    textBox18.Text = "253";
                    textBox17.Text = "RSD";
                    textBox16.Text = "240";
                    textBox8.Text = "840-742221843-57";
                    textBox9.Text = "97";
                    textBox12.Text = "19-091";
                }
                else
                {
                    richTextBox6.Text = "Registraciona nalepnica, " + txtVBrojSasije.Text;
                    richTextBox5.Text = "REPUBLIČKA ADMINISTRATIVNA TAKSA\nBeograd\nBeograd";
                    textBox27.Text = "153";
                    textBox26.Text = "RSD";
                    textBox25.Text = "240";
                    textBox23.Text = "840-742221843-57";
                    textBox21.Text = "97";
                    textBox19.Text = "19-091";
                }

            }

            if (comboBox1.SelectedIndex == 3)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Saobraćajna dozvola, " + txtVBrojSasije.Text;
                    richTextBox3.Text = "ZAVOD ZA IZRADU NOV. I KOV. NOVCA\nBeograd\nBeograd";
                    textBox18.Text = "253";
                    textBox17.Text = "RSD";
                    textBox16.Text = "660";
                    textBox8.Text = "980-333-07"; ;
                    textBox12.Text = "090901";
                }
                else
                {
                    richTextBox6.Text = "Saobraćajna dozvola, " + txtVBrojSasije.Text;
                    richTextBox5.Text = "ZAVOD ZA IZRADU NOV. I KOV. NOVCA\nBeograd\nBeograd";
                    textBox27.Text = "153";
                    textBox26.Text = "RSD";
                    textBox25.Text = "660";
                    textBox23.Text = "980-333-07"; ;
                    textBox19.Text = "090901";
                }

            }

            if (comboBox1.SelectedIndex == 4)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Saobraćajna dozvola, " + txtVBrojSasije.Text;
                    richTextBox3.Text = "REPUBLIČKA ADMINISTRATIVNA TAKSA\nBeograd\nBeograd";
                    textBox18.Text = "253";
                    textBox17.Text = "RSD";
                    textBox16.Text = "780";
                    textBox8.Text = "840-742221843-57";
                    textBox9.Text = "97";
                    textBox12.Text = "19-091";
                }
                else
                {
                    richTextBox6.Text = "Saobraćajna dozvola, " + txtVBrojSasije.Text;
                    richTextBox5.Text = "REPUBLIČKA ADMINISTRATIVNA TAKSA\nBeograd\nBeograd";
                    textBox27.Text = "153";
                    textBox26.Text = "RSD";
                    textBox25.Text = "780";
                    textBox23.Text = "840-742221843-57";
                    textBox21.Text = "97";
                    textBox19.Text = "19-091";
                }
            }

            if (comboBox1.SelectedIndex == 5)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Registarske tablice, " + txtVBrojSasije.Text;
                    richTextBox3.Text = "ZAVOD ZA IZRADU NOV. I KOV. NOVCA\nBeograd\nBeograd";
                    textBox18.Text = "253";
                    textBox17.Text = "RSD";
                    textBox16.Text = "1093";
                    textBox8.Text = "980-333-07"; ;
                    textBox12.Text = "041493";
                }
                else
                {
                    richTextBox6.Text = "Registarske tablice, " + txtVBrojSasije.Text;
                    richTextBox5.Text = "ZAVOD ZA IZRADU NOV. I KOV. NOVCA\nBeograd\nBeograd";
                    textBox27.Text = "153";
                    textBox26.Text = "RSD";
                    textBox25.Text = "1093";
                    textBox23.Text = "980-333-07"; ;
                    textBox19.Text = "041493";
                }
            }
            if (comboBox1.SelectedIndex == 6)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Registarske tablice, " + txtVBrojSasije.Text;
                    richTextBox3.Text = "REPUBLIČKA ADMINISTRATIVNA TAKSA\nBeograd\nBeograd";
                    textBox18.Text = "253";
                    textBox17.Text = "RSD";
                    textBox16.Text = "1310";
                    textBox8.Text = "840-742221843-57";
                    textBox9.Text = "97";
                    textBox12.Text = "19-091";
                }
                else
                {
                    richTextBox6.Text = "Registarske tablice, " + txtVBrojSasije.Text;
                    richTextBox5.Text = "REPUBLIČKA ADMINISTRATIVNA TAKSA\nBeograd\nBeograd";
                    textBox27.Text = "153";
                    textBox26.Text = "RSD";
                    textBox25.Text = "1310";
                    textBox23.Text = "840-742221843-57";
                    textBox21.Text = "97";
                    textBox19.Text = "19-091";
                }
            }
            if (comboBox1.SelectedIndex == 7)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Komunalna taksa, " + txtVBrojSasije.Text;
                    richTextBox3.Text = "KOMUNALNA TAKSA\nSjenica\nSjenica";
                    textBox18.Text = "253";
                    textBox17.Text = "RSD";
                    textBox16.Text = "";
                    textBox8.Text = "840-714513843-04";
                    textBox9.Text = "97";
                    textBox12.Text = "19-091";
                }
                else
                {
                    richTextBox6.Text = "Komunalna taksa, " + txtVBrojSasije.Text;
                    richTextBox5.Text = "KOMUNALNA TAKSA\nSjenica\nSjenica";
                    textBox27.Text = "153";
                    textBox26.Text = "RSD";
                    textBox25.Text = "";
                    textBox23.Text = "840-714513843-04";
                    textBox21.Text = "97";
                    textBox19.Text = "19-091";
                }
            }
            if (comboBox1.SelectedIndex == 9)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Uplata premije za osiguranje od autoodgovornosti po polisi\nbr. ";
                    richTextBox3.Text = "GFO NOVI PAZAR\nŠabana Koče 71\nNovi Pazar";
                    textBox18.Text = "221";
                    textBox17.Text = "RSD";
                    textBox16.Text = "";
                    textBox8.Text = "200-0810000001679-53";
                    textBox12.Text = "204-1895-";
                }
                else
                {
                    richTextBox6.Text = "Uplata premije za osiguranje od autoodgovornosti po polisi\nbr. ";
                    richTextBox5.Text = "GFO NOVI PAZAR\nŠabana Koče 71\nNovi Pazar";
                    textBox27.Text = "160";
                    textBox26.Text = "RSD";
                    textBox25.Text = "";
                    textBox23.Text = "200-0810000001679-53";
                    textBox19.Text = "204-1895-";
                }
            }
            if (comboBox1.SelectedIndex == 8)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Porez na upotrebu m/v, " + txtVBrojSasije.Text;
                    richTextBox3.Text = "POREZ NA UPOTREBU M/V\nSjenica\nSjenica";
                    textBox18.Text = "253";
                    textBox17.Text = "RSD";
                    textBox16.Text = "";
                    textBox8.Text = "840000071451184387";
                    textBox9.Text = "97";
                    textBox12.Text = "19-091";
                }
                else
                {
                    richTextBox6.Text = "Porez na upotrebu m/v, " + txtVBrojSasije.Text;
                    richTextBox5.Text = "POREZ NA UPOTREBU M/V\nSjenica\nSjenica";
                    textBox27.Text = "153";
                    textBox26.Text = "RSD";
                    textBox25.Text = "";
                    textBox21.Text = "97";
                    textBox23.Text = "840000071451184387";
                    textBox19.Text = "19-091";
                }
            }
            if (comboBox1.SelectedIndex == 10)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Uplata po zelenoj karti broj ";
                    richTextBox3.Text = "GFO NOVI PAZAR\nŠabana Koče 71\nNovi Pazar";
                    textBox18.Text = "221";
                    textBox17.Text = "RSD";
                    textBox16.Text = "2000";
                    textBox8.Text = "360-1912-84";
                    textBox12.Text = "1110-1895-";
                }
                else
                {
                    richTextBox6.Text = "Uplata po zelenoj karti broj ";
                    richTextBox5.Text = "GFO NOVI PAZAR\nŠabana Koče 71\nNovi Pazar";
                    textBox27.Text = "160";
                    textBox26.Text = "RSD";
                    textBox25.Text = "2000";
                    textBox23.Text = "360-1912-84";
                    textBox19.Text = "1110-1895-";
                }
            }
            if (comboBox1.SelectedIndex == 11)
            {
                if (checkBox2.Checked)
                {
                    richTextBox2.Text = "Tehnički pregled vozila ";
                    richTextBox3.Text = "VAAS DOO SJENICA \nUl. Nikole Tesle 1\n36310 Sjenica";
                    textBox18.Text = "221";
                    textBox17.Text = "RSD";
                    textBox16.Text = "5000";
                    textBox8.Text = "200-3410880101944-68";

                }
                else
                {
                    richTextBox6.Text = "Tehnički pregled vozila ";
                    richTextBox5.Text = "VAAS DOO SJENICA \nUl. Nikole Tesle 1\n36310 Sjenica";
                    textBox27.Text = "153";
                    textBox26.Text = "RSD";
                    textBox25.Text = "5000";
                    textBox23.Text = "200-3410880101944-68";
                }
            }
        }

        //pretraga musterija po prezimenu
        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            if (txtImePrezimePretraga.Text != "")
            {
                cbMusterije.DataSource = null;
                cbMusterije.DataSource = b.ucitajMusterije(txtImePrezimePretraga.Text);
            }
            else
                cbMusterije.DataSource = null;

        }

        //ispisivanje podataka za odabranog musteriju
        private void cbMusterije_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (txtImePrezimePretraga.Text != "")
            vozilo2 v = cbMusterije.SelectedItem as vozilo2;
            if (v != null)
            {

                txtTPNap.Text = "";
                txtVezaID.Text = "";
                txtVezaGod.Text = "";
                txtTPID.Text = v.IDBroj;
                txtRegBr.Text = v.RegistarskiBroj;
                cbVrstaVoz.Text = v.VrstaVozila;
                cbVPV.Text = v.PodvrstaVozila;
                cbVKlasa.Text = v.klasa;
                cbVOblikKar.Text = v.oblikkaroserije;
                cbUredjajZaZaust.Text = v.zaustavljanje;
                cbVOKS.Text = v.OblikKaroserijeSkraceno;
                cbVSpec.Text = v.specijalno;
                if (cbVSpec.Text != "")
                    checkBox12.Checked = true;
                else
                    checkBox12.Checked = false;
                cbVMarka.Text = v.Marka;
                txtVModel.Text = v.Model;
                txtVTip.Text = v.Tip;
                txtVGodinaProizvodnje.Text = v.GodinaProizvodnje;
                txtVBrojSasije.Text = v.BrojSasije;
                txtVBrojMotora.Text = v.BrojMotora;
                cbVBoja.Text = v.Boja;
                txtVPrvaReg.Text = v.DatumPrveRegostracije;
                txtVOdnos.Text = v.Odnos;
                txtVSnaga.Text = v.Snaga;
                txtVZapremina.Text = v.Zapremina;
                txtVNajMasa.Text = v.NajvecaMasa;
                txtVTezina.Text = v.Tezina;
                txtVNosivost.Text = v.Nosivost;
                txtVSedenje1.Text = v.Sedenje;
                txtVStajanje.Text = v.Stajanje;
                txtVBrojOsovina.Text = v.BrojOsovina;
                cbVGorivo.Text = v.Gorivo;
                txtVHomologacija.Text = v.Homologacija;
                txtVIme.Text = v.ImeVL;
                txtVPrezime.Text = v.PrezimeVL;
                txtVJMBG.Text = v.MaticniBrojVL;
                txtVBRLK.Text = v.BRLKVL;
                txtVBrPoste.Text = v.BrojPosteVL;
                txtVBrDokumenta.Text = v.BrojDokumenta;
                txtVSerijskiBroj.Text = v.SerijskiBroj;
                cbVMesto.Text = v.MestoVL;
                txtVUlica.Text = v.UlicaVL;
                txtVBroj.Text = v.BrojVL;
                txtVBrTelefona.Text = v.TelefonaVL;
                txtVZiroRacun.Text = v.ZiroRacunVL;
                txtDLIme.Text = v.ImeDL;
                txtDLPrezime.Text = v.PrezimeDL;
                txtDLJMBG.Text = v.MaticniBrojDL;
                txtDLBRLK.Text = v.BRLKDL;
                txtDLBrPoste.Text = v.BrojPosteDL;
                cbDLMesto.Text = v.MestoDL;
                txtDLUlica.Text = v.UlicaDL;
                txtDLBroj.Text = v.BrojDL;
                txtDLBrTelefona.Text = v.TelefonaDL;
                txtTPNap.Text = v.Napomena;
                txtTPDatum.Text = v.Datum;
                txtTPVremeOd.CustomFormat = "HH:mm";
                txtTPVremeDo.CustomFormat = "HH:mm";
                txtTPVremeOd.Text = v.VremeOd;
                txtTPVremeDo.Text = v.VremeDo;
                txtVD.Text = v.duzina;
                txtVS.Text = v.sirina;
                txtVV.Text = v.visina;
                cbTPKontr1.Text = v.Kontrolor1;
                cbTPKontr2.Text = v.Kontrolor2;
                txtVKM.Text = v.Predjenikm;
                txtPoslednjeUverenjeBroj.Text = v.PoslednjeUverenjeBroj;
                txtPoslednjeUverenjeDatum.Text = v.PoslednjeUverenjeDatum;
                cbDecOzn.Text = v.DecimalnaOznaka;
                cbBojaL.Text = v.BojaL;
                cbBojaD.Text = v.BojaD;

                txtVBrSati.Text = v.BrSati;
                cbVKuka.Text = v.Kuka;
                txtVOpterecenjeN1.Text = v.OpterecenjeN1; txtVOpterecenjeN2.Text = v.OpterecenjeN2;
                txtVOpterecenjeN3.Text = v.OpterecenjeN3;
                txtVOpterecenjeN4.Text = v.OpterecenjeN4;
                txtVOpterecenjeN5.Text = v.OpterecenjeN5;
                txtVPneumatikL1.Text = v.PneumatikL1;
                txtVPneumatikL2.Text = v.PneumatikL2;
                txtVPneumatikL3.Text = v.PneumatikL3;
                txtVPneumatikL4.Text = v.PneumatikL4;
                txtVPneumatikL5.Text = v.PneumatikL5;
                txtVPneumatikD1.Text = v.PneumatikD1;
                txtVPneumatikD2.Text = v.PneumatikD2;
                txtVPneumatikD3.Text = v.PneumatikD3;
                txtVPneumatikD4.Text = v.PneumatikD4;
                txtVPneumatikD5.Text = v.PneumatikD5;
                cbUredjajZaZaust.Text = v.zaustavljanje;
                txtVKomentar.Text = v.komentar;
                txtVrstaPlusDecOznk.Text = v.decplusoznaka;
                txtVVlaznostKT.Text = v.VlaznostKT;
                datumIsteka.Text = v.datumIsteka;
                txtDLZiroRacun.Text = v.DLZiroRacun;
                tbDatumIzdavanja.Text= v.DatumIzdavanjaSaobracajne;
                if (txtDLIme.Text != "")
                    rbDL.Checked = true;
                else
                    rbVL.Checked = true;
                radioButton2_CheckedChanged(sender, e);
                int a = 1;
                try
                {
                    a = int.Parse(v.rejting);
                }
                catch
                {
                    a = 1;
                }
                txtVRejting.Value = a;

                if (v.Ispravnost == "Da")
                {
                    rbTehIsp.Checked = true;
                    ispravnost = "Da";
                }
                else if (v.Ispravnost == "Ne")
                {
                    rbTehNeis.Checked = true;
                    ispravnost = "Ne";
                }
                else if (v.Ispravnost == "Prekid")
                {
                    rbTehIspPrekid.Checked = true;
                    ispravnost = "Prekid";
                }
                if (v.terensko == "Da")
                {
                    cbTerensko.Checked = true;
                    terensko = "Da";
                }
                else if (v.terensko == "Ne")
                {
                    cbTerensko.Checked = false;
                    terensko = "Ne";
                }
                popuniTabeluNeispravnosti(txtTPID.Text);
                cbVPV.Text = v.PodvrstaVozila;
                btnIzv.Visible = true;
                btnPot2.Visible = true;
                btnReg2.Visible = true;
                btnZap1.Visible = true;
                txtDLZiroRacun.Text = v.DLZiroRacun;
                cbTablice.Text = "";
                if (v.ZiroRacunVL != "")
                    checkBox2.Checked = true;
                else
                    checkBox4.Checked = true;
                if (v.DLZiroRacun != "")
                    cbDLPravno.Checked = true;
                else
                    cbDLFizicko.Checked = true;
                popuniFakture();
                string datum = txtTPDatum.Text;
                proveriSlike(datum);
                oznaciVrstuPregleda(txtTPID.Text[8]);
                txtTPID.Text = v.IDBroj;
            }
        }

        //ispisivanje podataka za odabrane tablice
        private void cbTablice_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (prvoUcitavanje == false && /*((ComboBox)sender).SelectedItem !=null*/ cbTabliceEnter == true)
            {
                txtTPNap.Text = "";
                txtVezaID.Text = "";
                txtVezaGod.Text = "";

                vozilo v = cbTablice.SelectedItem as vozilo;

                txtTPID.Text = v.IDBroj;
                txtRegBr.Text = v.RegistarskiBroj;
                cbVrstaVoz.Text = v.VrstaVozila;
                cbVPV.Text = v.PodvrstaVozila;
                cbVKlasa.Text = v.klasa;
                cbVOblikKar.Text = v.oblikkaroserije;
                cbUredjajZaZaust.Text = v.zaustavljanje;
                cbVOKS.Text = v.OblikKaroserijeSkraceno;
                cbVSpec.Text = v.specijalno;
                if (cbVSpec.Text != "")
                    checkBox12.Checked = true;
                else
                    checkBox12.Checked = false;
                cbVMarka.Text = v.Marka;
                txtVModel.Text = v.Model;
                txtVTip.Text = v.Tip;
                txtVGodinaProizvodnje.Text = v.GodinaProizvodnje;
                txtVBrojSasije.Text = v.BrojSasije;
                txtVBrojMotora.Text = v.BrojMotora;
                cbVBoja.Text = v.Boja;
                txtVPrvaReg.Text = v.DatumPrveRegostracije;
                tbDatumIzdavanja.Text = v.DatumIzdavanjaSaobracajne;
                txtVOdnos.Text = v.Odnos;
                txtVSnaga.Text = v.Snaga;
                txtVZapremina.Text = v.Zapremina;
                txtVNajMasa.Text = v.NajvecaMasa;
                txtVTezina.Text = v.Tezina;
                txtVNosivost.Text = v.Nosivost;
                txtVSedenje1.Text = v.Sedenje;
                txtVStajanje.Text = v.Stajanje;
                txtVBrojOsovina.Text = v.BrojOsovina;
                cbVGorivo.Text = v.Gorivo;
                txtVHomologacija.Text = v.Homologacija;
                txtVIme.Text = v.ImeVL;
                txtVPrezime.Text = v.PrezimeVL;
                txtVJMBG.Text = v.MaticniBrojVL;
                txtVBRLK.Text = v.BRLKVL;
                txtVBrPoste.Text = v.BrojPosteVL;
                cbVMesto.Text = v.MestoVL;
                txtVUlica.Text = v.UlicaVL;
                txtVBroj.Text = v.BrojVL;
                txtVBrTelefona.Text = v.TelefonaVL;
                txtVZiroRacun.Text = v.ZiroRacunVL;
                txtDLIme.Text = v.ImeDL;
                txtDLPrezime.Text = v.PrezimeDL;
                txtDLJMBG.Text = v.MaticniBrojDL;
                txtDLBRLK.Text = v.BRLKDL;
                txtDLBrPoste.Text = v.BrojPosteDL;
                cbDLMesto.Text = v.MestoDL;
                txtDLUlica.Text = v.UlicaDL;
                txtDLBroj.Text = v.BrojDL;
                txtDLBrTelefona.Text = v.TelefonaDL;
                txtTPNap.Text = v.Napomena;
                txtTPDatum.Text = v.Datum;
                txtTPVremeOd.CustomFormat = "HH:mm";
                txtTPVremeDo.CustomFormat = "HH:mm";
                txtTPVremeDo.Format = DateTimePickerFormat.Custom;
                txtTPVremeOd.Format = DateTimePickerFormat.Custom;
                txtTPVremeOd.Text = v.VremeOd;
                txtTPVremeDo.Text = v.VremeDo;
                cbTPKontr1.Text = v.Kontrolor1;
                cbTPKontr2.Text = v.Kontrolor2;
                txtVKM.Text = v.Predjenikm;
                txtPoslednjeUverenjeBroj.Text = v.PoslednjeUverenjeBroj;
                txtPoslednjeUverenjeDatum.Text = v.PoslednjeUverenjeDatum;
                cbDecOzn.Text = v.DecimalnaOznaka;
                cbBojaL.Text = v.BojaL;
                cbBojaD.Text = v.BojaD;
                txtVD.Text = v.duzina;
                txtVS.Text = v.sirina;
                txtVV.Text = v.visina;
                txtVBrDokumenta.Text = v.BrojDokumenta;
                txtVSerijskiBroj.Text = v.SerijskiBroj;
                txtVBrSati.Text = v.BrSati;
                cbVKuka.Text = v.Kuka;
                txtVOpterecenjeN1.Text = v.OpterecenjeN1; txtVOpterecenjeN2.Text = v.OpterecenjeN2;
                txtVOpterecenjeN3.Text = v.OpterecenjeN3;
                txtVOpterecenjeN4.Text = v.OpterecenjeN4;
                txtVOpterecenjeN5.Text = v.OpterecenjeN5;
                txtVPneumatikL1.Text = v.PneumatikL1;
                txtVPneumatikL2.Text = v.PneumatikL2;
                txtVPneumatikL3.Text = v.PneumatikL3;
                txtVPneumatikL4.Text = v.PneumatikL4;
                txtVPneumatikL5.Text = v.PneumatikL5;
                txtVPneumatikD1.Text = v.PneumatikD1;
                txtVPneumatikD2.Text = v.PneumatikD2;
                txtVPneumatikD3.Text = v.PneumatikD3;
                txtVPneumatikD4.Text = v.PneumatikD4;
                txtVPneumatikD5.Text = v.PneumatikD5;
                cbUredjajZaZaust.Text = v.zaustavljanje;
                txtVKomentar.Text = v.komentar;
                txtVrstaPlusDecOznk.Text = v.decplusoznaka;
                txtVVlaznostKT.Text = v.VlaznostKT;
                datumIsteka.Text = v.datumIsteka;
                txtDLZiroRacun.Text = v.DLZiroRacun;
                tbDatumIzdavanja.Text = v.DatumIzdavanjaSaobracajne;
                //if (txtDLIme.Text != "")
                //    rbDL.Checked = true;
                //else
                //    rbVL.Checked = true;
                if (v.Prioritet == "DRUGO LICE")
                    rbDL.Checked = true;
                else
                    rbVL.Checked = true;
                radioButton2_CheckedChanged(sender, e);
                int a = 1;
                try
                {
                    a = int.Parse(v.rejting);
                }
                catch
                {
                    a = 1;
                }
                txtVRejting.Value = a;

                if (v.Ispravnost == "Da")
                {
                    rbTehIsp.Checked = true;
                    ispravnost = "Da";
                }
                else if (v.Ispravnost == "Ne")
                {
                    rbTehNeis.Checked = true;
                    ispravnost = "Ne";
                }
                else if (v.Ispravnost == "Prekid")
                {
                    rbTehIspPrekid.Checked = true;
                    ispravnost = "Prekid";
                }
                if (v.terensko == "Da")
                {
                    cbTerensko.Checked = true;
                    terensko = "Da";
                }
                else if (v.terensko == "Ne")
                {
                    cbTerensko.Checked = false;
                    terensko = "Ne";
                }
                popuniTabeluNeispravnosti(txtTPID.Text);
                cbVPV.Text = v.PodvrstaVozila;
                btnIzv.Visible = true;
                btnPot2.Visible = true;
                btnReg2.Visible = true;
                btnZap1.Visible = true;
                txtImePrezimePretraga.Text = "";
                cbOsiguranje.SelectedItem = v.Osiguranje;
                if (v.ZiroRacunVL != "")
                    checkBox2.Checked = true;
                else
                    checkBox4.Checked = true;
                if (v.DLZiroRacun != "")
                    cbDLPravno.Checked = true;
                else
                    cbDLFizicko.Checked = true;
                string datum = txtTPDatum.Text;
                proveriSlike(datum);
                popuniFakture();
                oznaciVrstuPregleda(txtTPID.Text[8]);
                txtTPID.Text = v.IDBroj;
                cbSasije.SelectedIndex = -1;
                cbSasije.Text = "";
                cbIdjevi.SelectedIndex = -1;
                cbIdjevi.Text = "";
            }
        }

        private void oznaciVrstuPregleda(char v)
        {
            int index = int.Parse(v.ToString());
            comboBox2.SelectedIndex = (index - 1);
        }

        int iTotalFields = 0;

        //kreiranje id broja promenom na cbu
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    txtTPID.Text = "";
            //    string id = "189046";
            //    id += (comboBox2.SelectedIndex + 1).ToString();
            //    string poslednji = b.nadjiPoslednjiPregled();
            //    int sledeci = int.Parse(poslednji);
            //    sledeci++;
            //    id += sledeci;
            //    txtTPID.Text = id;
            //}
            try
            {
                string sifra = "141022";
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 6; i++)
                    sb.Append(sifra[i]);
                sifra = (DateTime.Now.Year % 100).ToString();
                for (int i = 0; i < 2; i++)
                    sb.Append(sifra[i]);
                sb.Append(comboBox2.SelectedIndex + 1);
                string poslednji = b.nadjiPoslednjiPregled();
                int sledeci = int.Parse(poslednji);
                sledeci++;
                poslednji = sledeci.ToString().PadLeft(5, '0');
                Brpregleda = poslednji;
                for (int i = 0; i < 5; i++)
                    sb.Append(poslednji[i]);
                txtTPID.Text = sb.ToString();
            }
            catch { }
        }


        //racunanje predjenih kilometara za uneseni broj sasije
        private void txtVKM_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtVBrojSasije.Text != "")
                {
                    try
                    {
                        vozilo v = b.reg(txtVBrojSasije.Text.Trim());
                        txtRazlika.Text = (float.Parse(txtVKM.Text) - float.Parse(v.Predjenikm)).ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
                else
                    MessageBox.Show("Unesite broj šasije!");
            }
        }

        //popunjavanje podataka nakon odabranog oblika karoserije iz cb
        private void cbVOblikKar_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbVOKS.Text = "";

            List<string> lista = new List<string>();
            foreach (var k in b.lista_kar_sif(cbVOblikKar.Text))
            {
                lista.Add(k.KaroserijaNaziv);
            }
            cbVOKS.DataSource = lista;
            //cbVOKS.Text = "";


            if (cbVrstaVoz.Text == "TERETNO VOZILO" || cbVrstaVoz.Text == "PRIKLJUCNO VOZILO")
            {

                cbDecOzn.Enabled = true;
                lblDecOzn.Enabled = true;
            }
            else
            {

                cbDecOzn.Enabled = false;
                lblDecOzn.Enabled = false;
                cbDecOzn.Text = "";

            }

        }

        //popunjavanje podataka za listu kvarova nakon odabranog dijela iz cb
        private void cbNaziv_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<neispravnost> lista = new List<neispravnost>();
            foreach (var k in b.lista_kvarova(cbNaziv.Text))
            {
                lista.Add(k);
            }
            cbNeispravnost.DataSource = lista;
            cbNeispravnost.Text = "";
        }

        private void UkljuciIskljuciNeispravnosti()
        {
            if (ispravnost == "Da")
            {
                dataGridView1.Enabled = false;
                cbGlPodela.Enabled = false;
                cbNaziv.Enabled = false;
                cbNeispravnost.Enabled = false;
                bunifuThinButton213.Enabled = false;
            }
            else
            {
                dataGridView1.Enabled = true;
                cbGlPodela.Enabled = true;
                cbNaziv.Enabled = true;
                cbNeispravnost.Enabled = true;
                bunifuThinButton213.Enabled = true;
            }
        }

        //dodavanje neispravnosti u datagrid
        private void bunifuThinButton213_Click_1(object sender, EventArgs e)
        {
            if (rbTehNeis.Checked || rbTehIspPrekid.Checked)
            {
                neispravnost n = cbNeispravnost.SelectedItem as neispravnost;
                dataGridView1.Rows.Add(n.ID, cbNaziv.Text, cbNeispravnost.Text, "");
                brNeispravnost = dataGridView1.Rows.Count - 1;
            }
            else MessageBox.Show("Tehnička ispravnost je obeležena sa DA, nije moguće dodati neispravnost!");
        }

        //popunjavanje decimalne oznake iz cb
        private void cbVOKS_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbVrstaVoz.Text == "TERETNO VOZILO" || cbVrstaVoz.Text == "PRIKLJUCNO VOZILO")
            {
                List<string> listadec = new List<string>();
                foreach (var k in b.lista_dec(cbVOKS.Text))
                {
                    listadec.Add(k.KaroserijaNaziv);
                }
                cbDecOzn.DataSource = listadec;
                cbDecOzn.Text = "";

                cbDecOzn.Enabled = true;
                lblDecOzn.Enabled = true;
                string s;
                if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
                    s = cbVKlasa.Text + "/" + cbVOKS.Text;
                else
                    if (cbDecOzn.Text != "")
                    if (checkBox12.Checked)
                        s = (cbVSpec.Text.Split(' '))[0].ToString() + (cbDecOzn.Text.Split(' '))[0].ToString();
                    else
                        s = cbVOKS.Text + (cbDecOzn.Text.Split(' '))[0].ToString();
                else
                    s = cbVOKS.Text;
                txtVrstaPlusDecOznk.Text = s;
            }
            else if (cbVrstaVoz.Text != "AUTOBUS" && cbVrstaVoz.Text != "TERETNO VOZILO" && cbVrstaVoz.Text != "PRIKLJUCNO VOZILO")
                txtVrstaPlusDecOznk.Text = cbVOKS.Text;
            else
            {

                cbDecOzn.Enabled = false;
                lblDecOzn.Enabled = false;


            }
        }


        //private void txtPoslednjeUverenjeDatum_KeyDown(object sender, KeyEventArgs e)
        //{
        //    //txtPoslednjeUverenjeDatum.MaxLength = 10;
        //    //if (e.KeyCode!= Keys.Back)
        //    //{

        //    //    if (txtPoslednjeUverenjeDatum.Text.Length == 2)
        //    //        txtPoslednjeUverenjeDatum.Text += ".";
        //    //    if (txtPoslednjeUverenjeDatum.Text.Length == 5)
        //    //        txtPoslednjeUverenjeDatum.Text += ".";
        //    //    txtPoslednjeUverenjeDatum.SelectionStart = txtPoslednjeUverenjeDatum.Text.Length;
        //    //}
        //}


        //private void txtTPVremeOd_KeyDown_1(object sender, KeyEventArgs e)
        //{
        //    txtTPVremeOd.MaxLength = 5;
        //    if (e.KeyCode != Keys.Back)
        //    {

        //        if (txtTPVremeOd.Text.Length == 2)
        //            txtTPVremeOd.Text += ":";
        //        txtTPVremeOd.SelectionStart = txtTPVremeOd.Text.Length;


        //    }
        //}

        //racunanje vremena provedenog na pregledu
        private void maskedTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DateTime vreme = Convert.ToDateTime(txtTPVremeOd.Text);
                int najmasa;
                try
                {
                    najmasa = int.Parse(txtVNajMasa.Text);
                }
                catch
                {
                    najmasa = 0;
                }

                txtTPVremeDo.CustomFormat = "HH:mm";
                if (cbVrstaVoz.Text == "MOPED" || cbVrstaVoz.Text == "MOTOCIKL" || cbVrstaVoz.Text == "LAKI TRICIKL" || cbVrstaVoz.Text == "MOTOCIKL SA BOCNIM SEDISTEM"
                    || cbVrstaVoz.Text == "TESKI TRICIKL" || cbVrstaVoz.Text == "LAKI CETVOROCIKL" || cbVrstaVoz.Text == "TESKI CETVOROCIKL" || cbVrstaVoz.Text == "BICIKL SA MOTOROM"
                     || cbVrstaVoz.Text == "RADNA MASINA" || cbVrstaVoz.Text == "MOTOKULTIVATOR" || cbVrstaVoz.Text == "PRIKLJUCNO ZA TRAKTOR")
                {
                    TimeSpan p = new TimeSpan(0, 15, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;

                }
                else if (cbVrstaVoz.Text == "TRAKTOR")
                {
                    TimeSpan p = new TimeSpan(0, 20, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;
                }
                else if (cbVrstaVoz.Text == "PUTNICKO VOZILO")
                {
                    TimeSpan p = new TimeSpan(0, 25, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;
                }
                else if (cbVrstaVoz.Text == "TERETNO VOZILO" && cbUredjajZaZaust.Text == "Hidraulični" && najmasa <= 3500)
                {
                    TimeSpan p = new TimeSpan(0, 25, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;
                }
                else if (cbVrstaVoz.Text == "TERETNO VOZILO" && cbUredjajZaZaust.Text == "Hidraulični" && najmasa > 3500)
                {
                    TimeSpan p = new TimeSpan(0, 30, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;
                }
                else if (cbVrstaVoz.Text == "TERETNO VOZILO" && (cbUredjajZaZaust.Text == "Pneumatski" || cbUredjajZaZaust.Text == "Hidropneumatski"))
                {
                    TimeSpan p = new TimeSpan(0, 40, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;
                }
                else if (cbVrstaVoz.Text == "AUTOBUS" && cbUredjajZaZaust.Text == "Hidraulični" && najmasa <= 3500)
                {
                    TimeSpan p = new TimeSpan(0, 25, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;
                }
                else if (cbVrstaVoz.Text == "AUTOBUS" && cbUredjajZaZaust.Text == "Hidraulični" && najmasa > 3500)
                {
                    TimeSpan p = new TimeSpan(0, 30, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;
                }
                else if (cbVrstaVoz.Text == "AUTOBUS" && (cbUredjajZaZaust.Text == "Pneumatski" || cbUredjajZaZaust.Text == "Hidropneumatski"))
                {
                    TimeSpan p = new TimeSpan(0, 40, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;
                }
                else if (cbVrstaVoz.Text == "PRIKLJUCNO VOZILO" && (cbUredjajZaZaust.Text == "Inerciona komanda" || cbUredjajZaZaust.Text == "Bez uređaja za zaustavljanje" || cbUredjajZaZaust.Text == "Ostalo"))
                {
                    TimeSpan p = new TimeSpan(0, 15, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;
                }
                else
                {
                    TimeSpan p = new TimeSpan(0, 40, 0);
                    vreme += p;
                    txtTPVremeDo.Text = vreme.ToString("HH:mm");
                    duzinaTehPregleda = p.Minutes;
                }
            }
        }

        //postavljanje Da kod ispravnosti vozila
        private void rbTehIsp_CheckedChanged(object sender, EventArgs e)
        {
            if (rbTehIsp.Checked)
            {
                //if (ispravnost == "Ne" && txtTPDatum.Text.Substring(6, 4) == DateTime.Now.Year.ToString())
                //{
                //    txtTPNap.Text = txtTPDatum.Text + "\r\n" + txtTPID.Text;
                //    txtVezaID.Text = txtTPID.Text;
                //    txtVezaGod.Text = txtTPDatum.Text;
                //}
                ispravnost = "Da";
                UkljuciIskljuciNeispravnosti();
                dataGridView1.Rows.Clear();
            }

        }

        //postavljanje Ne kod ispravnosti vozila
        private void rbTehNeis_CheckedChanged(object sender, EventArgs e)
        {
            if (rbTehNeis.Checked)
            {
                ispravnost = "Ne";
                UkljuciIskljuciNeispravnosti();
            }
        }

        //brisanje podataka iz datagrida na izvestaju klikom na rowheder 
        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                dataGridView1.Rows.RemoveAt(e.RowIndex);
                brNeispravnost = dataGridView1.Rows.Count - 1;
            }
            catch (Exception)
            {

            }
        }

        //brisanje podataka iz datagrida na fakturi klikom na rowheder
        private void dataGridView2_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                dataGridView2.Rows.RemoveAt(dataGridView2.CurrentCell.RowIndex);
                i--;
                int k = 0;
                double u = 0;
                while (k < dataGridView2.Rows.Count - 1)
                {
                    u += Convert.ToDouble(dataGridView2.Rows[k].Cells[11].FormattedValue);
                    k++;
                }
                textBox5.Text = u.ToString();
            }
            catch (Exception)
            {


            }
        }

        //Otvaranje slika za dati ID broj
        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {
            int n;

            if (txtTPID.Text.Length < 14 || txtTPID.Text.Length > 14)
            {
                MessageBox.Show("ID Broj je nepravilan!");
            }
            else if (int.TryParse(txtTPID.Text.Substring(9), out n) == false)
            {
                MessageBox.Show("ID Broj je nepravilan!");
            }
            else
            {
                try
                {
                    string datum = txtTPDatum.Text;
                    Process.Start(@"D:\Slike\" + datum.Substring(6) + @"\" + datum.Substring(0, 2) + ". " + nadjiMesec(datum.Substring(3, 2)) + @"\" + int.Parse(txtTPID.Text.Substring(9)));
                }
                catch (Exception)
                {
                    MessageBox.Show("Ne postoji folder za dati ID broj!");
                }
            }


        }

        private string nadjiMesec(string v)
        {
            switch (v)
            {
                case "01": return "Januar";
                case "02": return "Februar";
                case "03": return "Mart";
                case "04": return "April";
                case "05": return "Maj";
                case "06": return "Jun";
                case "07": return "Jul";
                case "08": return "Avgust";
                case "09": return "Septembar";
                case "10": return "Oktobar";
                case "11": return "Novembar";
                default: return "Decembar";

            }
        }

        //Onemogucuje promenu indexa na tablicama kada se skine fokus sa njih 
        private void cbTablice_Leave(object sender, EventArgs e)
        {
            cbTabliceEnter = false;
        }

        //Omogucuje promenu indexa na tablicama kada se fokus stavi na njih
        private void cbTablice_Enter(object sender, EventArgs e)
        {
            cbTabliceEnter = true;
        }

        //Postavljanje vrijednosti vremena OD kad se fokus stavi na njega
        private void txtTPVremeOd_Enter(object sender, EventArgs e)
        {
            txtTPVremeOd.CustomFormat = "HH:mm";
        }

        //Postavljanje vrijednosti vremena DO kad se fokus stavi na njega
        private void txtTPVremeDo_Enter(object sender, EventArgs e)
        {
            txtTPVremeDo.CustomFormat = "HH:mm";

        }



        private void txtVBrojSasije_TextChanged(object sender, EventArgs e)
        {
            lblDuzinaSasija.Text = txtVBrojSasije.Text.Length.ToString();
            if (File.Exists(@"D:\Slike\Uverenja\" + (txtVBrojSasije.Text).Replace("/", "") + ".pdf"))
                btbUverenje.Visible = true;
            else
                btbUverenje.Visible = false;
        }

        private void label90_Click(object sender, EventArgs e)
        {

        }

        private void cbTerensko_CheckedChanged(object sender, EventArgs e)
        {
            if (cbTerensko.Checked)
                terensko = "Da";
            else
                terensko = "Ne";
        }

        private void bunifuThinButton218_Click(object sender, EventArgs e)
        {

        }

        private void bunifuThinButton217_Click(object sender, EventArgs e)
        {

        }

        private void bunifuThinButton216_Click(object sender, EventArgs e)
        {

        }

        private void bunifuThinButton215_Click(object sender, EventArgs e)
        {

        }

        private void bunifuThinButton217_Click_1(object sender, EventArgs e)
        {
            try
            {
                var data = ElectronicIdReader.ReadAll();
                //    Console.WriteLine("Fiksni podaci ------------------------------------------");

                txtJMBG.Text = data.FixedPersonalData.PersonalNumber;
                txtPILICA.Text = data.FixedPersonalData.Surname + " " + data.FixedPersonalData.GivenName;
                //txtULiBR.Text = d1.placeOfBirth;
                //txtMestoRodjenja.Text += "," + d1.communityOfBirth;
                //txtGRAD.Text += ", " + d1.stateOfBirth;


                //    Console.WriteLine("Variablini podaci ------------------------------------------");

                //txtPrebivaliste.Text= d2.state;
                //txtPrebivaliste.Text = d2.community;
                txtGRAD.Text = data.VariablePersonalData.Place + ", " + data.VariablePersonalData.Community;
                txtULiBR.Text = data.VariablePersonalData.Street + " " + data.VariablePersonalData.HouseNumber + (string.IsNullOrWhiteSpace(data.VariablePersonalData.Floor)
        ? ""
        : "/" + data.VariablePersonalData.Floor) +
    (string.IsNullOrWhiteSpace(data.VariablePersonalData.ApartmentNumber)
        ? ""
        : "/" + data.VariablePersonalData.ApartmentNumber); ;
                //Console.WriteLine("houseLetter: {0}", d2.houseLetter);
                //Console.WriteLine("entrance: {0}", d2.entrance);
                //Console.WriteLine("floor: {0}", d2.floor);
                //Console.WriteLine("apartmentNumber: {0}", d2.apartmentNumber);

                //    Console.WriteLine("Slika ------------------------------------------");
                //Bitmap c = x.ReadProfileImage();
                //pictureBox2.Image = c;

                //Console.ReadLine();
            }
            catch (Exception ex)
            {

                MessageBox.Show("Molim vas ubacite ličnu kartu");
            }
        }

        private void bunifuThinButton218_Click_1(object sender, EventArgs e)
        {
            try
            {
                var data = ElectronicIdReader.ReadAll();
                txtLicnaKartaK.Text = data.Document.RegistrationNumber;

                //    Console.WriteLine("Fiksni podaci ------------------------------------------");

                txtJMBGK.Text = data.FixedPersonalData.PersonalNumber;
                txtPIK.Text = data.FixedPersonalData.Surname + " " + data.FixedPersonalData.GivenName;
                //txtULiBR.Text = d1.placeOfBirth;
                //txtMestoRodjenja.Text += "," + d1.communityOfBirth;
                //txtGRAD.Text += ", " + d1.stateOfBirth;


                //    Console.WriteLine("Variablini podaci ------------------------------------------");

                //txtPrebivaliste.Text= d2.state;
                //txtPrebivaliste.Text = d2.community;
                txtGradK.Text = data.VariablePersonalData.Place + ", " + data.VariablePersonalData.Community;
                txtULiBK.Text = data.VariablePersonalData.Street + " " + data.VariablePersonalData.HouseNumber + (string.IsNullOrWhiteSpace(data.VariablePersonalData.Floor)
        ? ""
        : "/" + data.VariablePersonalData.Floor) +
    (string.IsNullOrWhiteSpace(data.VariablePersonalData.ApartmentNumber)
        ? ""
        : "/" + data.VariablePersonalData.ApartmentNumber); ;
                //Console.WriteLine("houseLetter: {0}", d2.houseLetter);
                //Console.WriteLine("entrance: {0}", d2.entrance);
                //Console.WriteLine("floor: {0}", d2.floor);
                //Console.WriteLine("apartmentNumber: {0}", d2.apartmentNumber);

                //    Console.WriteLine("Slika ------------------------------------------");
                //Bitmap c = x.ReadProfileImage();
                //pictureBox2.Image = c;

                //Console.ReadLine();
            }
            catch (Exception ex)
            {

                MessageBox.Show("Molim vas ubacite ličnu kartu");

            }
        }

        private void bunifuThinButton210_Click_1(object sender, EventArgs e)
        {
            Object oMissing = System.Reflection.Missing.Value;

            Object oTrue = true;
            Object oFalse = false;

            Word.Application oWord = new Word.Application();
            Word.Document oWordDoc = new Word.Document();

            oWord.Visible = true;
            Object oTemplatePath;
            long uk = 0;
            oTemplatePath = System.Windows.Forms.Application.StartupPath + "\\UgovorKP.dot";

            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            foreach (Word.Field myMergeField in oWordDoc.Fields)
            {

                iTotalFields++;
                Word.Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;

                if (fieldText.StartsWith(" MERGEFIELD"))
                {

                    Int32 endMerge = fieldText.IndexOf("\\");
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(11, endMerge - 11);

                    fieldName = fieldName.Trim();
                    if (fieldName == "VRSTA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVrstaVoz.SelectedItem as string);

                    }
                    if (fieldName == "SASIJA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBrojSasije.Text);
                    }
                    if (fieldName == "GODINA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVGodinaProizvodnje.Text);
                    }
                    if (fieldName == "SNAGA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVSnaga.Text);
                    }

                    if (fieldName == "BRMOTORA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBrojMotora.Text);
                    }

                    if (fieldName == "MARKA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVMarka.Text);
                    }

                    if (fieldName == "MODEL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVModel.Text);
                    }

                    if (fieldName == "RZ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVZapremina.Text);
                    }

                    if (fieldName == "REG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtRegBr.Text);
                    }
                    if (fieldName == "DATUM")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(DateTime.Now.ToString("d.M.yyyy"));
                    }
                    if (fieldName == "PI")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVPrezime.Text + " " + txtVIme.Text);
                    }
                    if (fieldName == "ULICA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVUlica.Text);
                    }
                    if (fieldName == "BROJ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBroj.Text);
                    }
                    if (fieldName == "GRAD")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVMesto.Text);
                    }
                    if (fieldName == "JMBG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVJMBG.Text);
                    }
                    if (fieldName == "LICNA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBRLK.Text);
                    }
                    if (fieldName == "PIOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtPILICA.Text);
                    }
                    if (fieldName == "PIK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtPIK.Text);
                    }
                    if (fieldName == "ULIBR")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtULiBR.Text);
                    }
                    if (fieldName == "ULIBRK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtULiBK.Text);
                    }
                    if (fieldName == "GRADOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtGRAD.Text);
                    }
                    if (fieldName == "GRADK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtGradK.Text);
                    }
                    if (fieldName == "JMBGOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtJMBG.Text);
                    }
                    if (fieldName == "JMBGK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtJMBGK.Text);
                    }
                    if (fieldName == "LICNAK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtLicnaKartaK.Text);
                    }
                    if (fieldName == "OVBR")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtBrOvl.Text);
                    }
                    if (fieldName == "OVDATUM")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtDatumOvl.Text);
                    }
                }

            }
        }

        private void bunifuThinButton26_Click(object sender, EventArgs e)
        {
            Object oMissing = System.Reflection.Missing.Value;

            Object oTrue = true;
            Object oFalse = false;

            Word.Application oWord = new Word.Application();
            Word.Document oWordDoc = new Word.Document();

            oWord.Visible = true;
            Object oTemplatePath;
            long uk = 0;
            oTemplatePath = System.Windows.Forms.Application.StartupPath + "\\UgovorPOK.dot";

            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            foreach (Word.Field myMergeField in oWordDoc.Fields)
            {

                iTotalFields++;
                Word.Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;

                if (fieldText.StartsWith(" MERGEFIELD"))
                {

                    Int32 endMerge = fieldText.IndexOf("\\");
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(11, endMerge - 11);

                    fieldName = fieldName.Trim();

                    if (fieldName == "VRSTA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVrstaVoz.SelectedItem as string);
                    }
                    if (fieldName == "SASIJA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBrojSasije.Text);
                    }
                    if (fieldName == "GODINA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVGodinaProizvodnje.Text);
                    }
                    if (fieldName == "SNAGA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVSnaga.Text);
                    }

                    if (fieldName == "BRMOTORA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBrojMotora.Text);
                    }

                    if (fieldName == "MARKA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVMarka.Text);
                    }

                    if (fieldName == "MODEL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVModel.Text);
                    }

                    if (fieldName == "RZ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVZapremina.Text);
                    }

                    if (fieldName == "REG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtRegBr.Text);
                    }
                    if (fieldName == "DATUM")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(DateTime.Now.ToString("d.M.yyyy"));
                    }
                    if (fieldName == "PI")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVPrezime.Text + " " + txtVIme.Text);
                    }
                    if (fieldName == "ULICA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVUlica.Text);
                    }
                    if (fieldName == "BROJ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBroj.Text);
                    }
                    if (fieldName == "GRAD")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVMesto.Text);
                    }
                    if (fieldName == "JMBG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVJMBG.Text);
                    }
                    if (fieldName == "LICNA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBRLK.Text);
                    }
                    if (fieldName == "PIOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtPILICA.Text);
                    }
                    if (fieldName == "PIK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtPIK.Text);
                    }
                    if (fieldName == "ULIBR")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtULiBR.Text);
                    }
                    if (fieldName == "ULIBRK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtULiBK.Text);
                    }
                    if (fieldName == "GRADOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtGRAD.Text);
                    }
                    if (fieldName == "GRADK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtGradK.Text);
                    }
                    if (fieldName == "JMBGOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtJMBG.Text);
                    }
                    if (fieldName == "JMBGK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtJMBGK.Text);
                    }
                    if (fieldName == "LICNAK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtLicnaKartaK.Text);
                    }
                    if (fieldName == "OVBR")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtBrOvl.Text);
                    }
                    if (fieldName == "OVDATUM")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtDatumOvl.Text);
                    }

                }

            }
        }

        private void bunifuThinButton28_Click(object sender, EventArgs e)
        {
            Object oMissing = System.Reflection.Missing.Value;

            Object oTrue = true;
            Object oFalse = false;

            Word.Application oWord = new Word.Application();
            Word.Document oWordDoc = new Word.Document();

            oWord.Visible = true;
            Object oTemplatePath;
            long uk = 0;
            oTemplatePath = System.Windows.Forms.Application.StartupPath + "\\OvlascenjeINV.dot";

            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            foreach (Word.Field myMergeField in oWordDoc.Fields)
            {

                iTotalFields++;
                Word.Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;

                if (fieldText.StartsWith(" MERGEFIELD"))
                {

                    Int32 endMerge = fieldText.IndexOf("\\");
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(11, endMerge - 11);

                    fieldName = fieldName.Trim();

                    if (fieldName == "VRSTA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVrstaVoz.SelectedItem as string);
                    }
                    if (fieldName == "SASIJA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBrojSasije.Text);
                    }
                    if (fieldName == "GODINA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVGodinaProizvodnje.Text);
                    }
                    if (fieldName == "SNAGA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVSnaga.Text);
                    }

                    if (fieldName == "BRMOTORA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBrojMotora.Text);
                    }

                    if (fieldName == "MARKA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVMarka.Text);
                    }

                    if (fieldName == "MODEL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVModel.Text);
                    }

                    if (fieldName == "RZ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVZapremina.Text);
                    }

                    if (fieldName == "REG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtRegBr.Text);
                    }
                    if (fieldName == "DATUM")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(DateTime.Now.ToString("d.M.yyyy"));
                    }
                    if (fieldName == "PI")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVPrezime.Text + " " + txtVIme.Text);
                    }
                    if (fieldName == "ULICA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVUlica.Text);
                    }
                    if (fieldName == "BROJ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBroj.Text);
                    }
                    if (fieldName == "GRAD")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVMesto.Text);
                    }
                    if (fieldName == "JMBG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVJMBG.Text);
                    }
                    if (fieldName == "LICNA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBRLK.Text);
                    }
                    if (fieldName == "PIOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtPILICA.Text);
                    }
                    if (fieldName == "PIK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtPIK.Text);
                    }
                    if (fieldName == "ULIBR")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtULiBR.Text);
                    }
                    if (fieldName == "ULIBRK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtULiBK.Text);
                    }
                    if (fieldName == "GRADOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtGRAD.Text);
                    }
                    if (fieldName == "GRADK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtGradK.Text);
                    }
                    if (fieldName == "JMBGOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtJMBG.Text);
                    }
                    if (fieldName == "JMBGK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtJMBGK.Text);
                    }
                    if (fieldName == "LICNAK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtLicnaKartaK.Text);
                    }

                }

            }
        }

        private void bunifuThinButton212_Click_1(object sender, EventArgs e)
        {
            Object oMissing = System.Reflection.Missing.Value;

            Object oTrue = true;
            Object oFalse = false;

            Word.Application oWord = new Word.Application();
            Word.Document oWordDoc = new Word.Document();

            oWord.Visible = true;
            Object oTemplatePath;
            long uk = 0;
            oTemplatePath = System.Windows.Forms.Application.StartupPath + "\\Ovlascenje.dot";

            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            foreach (Word.Field myMergeField in oWordDoc.Fields)
            {

                iTotalFields++;
                Word.Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;

                if (fieldText.StartsWith(" MERGEFIELD"))
                {

                    Int32 endMerge = fieldText.IndexOf("\\");
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(11, endMerge - 11);

                    fieldName = fieldName.Trim();

                    if (fieldName == "VRSTA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVrstaVoz.SelectedItem as string);
                    }
                    if (fieldName == "SASIJA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBrojSasije.Text);
                    }
                    if (fieldName == "GODINA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVGodinaProizvodnje.Text);
                    }
                    if (fieldName == "SNAGA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVSnaga.Text);
                    }

                    if (fieldName == "BRMOTORA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBrojMotora.Text);
                    }

                    if (fieldName == "MARKA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVMarka.Text);
                    }

                    if (fieldName == "MODEL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVModel.Text);
                    }

                    if (fieldName == "RZ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVZapremina.Text);
                    }

                    if (fieldName == "REG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtRegBr.Text);
                    }
                    if (fieldName == "DATUM")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(DateTime.Now.ToString("d.M.yyyy"));
                    }
                    if (fieldName == "PI")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVPrezime.Text + " " + txtVIme.Text);
                    }
                    if (fieldName == "ULICA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVUlica.Text);
                    }
                    if (fieldName == "BROJ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBroj.Text);
                    }
                    if (fieldName == "GRAD")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(cbVMesto.Text);
                    }
                    if (fieldName == "JMBG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVJMBG.Text);
                    }
                    if (fieldName == "LICNA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtVBRLK.Text);
                    }
                    if (fieldName == "PIOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtPILICA.Text);
                    }
                    if (fieldName == "PIK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtPIK.Text);
                    }
                    if (fieldName == "ULIBR")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtULiBR.Text);
                    }
                    if (fieldName == "ULIBRK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtULiBK.Text);
                    }
                    if (fieldName == "GRADOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtGRAD.Text);
                    }
                    if (fieldName == "GRADK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtGradK.Text);
                    }
                    if (fieldName == "JMBGOL")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtJMBG.Text);
                    }
                    if (fieldName == "JMBGK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtJMBGK.Text);
                    }
                    if (fieldName == "LICNAK")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtLicnaKartaK.Text);
                    }

                }

            }
        }

        private void txtTPID_TextChanged(object sender, EventArgs e)
        {
            //lblDuzinaId.Text = txtTPID.Text.Length.ToString();
        }
        private void ResetFormData()
        {
            // Pauziraj renderovanje da ne trese UI
            this.SuspendLayout();
            dgwPoruke.SuspendLayout();

            // Obriši sve podatke iz DataTable i DataGridView
            dt.Clear();
            dt2.Clear();
            dt3.Clear();

            // Ako treba, resetuj ostale kontrole na default
            txtTPDatum.Text = DateTime.Now.ToString("dd.MM.yyyy");
            datumIsteka.Text = DateTime.Now.AddYears(1).ToString("dd.MM.yyyy");
            txtUplatniceDatum.Text = DateTime.Now.ToString("dd.MM.yyyy");
            cbVKuka.SelectedIndex = 1;

            this.ResumeLayout();
            dgwPoruke.ResumeLayout();
        }
        private void ResetujPoljaForme()
{
    // Umesto InitializeComponent, ovde resetuj samo ono što je neophodno
    foreach (Control c in this.Controls)
    {
        if (c is TextBox) c.Text = string.Empty;
        if (c is ComboBox) ((ComboBox)c).SelectedIndex = -1;
    }
}
        private void bunifuThinButton219_Click(object sender, EventArgs e)
        {
            //tabControl1.Refresh();
            //Panel parentForm = (this.Parent as Panel);
            //ucRegistracija p = new ucRegistracija();
            //p.Dock = DockStyle.Fill;
            //parentForm.Controls.Clear();
            //parentForm.Controls.Add(p);
            this.Controls.Clear();
            this.InitializeComponent();
            //dataGridView2.DataSource = null;
            dt.Clear();
            dataGridView2.Rows.Clear();
            dt2.Columns.Clear();
            dt2.Columns.Add("Ime i prezime vlasnika", typeof(string));
            dt2.Columns.Add("RegistarskiBroj", typeof(string));
            dt2.Columns.Add("Marka", typeof(string));
            dt2.Columns.Add("Model", typeof(string));
            dt2.Columns.Add("Ime i prezime stranke", typeof(string));
            dt2.Columns.Add("Broj telefona", typeof(string));
            dt2.Columns.Add("Datum isteka", typeof(string));

            dt3.Columns.Clear();
            dt3.Columns.Add("cbPoruka", typeof(bool));
            dt3.Columns.Add("Ime i prezime vlasnika", typeof(string));
            dt3.Columns.Add("RegistarskiBroj", typeof(string));
            dt3.Columns.Add("Ime i prezime stranke", typeof(string));
            dt3.Columns.Add("Broj telefona", typeof(string));
            dt3.Columns.Add("Datum isteka", typeof(string));
            dt3.Columns.Add("IDBroj", typeof(string));
            dtpOd.Value = DateTime.Now;
            dtpDo.Value = DateTime.Now.AddDays(7);

            var dolazilaVozila = b.ucitajTabliceDolazili(dtpOd.Value.AddDays(-21).ToString("dd/MM/yyyy"), dtpOd.Value.AddDays(1).ToString("dd/MM/yyyy"));
            var vozilas2 = new List<vozilo>();
            if (DateTime.Now.DayOfWeek == DayOfWeek.Monday)
                vozilas2 = b.ucitajTabliceIsticu(dtpDo.Value.AddDays(-2).ToString("dd/MM/yyyy"), dtpDo.Value.ToString("dd/MM/yyyy"));
            else
                vozilas2 = b.ucitajTabliceIsticu(dtpDo.Value.ToString("dd/MM/yyyy"), dtpDo.Value.ToString("dd/MM/yyyy"));
            var vozilas = vozilas2.Where(x => x.IDBroj[8] == '1' && x.Ispravnost == "Da");
            dt3.Clear();
            //vozilas.Reverse();
            var vozila = vozilas.GroupBy(x => new { x.ImeVL, x.RegistarskiBroj }).Select(y => y.First());
            //TimeSpan razlika2 = new TimeSpan(365, 0, 0, 0);
            int j = 0;
            foreach (var auto in vozila)
            {
                //DateTime datum;
                try
                {
                    //datum = DateTime.ParseExact(auto.Datum, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                    //if ((datum + razlika2) >= dtpOd.Value && (datum+razlika2) <= dtpDo.Value)
                    var dolazio = dolazilaVozila.Where(a => a.BrojSasije == auto.BrojSasije).FirstOrDefault();
                    if (dolazio != null) auto.Poruka = "Da";
                    dgwPoruke.Rows.Add((auto.Poruka != "Da") ? true : false, auto.ImeVL + " " + auto.PrezimeVL, auto.RegistarskiBroj, auto.ImeDL + " " + auto.PrezimeDL, (auto.TelefonaDL != "") ? auto.TelefonaDL : auto.TelefonaVL, auto.datumIsteka, auto.IDBroj);
                    if (auto.Poruka == "Da")
                    {
                        dgwPoruke.Rows[j].Cells[0].Style.BackColor = Color.FromArgb(229, 226, 244);
                        dgwPoruke.Rows[j].Cells[1].Style.BackColor = Color.FromArgb(229, 226, 244);
                        dgwPoruke.Rows[j].Cells[2].Style.BackColor = Color.FromArgb(229, 226, 244);
                        dgwPoruke.Rows[j].Cells[3].Style.BackColor = Color.FromArgb(229, 226, 244);
                        dgwPoruke.Rows[j].Cells[4].Style.BackColor = Color.FromArgb(229, 226, 244);
                        dgwPoruke.Rows[j].Cells[5].Style.BackColor = Color.FromArgb(229, 226, 244);
                    }
                }
                catch { }
                j++;
            }
            //var timer = System.Diagnostics.Stopwatch.StartNew();


            //try
            //{
            //var client = new TcpClient("time.nist.gov", 13);
            //var streamReader = new StreamReader(client.GetStream());
            //var response = streamReader.ReadToEnd();
            //var utcDateTimeString = response.Substring(7, 17);
            //var localDateTime = DateTime.ParseExact(utcDateTimeString, "yy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal);
            //// if (localDateTime.Date < DateTime.Parse("31.12.2020"))
            //{

            txtTPDatum.Text = DateTime.Now.ToString("dd.MM.yyyy");
            datumIsteka.Text = (DateTime.Now).AddYears(1).ToString("dd.MM.yyyy");
            txtUplatniceDatum.Text = DateTime.Now.ToString("dd.MM.yyyy");
            cbVKuka.SelectedIndex = 1;
            textBox10.Text = b.nadjiMarginuLevo();
            textBox11.Text = b.nadjiMarginuDesno();
            brOtpremnice = b.nadjiOtpremnicu();
            txtOtp.Text = brOtpremnice.ToString();
            //podesavanje cb autocomplete
            cbTPKontr1.DropDownStyle = ComboBoxStyle.DropDown;
            cbTPKontr1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbTPKontr1.AutoCompleteSource = AutoCompleteSource.ListItems;

            cbTPKontr2.DropDownStyle = ComboBoxStyle.DropDown;
            cbTPKontr2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbTPKontr2.AutoCompleteSource = AutoCompleteSource.ListItems;

            cbTablice.DropDownStyle = ComboBoxStyle.DropDown;
            cbTablice.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbTablice.AutoCompleteSource = AutoCompleteSource.ListItems;

            cbFakture.DropDownStyle = ComboBoxStyle.DropDown;
            cbFakture.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbFakture.AutoCompleteSource = AutoCompleteSource.ListItems;

            cbSasije.DropDownStyle = ComboBoxStyle.DropDown;
            cbSasije.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbSasije.AutoCompleteSource = AutoCompleteSource.ListItems;

            cbIdjevi.DropDownStyle = ComboBoxStyle.DropDown;
            cbIdjevi.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cbIdjevi.AutoCompleteSource = AutoCompleteSource.ListItems;
            //cbMusterije.DropDownStyle = ComboBoxStyle.DropDown;
            //cbMusterije.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            //cbMusterije.AutoCompleteSource = AutoCompleteSource.ListItems;

            //ucitavanje tablica, sasije i idjeva
            cbFakture.DataSource = b.ucitajFakture();
            cbTablice.DataSource = b.ucitajTablice();
            cbSasije.DataSource = b.ucitajSasije();
            cbIdjevi.DataSource = b.ucitajIdjeve();
            prvoUcitavanje = false;
            cbTablice.SelectedIndex = -1;
            cbSasije.SelectedIndex = -1;
            cbIdjevi.SelectedIndex = -1;
            cbIdjevi.Text = "141022" + DateTime.Now.Year % 100;

            //ucitavanje kontrolora
            cbTPKontr1.DataSource = b.nadjiKotrolore();
            cbTPKontr1.SelectedIndex = 0;
            cbTPKontr2.DataSource = b.nadjiKotrolore();
            cbTPKontr2.SelectedIndex = 1;

            //podesavanje datetimepickera
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd.MM.yyyy";


            //podesavanje polja na fakturi
            txtNaziv.Text = "Tehnički pregled vozila";
            //txtSfr.Text = "3000";
            txtRabat.Text = "0";

            //podesavanje datatable za fakturu
            dt.Columns.Clear(); ;
            dt.Columns.Add("Šifra", typeof(int));
            dt.Columns.Add("Naziv", typeof(string));
            dt.Columns.Add("JM", typeof(string));
            dt.Columns.Add("Količina", typeof(double));
            dt.Columns.Add("Cena", typeof(double));
            dt.Columns.Add("Iznos", typeof(double));
            dt.Columns.Add("Rabat %", typeof(double));
            dt.Columns.Add("Rabat", typeof(double));
            dt.Columns.Add("PDV %", typeof(string));
            dt.Columns.Add("PDV", typeof(double));
            dt.Columns.Add("Cena sa PDV", typeof(double));
            dt.Columns.Add("Iznos sa PDV", typeof(double));
            dataGridView2.DataSource = dt;

            //podesavanje kontrolora
            cbTPKontr1.SelectedIndex = 0;
            cbTPKontr2.SelectedIndex = 1;

            //podesavanje praznog polja za vrijeme
            txtTPVremeOd.CustomFormat = " ";
            txtTPVremeDo.CustomFormat = " ";
            txtTPVremeDo.Format = DateTimePickerFormat.Custom;
            txtTPVremeOd.Format = DateTimePickerFormat.Custom;


            //pravljenje id broja
            try
            {
                string sifra = "141022";
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 6; i++)
                    sb.Append(sifra[i]);
                sifra = (DateTime.Now.Year % 100).ToString();
                for (int i = 0; i < 2; i++)
                    sb.Append(sifra[i]);
                sb.Append(1);
                string poslednji = b.nadjiPoslednjiPregled();
                int sledeci = int.Parse(poslednji);
                sledeci++;
                poslednji = sledeci.ToString().PadLeft(5, '0');
                Brpregleda = poslednji;
                for (int i = 0; i < 5; i++)
                    sb.Append(poslednji[i]);
                txtTPID.Text = sb.ToString();
            }
            catch { }

            //odabir prvog id broja
            comboBox2.SelectedIndex = 0;

            //dodavanje podataka u cb za neispravnosti
            List<string> lista = b.lista_grupa();
            cbGlPodela.DataSource = lista;
            cbGlPodela.Text = "";

            popuniUplatnice();

            ispravnost = "Da";
            UkljuciIskljuciNeispravnosti();
            //popunjavanje marki vozila
            //List<string> marke = new List<string>();
            //foreach (var b in b.lista_marka())
            //{
            //    marke.Add(b.KaroserijaNaziv);
            //}
            //cbVMarka.DataSource = marke;
            //cbVMarka.Text = "";

            ucRegistracija_Load(sender, e);
            //}
            //else MessageBox.Show("Windows couldn’t complete the requested changes.The changes couldn’t be completed.\nPlease connect to the Internet and try again.Error code: 0x800F0954");

            //}
            //catch (Exception)
            //{

            //    MessageBox.Show("Windows couldn’t complete the requested changes.The changes couldn’t be completed.\nPlease connect to the Internet and try again.Error code: 0x800F0954");
            //}
        }

        private void cbVPV_SelectedIndexChanged(object sender, EventArgs e)
        {
            //string s;
            //if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
            //    s = cbVKlasa.Text + "/" + cbVOKS.Text;
            //else
            //    if (cbDecOzn.Text != "")
            //    if (checkBox12.Checked)
            //        s = (cbVSpec.Text.Split(' '))[0].ToString() + (cbDecOzn.Text.Split(' '))[0].ToString();
            //    else
            //        s = cbVOKS.Text + (cbDecOzn.Text.Split(' '))[0].ToString();
            //else
            //    s = cbVOKS.Text;
            //txtVrstaPlusDecOznk.Text = s;
        }

        private void cbDecOzn_SelectedIndexChanged(object sender, EventArgs e)
        {
            string s;
            if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
                s = cbVKlasa.Text + "/" + cbVOKS.Text;
            else
                if (cbDecOzn.Text != "")
                if (checkBox12.Checked)
                    s = (cbVSpec.Text.Split(' '))[0].ToString() + (cbDecOzn.Text.Split(' '))[0].ToString();
                else
                    s = cbVOKS.Text + (cbDecOzn.Text.Split(' '))[0].ToString();
            else
                s = cbVOKS.Text;
            txtVrstaPlusDecOznk.Text = s;
        }

        private void txtVJMBG_TextChanged(object sender, EventArgs e)
        {

            lblDuzinaMatBr.Text = txtVJMBG.Text.Length.ToString();
            popuniFakture();
        }

        private void txtVBRLK_TextChanged(object sender, EventArgs e)
        {
            lblDuzinaLK.Text = txtVBRLK.Text.Length.ToString();
            popuniFakture();
        }

        private void txtDLJMBG_TextChanged(object sender, EventArgs e)
        {

            lblDLDuzinaMatBr.Text = txtDLJMBG.Text.Length.ToString();
        }

        private void txtDLBRLK_TextChanged(object sender, EventArgs e)
        {
            lblDLDuzinaLK.Text = txtDLBRLK.Text.Length.ToString();
        }

        private void label83_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@Application.StartupPath + "\\" + "Pravilnik o podeli vozila.pdf");
            }
            catch (Exception)
            {
                MessageBox.Show("Ne postoji fajl sa imenom \'Pravilnik o podeli vozila.pdf\'!");
            }
        }

        private void label86_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@Application.StartupPath + "\\" + "Pravilnik o teh. pregledu.pdf");
            }
            catch (Exception)
            {
                MessageBox.Show("Ne postoji fajl sa imenom \'Pravilnik o teh. pregledu.pdf\'!");
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                txtRegBr.Clear();
        }

        private void bunifuThinButton214_Click_1(object sender, EventArgs e)
        {
            rbTehIsp.Checked = true;
            txtTPNap.Text = "";
            rbTehIsp_CheckedChanged(sender, e);
            ispravnost = "Da";
            if (txtTPID.Text[8] == '1')
            {
                try
                {
                    var datumIst = DateTime.Parse(datumIsteka.Text);
                    if (datumIst >= DateTime.Now)
                    {
                        if (datumIst.Year < DateTime.Now.Year + 1)
                            datumIsteka.Text = datumIst.AddYears(1).ToString("dd.MM.yyyy");
                    }
                    else
                        datumIsteka.Text = (DateTime.Now).AddYears(1).ToString("dd.MM.yyyy");
                }
                catch
                {
                    datumIsteka.Text = (DateTime.Now).AddYears(1).ToString("dd.MM.yyyy");
                }
            }
            try
            {
                string sifra = "141022";
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 6; i++)
                    sb.Append(sifra[i]);
                sifra = (DateTime.Now.Year % 100).ToString();
                for (int i = 0; i < 2; i++)
                    sb.Append(sifra[i]);
                sb.Append(1);
                string poslednji = b.nadjiPoslednjiPregled();
                int sledeci = int.Parse(poslednji);
                sledeci++;
                poslednji = sledeci.ToString().PadLeft(5, '0');
                Brpregleda = poslednji;
                for (int i = 0; i < 5; i++)
                    sb.Append(poslednji[i]);
                txtTPID.Text = sb.ToString();
                comboBox2.SelectedIndex = 0;
            }
            catch { }
            txtTPDatum.Text = DateTime.Now.ToString("dd.MM.yyyy");
            //datumIsteka.Text = (DateTime.Now).AddYears(1).ToString("dd.MM.yyyy");
            txtTPVremeOd.Text = DateTime.Now.ToString("HH:mm");
            dataGridView1.Rows.Clear();
            txtTPVremeDo.CustomFormat = " ";
            txtTPVremeDo.Format = DateTimePickerFormat.Custom;
            btnIzv.Visible = false;
            btnPot2.Visible = false;
            btnReg2.Visible = false;
            btnZap1.Visible = false;
            txtImePrezimePretraga.Text = "";
            popuniPrethodneNeispravnosti();
            dataGridView1.Rows.Clear();
        }

        private void bunifuThinButton220_Click(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();

            PrintDocument printDocument = new PrintDocument();

            PaperSize ps = new PaperSize();
            ps.RawKind = (int)PaperKind.A4;
            printDocument.DefaultPageSettings.PaperSize = ps;
            //printDocument.DefaultPageSettings.Landscape = true;
            printDialog.Document = printDocument; //add the document to the dialog box...        

            printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16); //add an event handler that will do the printing

            //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

            DialogResult result = printDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                printDocument.Print();

            }
        }

        private void txtVZiroRacun_TextChanged(object sender, EventArgs e)
        {

            if (checkBox2.Checked)
                textBox15.Text = txtVZiroRacun.Text;
            else
                textBox15.Text = txtDLZiroRacun.Text;

            popuniFakture();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                comboBox1.SelectedIndex = 0;
                bunifuThinButton220_Click(sender, e);
            }
        }

        private void bunifuThinButton221_Click(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();

            PrintDocument printDocument = new PrintDocument();

            PaperSize ps = new PaperSize();
            ps.RawKind = (int)PaperKind.A4;
            printDocument.DefaultPageSettings.PaperSize = ps;
            //printDocument.DefaultPageSettings.Landscape = true;
            printDialog.Document = printDocument; //add the document to the dialog box...        

            //add an event handler that will do the printing

            //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

            DialogResult result = printDialog.ShowDialog();

            if (result == DialogResult.OK)
            {


                if (checkBox4.Checked)
                {
                    if (checkBox3.Checked)
                    {
                        comboBox1.SelectedIndex = 0;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }

                    if (checkBox5.Checked)
                    {
                        comboBox1.SelectedIndex = 1;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                    if (checkBox6.Checked)
                    {
                        comboBox1.SelectedIndex = 2;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                    if (checkBox7.Checked)
                    {
                        comboBox1.SelectedIndex = 3;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                    if (checkBox8.Checked)
                    {
                        comboBox1.SelectedIndex = 4;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                    if (checkBox9.Checked)
                    {
                        comboBox1.SelectedIndex = 5;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                    if (checkBox10.Checked)
                    {
                        comboBox1.SelectedIndex = 6;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                    if (checkBox11.Checked)
                    {
                        comboBox1.SelectedIndex = 7;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                    if (checkBox14.Checked)
                    {
                        comboBox1.SelectedIndex = 8;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                    if (checkBox13.Checked)
                    {
                        comboBox1.SelectedIndex = 9;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                    if (checkBox15.Checked)
                    {
                        comboBox1.SelectedIndex = 10;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                    if (checkBox16.Checked)
                    {
                        comboBox1.SelectedIndex = 11;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt16);
                        printDocument.Print();
                    }
                }
                else
                {
                    if (checkBox3.Checked)
                    {
                        comboBox1.SelectedIndex = 0;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }

                    if (checkBox5.Checked)
                    {
                        comboBox1.SelectedIndex = 1;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                    if (checkBox6.Checked)
                    {
                        comboBox1.SelectedIndex = 2;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                    if (checkBox7.Checked)
                    {
                        comboBox1.SelectedIndex = 3;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                    if (checkBox8.Checked)
                    {
                        comboBox1.SelectedIndex = 4;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                    if (checkBox9.Checked)
                    {
                        comboBox1.SelectedIndex = 5;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                    if (checkBox10.Checked)
                    {
                        comboBox1.SelectedIndex = 6;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                    if (checkBox11.Checked)
                    {
                        comboBox1.SelectedIndex = 7;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                    if (checkBox14.Checked)
                    {
                        comboBox1.SelectedIndex = 8;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                    if (checkBox13.Checked)
                    {
                        comboBox1.SelectedIndex = 9;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                    if (checkBox15.Checked)
                    {
                        comboBox1.SelectedIndex = 10;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                    if (checkBox16.Checked)
                    {
                        comboBox1.SelectedIndex = 11;
                        printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt6);
                        printDocument.Print();
                    }
                }

            }
        }

        private void cbVKlasa_SelectedIndexChanged(object sender, EventArgs e)
        {
            string s;
            if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
                s = cbVKlasa.Text + "/" + cbVOKS.Text;
            else
                if (cbDecOzn.Text != "")
                if (checkBox12.Checked)
                    s = (cbVSpec.Text.Split(' '))[0].ToString() + (cbDecOzn.Text.Split(' '))[0].ToString();
                else
                    s = cbVOKS.Text + (cbDecOzn.Text.Split(' '))[0].ToString();
            else
                s = cbVOKS.Text;
            txtVrstaPlusDecOznk.Text = s;
        }

        private void cbVSpec_SelectedIndexChanged(object sender, EventArgs e)
        {
            string s;
            if (checkBox12.Checked)
            {
                if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
                    s = cbVKlasa.Text + "/" + (cbVSpec.Text.Split(' '))[0].ToString();
                else
                    if (cbDecOzn.Text != "")
                    s = (cbVSpec.Text.Split(' '))[0].ToString() + (cbDecOzn.Text.Split(' '))[0].ToString();
                else
                    s = (cbVSpec.Text.Split(' '))[0].ToString();
                txtVrstaPlusDecOznk.Text = s;
            }
        }

        private void txtTPVremeDo_ValueChanged(object sender, EventArgs e)
        {
            if (DateTime.TryParse(txtTPVremeDo.Text, out DateTime dt) &&
                DateTime.TryParse(txtTPVremeOd.Text, out DateTime dt2))
            {
                int raz = (int)(dt - dt2).TotalMinutes; // ← KLJUČNA IZMENA

                txtTPRazlikaVremena.Text = raz.ToString();

                if (raz < duzinaTehPregleda)
                    txtTPRazlikaVremena.BackColor = Color.LightPink;
                else
                    txtTPRazlikaVremena.BackColor = Color.FromArgb(240, 240, 240);
            }
        }

        //blanko uplatnica
        private void bunifuThinButton215_Click_1(object sender, EventArgs e)
        {
            richTextBox4.Text = "";
            richTextBox2.Text = "";
            richTextBox3.Text = "";
            textBox18.Text = "";
            textBox16.Text = "";
            textBox15.Text = "";
            textBox14.Text = "";
            textBox8.Text = "";
            textBox12.Text = "";

            textBox9.Text = "";
            richTextBox1.Text = "";
            richTextBox5.Text = "";
            richTextBox6.Text = "";
            textBox27.Text = "";
            textBox26.Text = "";
            textBox25.Text = "";
            textBox23.Text = "";
            textBox21.Text = "";
            textBox17.Text = "";
            textBox19.Text = "";
        }

        private void txtVPrezime_TextChanged(object sender, EventArgs e)
        {
            popuniFakture();
        }

        private void popuniFakture()
        {
            if (cbDLPravno.Checked)
            {
                textBox1.Text = txtDLPrezime.Text;
                textBox2.Text = txtDLUlica.Text + " " + txtDLBroj.Text;
                textBox3.Text = txtDLBrPoste.Text + " " + cbDLMesto.Text;
                textBox6.Text = txtDLBrTelefona.Text;
                textBox7.Text = txtDLZiroRacun.Text;
                textBox4.Text = txtDLBRLK.Text;
                textBox20.Text = txtDLJMBG.Text;
            }
            else if (checkBox2.Checked)
            {
                textBox1.Text = txtVPrezime.Text;
                textBox2.Text = txtVUlica.Text + " " + txtVBroj.Text;
                textBox3.Text = txtVBrPoste.Text + " " + cbVMesto.Text;
                textBox6.Text = txtVBrTelefona.Text;
                textBox7.Text = txtVZiroRacun.Text;
                textBox4.Text = txtVBRLK.Text;
                textBox20.Text = txtVJMBG.Text;
            }
        }
        private void txtVUlica_TextChanged(object sender, EventArgs e)
        {
            popuniFakture();
        }

        private void cbVMesto_TextChanged(object sender, EventArgs e)
        {
            popuniFakture();
        }

        private void txtVBrTelefona_TextChanged(object sender, EventArgs e)
        {
            popuniFakture();

        }

        private void textBox11_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                b.promeniMargine(textBox10.Text, textBox11.Text);
            }
        }

        private void cbGlPodela_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> lista = b.lista_delova(cbGlPodela.Text);
            cbNaziv.DataSource = lista;
            cbNaziv.Text = "";
        }

        private void dtpOd_onValueChanged(object sender, EventArgs e)
        {
            if (prvoUcitavanje == false)
            {
                //var timer = System.Diagnostics.Stopwatch.StartNew();
                var vozilas = b.ucitajTabliceIsticu(dtpOd.Value.ToString("dd/MM/yyyy"), dtpDo.Value.ToString("dd/MM/yyyy"));
                dt2.Clear();
                //vozilas.Reverse();
                var vozila = vozilas.GroupBy(x => new { x.ImeVL, x.RegistarskiBroj }).Select(y => y.First());
                //TimeSpan razlika2 = new TimeSpan(365, 0, 0, 0);
                foreach (var auto in vozila)
                {
                    //DateTime datum;
                    try
                    {
                        //datum = DateTime.ParseExact(auto.Datum, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                        //if ((datum + razlika2) >= dtpOd.Value && (datum+razlika2) <= dtpDo.Value)
                        dt2.Rows.Add(auto.ImeVL + " " + auto.PrezimeVL, auto.RegistarskiBroj, auto.Marka, auto.Model, auto.ImeDL + " " + auto.PrezimeDL, (auto.TelefonaDL != "") ? auto.TelefonaDL : auto.TelefonaVL, auto.Datum);
                    }
                    catch { }
                }
                dgwRacun.DataSource = null;
                dgwRacun.DataSource = dt2;
                //timer.Stop();
                //MessageBox.Show(timer.ElapsedMilliseconds.ToString());
            }
        }

        private void dtpDo_onValueChanged(object sender, EventArgs e)
        {
            if (prvoUcitavanje == false)
            {
                //var timer = System.Diagnostics.Stopwatch.StartNew();
                var vozilas = b.ucitajTabliceIsticu(dtpOd.Value.ToString("dd/MM/yyyy"), dtpDo.Value.ToString("dd/MM/yyyy"));
                dt2.Clear();
                //vozilas.Reverse();
                var vozila = vozilas.GroupBy(x => new { x.ImeVL, x.RegistarskiBroj }).Select(y => y.First());
                //TimeSpan razlika2 = new TimeSpan(365, 0, 0, 0);
                foreach (var auto in vozila)
                {
                    //DateTime datum;
                    try
                    {
                        //datum = DateTime.ParseExact(auto.Datum, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                        //if ((datum + razlika2) >= dtpOd.Value && (datum+razlika2) <= dtpDo.Value)
                        dt2.Rows.Add(auto.ImeVL + " " + auto.PrezimeVL, auto.RegistarskiBroj, auto.Marka, auto.Model, auto.ImeDL + " " + auto.PrezimeDL, (auto.TelefonaDL != "") ? auto.TelefonaDL : auto.TelefonaVL, auto.Datum, auto.BrojSasije);
                    }
                    catch { }
                }
                dgwRacun.DataSource = null;
                dgwRacun.DataSource = dt2;
                //timer.Stop();
                //MessageBox.Show(timer.ElapsedMilliseconds.ToString());
            }
        }

        private void dtpDo_ValueChanged(object sender, EventArgs e)
        {
            //if (prvoUcitavanje == false)
            {
                //var timer = System.Diagnostics.Stopwatch.StartNew();
                var vozilas2 = b.ucitajTabliceIsticu(dtpOd.Value.ToString("dd/MM/yyyy"), dtpDo.Value.ToString("dd/MM/yyyy"));
                dt2.Clear();
                dgwRacun.Rows.Clear();
                var vozilas = vozilas2.Where(x => x.IDBroj[8] == '1' && x.Ispravnost == "Da");

                //vozilas.Reverse();
                var vozila = vozilas.GroupBy(x => new { x.ImeVL, x.RegistarskiBroj }).Select(y => y.First());
                //TimeSpan razlika2 = new TimeSpan(365, 0, 0, 0);
                foreach (var auto in vozila)
                {
                    //DateTime datum;
                    try
                    {
                        //datum = DateTime.ParseExact(auto.Datum, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                        //if ((datum + razlika2) >= dtpOd.Value && (datum+razlika2) <= dtpDo.Value)
                        if (!string.IsNullOrWhiteSpace(auto.ImeVL) || !string.IsNullOrWhiteSpace(auto.Marka) || !string.IsNullOrWhiteSpace(auto.Model))
                            dgwRacun.Rows.Add(auto.ImeVL + " " + auto.PrezimeVL, auto.RegistarskiBroj, auto.Marka, auto.Model, auto.ImeDL + " " + auto.PrezimeDL, (auto.TelefonaDL != "") ? auto.TelefonaDL : auto.TelefonaVL, auto.datumIsteka, auto.IDBroj);
                    }
                    catch { }
                }
                //dgwRacun.DataSource = null;
                //dgwRacun.DataSource = dt2;
                //dgwRacun.AutoResizeColumns();
                //timer.Stop();
                //MessageBox.Show(timer.ElapsedMilliseconds.ToString());
            }
        }

        private void bunifuThinButton222_Click(object sender, EventArgs e)
        {
            var lista = b.ucitajSortPoId();
            int god = DateTime.Now.Year % 100;
            lista = lista.Where(x => x.IDBroj[6].ToString() == (god / 10).ToString() && x.IDBroj[7].ToString() == (god % 10).ToString()).ToList();
            var listaMOPED = lista.Where(x => x.VrstaVozila == "MOPED").ToList();
            var listaLAKITRICIKL = lista.Where(x => x.VrstaVozila == "LAKI TRICIKL").ToList();
            var listaMOTOCIKL = lista.Where(x => x.VrstaVozila == "MOTOCIKL").ToList();
            //var listaMOTOCIKLSABOCNIMSEDISTEM = lista.Where(x => x.VrstaVozila == "MOTOCIKL SA BOCNIM SEDISTEM").ToList();
            var listaTESKITRICIKL = lista.Where(x => x.VrstaVozila == "TESKI TRICIKL").ToList();
            var listaLAKICETVOROCIKL = lista.Where(x => x.VrstaVozila == "LAKI CETVOROCIKL").ToList();
            var listaTESKICETVOROCIKL = lista.Where(x => x.VrstaVozila == "TESKI CETVOROCIKL").ToList();
            var listaPUTNICKOVOZILO = lista.Where(x => x.VrstaVozila == "PUTNICKO VOZILO").ToList();
            var listaAUTOBUS = lista.Where(x => x.VrstaVozila == "AUTOBUS").ToList();
            var listaTERETNOVOZILO = lista.Where(x => x.VrstaVozila == "TERETNO VOZILO").ToList();
            var listaRADNAMASINA = lista.Where(x => x.VrstaVozila == "RADNA MASINA").ToList();
            var listaTRAKTOR = lista.Where(x => x.VrstaVozila == "TRAKTOR").ToList();
            var listaMOTOKULTIVATOR = lista.Where(x => x.VrstaVozila == "MOTOKULTIVATOR").ToList();
            var listaPRIKLJUCNOVOZILO = lista.Where(x => x.VrstaVozila == "PRIKLJUCNO VOZILO").ToList();
            var listaPRIKLJUCNOZATRAKTOR = lista.Where(x => x.VrstaVozila == "PRIKLJUCNO ZA TRAKTOR").ToList();
            //var listaOSTALO = lista.Where(x => x.VrstaVozila == "OSTALO").ToList();
            #region MOPED
            int brISPRAVNIHREDOVNIHMOPED = nadjiBrojPoVrsti('1', "Da", listaMOPED);
            int brNEISPRAVNIHREDOVNIHMOPED = nadjiBrojPoVrsti('1', "Ne", listaMOPED);
            int brISPRAVNIHSESTOMESECNIHMOPED = nadjiBrojPoVrsti('2', "Da", listaMOPED);
            int brNEISPRAVNIHSESTOMESECNIHMOPED = nadjiBrojPoVrsti('2', "Ne", listaMOPED);
            int brISPRAVNIHVANREDOVNIHMOPED = nadjiBrojPoVrsti('4', "Da", listaMOPED);
            int brNEISPRAVNIHVANREDOVNIHMOPED = nadjiBrojPoVrsti('4', "Ne", listaMOPED);
            int brISPRAVNIHKONTROLNIHMOPED = nadjiBrojPoVrsti('3', "Da", listaMOPED);
            int brNEISPRAVNIHKONTROLNIHMOPED = nadjiBrojPoVrsti('3', "Ne", listaMOPED);
            int brPREKINUTIHREDOVNIHMOPED = nadjiBrojPoVrsti('1', "Prekid", listaMOPED);
            int brPREKINUTIHSESTOMESECNIHMOPED = nadjiBrojPoVrsti('2', "Prekid", listaMOPED);
            int brPREKINUTIHVANREDOVNIHMOPED = nadjiBrojPoVrsti('4', "Prekid", listaMOPED);
            // MOPED - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHMOPEDPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaMOPED);
            int brNEISPRAVNIHREDOVNIHMOPEDPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaMOPED);
            int brISPRAVNIHSESTOMESECNIHMOPEDPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaMOPED);
            int brNEISPRAVNIHSESTOMESECNIHMOPEDPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaMOPED);
            int brISPRAVNIHVANREDOVNIHMOPEDPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaMOPED);
            int brNEISPRAVNIHVANREDOVNIHMOPEDPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaMOPED);
            #endregion
            #region LAKITRICIKL
            int brISPRAVNIHREDOVNIHLAKITRICIKL = nadjiBrojPoVrsti('1', "Da", listaLAKITRICIKL);
            int brNEISPRAVNIHREDOVNIHLAKITRICIKL = nadjiBrojPoVrsti('1', "Ne", listaLAKITRICIKL);
            int brISPRAVNIHSESTOMESECNIHLAKITRICIKL = nadjiBrojPoVrsti('2', "Da", listaLAKITRICIKL);
            int brNEISPRAVNIHSESTOMESECNIHLAKITRICIKL = nadjiBrojPoVrsti('2', "Ne", listaLAKITRICIKL);
            int brISPRAVNIHVANREDOVNIHLAKITRICIKL = nadjiBrojPoVrsti('4', "Da", listaLAKITRICIKL);
            int brNEISPRAVNIHVANREDOVNIHLAKITRICIKL = nadjiBrojPoVrsti('4', "Ne", listaLAKITRICIKL);
            int brISPRAVNIHKONTROLNIHLAKITRICIKL = nadjiBrojPoVrsti('3', "Da", listaLAKITRICIKL);
            int brNEISPRAVNIHKONTROLNIHLAKITRICIKL = nadjiBrojPoVrsti('3', "Ne", listaLAKITRICIKL);
            int brPREKINUTIHREDOVNIHLAKITRICIKL = nadjiBrojPoVrsti('1', "Prekid", listaLAKITRICIKL);
            int brPREKINUTIHSESTOMESECNIHLAKITRICIKL = nadjiBrojPoVrsti('2', "Prekid", listaLAKITRICIKL);
            int brPREKINUTIHVANREDOVNIHLAKITRICIKL = nadjiBrojPoVrsti('4', "Prekid", listaLAKITRICIKL);
            // LAKITRICIKL - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHLAKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaLAKITRICIKL);
            int brNEISPRAVNIHREDOVNIHLAKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaLAKITRICIKL);
            int brISPRAVNIHSESTOMESECNIHLAKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaLAKITRICIKL);
            int brNEISPRAVNIHSESTOMESECNIHLAKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaLAKITRICIKL);
            int brISPRAVNIHVANREDOVNIHLAKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaLAKITRICIKL);
            int brNEISPRAVNIHVANREDOVNIHLAKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaLAKITRICIKL);
            #endregion
            #region MOTOCIKL
            int brISPRAVNIHREDOVNIHMOTOCIKL = nadjiBrojPoVrsti('1', "Da", listaMOTOCIKL);
            int brNEISPRAVNIHREDOVNIHMOTOCIKL = nadjiBrojPoVrsti('1', "Ne", listaMOTOCIKL);
            int brISPRAVNIHSESTOMESECNIHMOTOCIKL = nadjiBrojPoVrsti('2', "Da", listaMOTOCIKL);
            int brNEISPRAVNIHSESTOMESECNIHMOTOCIKL = nadjiBrojPoVrsti('2', "Ne", listaMOTOCIKL);
            int brISPRAVNIHVANREDOVNIHMOTOCIKL = nadjiBrojPoVrsti('4', "Da", listaMOTOCIKL);
            int brNEISPRAVNIHVANREDOVNIHMOTOCIKL = nadjiBrojPoVrsti('4', "Ne", listaMOTOCIKL);
            int brISPRAVNIHKONTROLNIHMOTOCIKL = nadjiBrojPoVrsti('3', "Da", listaMOTOCIKL);
            int brNEISPRAVNIHKONTROLNIHMOTOCIKL = nadjiBrojPoVrsti('3', "Ne", listaMOTOCIKL);
            int brPREKINUTIHREDOVNIHMOTOCIKL = nadjiBrojPoVrsti('1', "Prekid", listaMOTOCIKL);
            int brPREKINUTIHSESTOMESECNIHMOTOCIKL = nadjiBrojPoVrsti('2', "Prekid", listaMOTOCIKL);
            int brPREKINUTIHVANREDOVNIHMOTOCIKL = nadjiBrojPoVrsti('4', "Prekid", listaMOTOCIKL);
            // MOTOCIKL - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHMOTOCIKLPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaMOTOCIKL);
            int brNEISPRAVNIHREDOVNIHMOTOCIKLPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaMOTOCIKL);
            int brISPRAVNIHSESTOMESECNIHMOTOCIKLPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaMOTOCIKL);
            int brNEISPRAVNIHSESTOMESECNIHMOTOCIKLPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaMOTOCIKL);
            int brISPRAVNIHVANREDOVNIHMOTOCIKLPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaMOTOCIKL);
            int brNEISPRAVNIHVANREDOVNIHMOTOCIKLPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaMOTOCIKL);
            #endregion
            #region TESKITRICIKL
            int brISPRAVNIHREDOVNIHTESKITRICIKL = nadjiBrojPoVrsti('1', "Da", listaTESKITRICIKL);
            int brNEISPRAVNIHREDOVNIHTESKITRICIKL = nadjiBrojPoVrsti('1', "Ne", listaTESKITRICIKL);
            int brISPRAVNIHSESTOMESECNIHTESKITRICIKL = nadjiBrojPoVrsti('2', "Da", listaTESKITRICIKL);
            int brNEISPRAVNIHSESTOMESECNIHTESKITRICIKL = nadjiBrojPoVrsti('2', "Ne", listaTESKITRICIKL);
            int brISPRAVNIHVANREDOVNIHTESKITRICIKL = nadjiBrojPoVrsti('4', "Da", listaTESKITRICIKL);
            int brNEISPRAVNIHVANREDOVNIHTESKITRICIKL = nadjiBrojPoVrsti('4', "Ne", listaTESKITRICIKL);
            int brISPRAVNIHKONTROLNIHTESKITRICIKL = nadjiBrojPoVrsti('3', "Da", listaTESKITRICIKL);
            int brNEISPRAVNIHKONTROLNIHTESKITRICIKL = nadjiBrojPoVrsti('3', "Ne", listaTESKITRICIKL);
            int brPREKINUTIHREDOVNIHTESKITRICIKL = nadjiBrojPoVrsti('1', "Prekid", listaTESKITRICIKL);
            int brPREKINUTIHSESTOMESECNIHTESKITRICIKL = nadjiBrojPoVrsti('2', "Prekid", listaTESKITRICIKL);
            int brPREKINUTIHVANREDOVNIHTESKITRICIKL = nadjiBrojPoVrsti('4', "Prekid", listaTESKITRICIKL);
            // TESKITRICIKL - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHTESKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaTESKITRICIKL);
            int brNEISPRAVNIHREDOVNIHTESKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaTESKITRICIKL);
            int brISPRAVNIHSESTOMESECNIHTESKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaTESKITRICIKL);
            int brNEISPRAVNIHSESTOMESECNIHTESKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaTESKITRICIKL);
            int brISPRAVNIHVANREDOVNIHTESKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaTESKITRICIKL);
            int brNEISPRAVNIHVANREDOVNIHTESKITRICIKLPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaTESKITRICIKL);
            #endregion
            #region LAKICETVOROCIKL
            int brISPRAVNIHREDOVNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('1', "Da", listaLAKICETVOROCIKL);
            int brNEISPRAVNIHREDOVNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('1', "Ne", listaLAKICETVOROCIKL);
            int brISPRAVNIHSESTOMESECNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('2', "Da", listaLAKICETVOROCIKL);
            int brNEISPRAVNIHSESTOMESECNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('2', "Ne", listaLAKICETVOROCIKL);
            int brISPRAVNIHVANREDOVNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('4', "Da", listaLAKICETVOROCIKL);
            int brNEISPRAVNIHVANREDOVNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('4', "Ne", listaLAKICETVOROCIKL);
            int brISPRAVNIHKONTROLNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('3', "Da", listaLAKICETVOROCIKL);
            int brNEISPRAVNIHKONTROLNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('3', "Ne", listaLAKICETVOROCIKL);
            int brPREKINUTIHREDOVNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('1', "Prekid", listaLAKICETVOROCIKL);
            int brPREKINUTIHSESTOMESECNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('2', "Prekid", listaLAKICETVOROCIKL);
            int brPREKINUTIHVANREDOVNIHLAKICETVOROCIKL = nadjiBrojPoVrsti('4', "Prekid", listaLAKICETVOROCIKL);
            // LAKICETVOROCIKL - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHLAKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaLAKICETVOROCIKL);
            int brNEISPRAVNIHREDOVNIHLAKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaLAKICETVOROCIKL);
            int brISPRAVNIHSESTOMESECNIHLAKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaLAKICETVOROCIKL);
            int brNEISPRAVNIHSESTOMESECNIHLAKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaLAKICETVOROCIKL);
            int brISPRAVNIHVANREDOVNIHLAKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaLAKICETVOROCIKL);
            int brNEISPRAVNIHVANREDOVNIHLAKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaLAKICETVOROCIKL);
            #endregion
            #region TESKICETVOROCIKL
            int brISPRAVNIHREDOVNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('1', "Da", listaTESKICETVOROCIKL);
            int brNEISPRAVNIHREDOVNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('1', "Ne", listaTESKICETVOROCIKL);
            int brISPRAVNIHSESTOMESECNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('2', "Da", listaTESKICETVOROCIKL);
            int brNEISPRAVNIHSESTOMESECNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('2', "Ne", listaTESKICETVOROCIKL);
            int brISPRAVNIHVANREDOVNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('4', "Da", listaTESKICETVOROCIKL);
            int brNEISPRAVNIHVANREDOVNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('4', "Ne", listaTESKICETVOROCIKL);
            int brISPRAVNIHKONTROLNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('3', "Da", listaTESKICETVOROCIKL);
            int brNEISPRAVNIHKONTROLNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('3', "Ne", listaTESKICETVOROCIKL);
            int brPREKINUTIHREDOVNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('1', "Prekid", listaTESKICETVOROCIKL);
            int brPREKINUTIHSESTOMESECNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('2', "Prekid", listaTESKICETVOROCIKL);
            int brPREKINUTIHVANREDOVNIHTESKICETVOROCIKL = nadjiBrojPoVrsti('4', "Prekid", listaTESKICETVOROCIKL);
            // TESKICETVOROCIKL - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHTESKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaTESKICETVOROCIKL);
            int brNEISPRAVNIHREDOVNIHTESKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaTESKICETVOROCIKL);
            int brISPRAVNIHSESTOMESECNIHTESKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaTESKICETVOROCIKL);
            int brNEISPRAVNIHSESTOMESECNIHTESKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaTESKICETVOROCIKL);
            int brISPRAVNIHVANREDOVNIHTESKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaTESKICETVOROCIKL);
            int brNEISPRAVNIHVANREDOVNIHTESKICETVOROCIKLPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaTESKICETVOROCIKL);
            #endregion
            #region PUTNICKO VOZILO
            int brISPRAVNIHREDOVNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('1', "Da", listaPUTNICKOVOZILO);
            int brNEISPRAVNIHREDOVNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('1', "Ne", listaPUTNICKOVOZILO);
            int brISPRAVNIHSESTOMESECNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('2', "Da", listaPUTNICKOVOZILO);
            int brNEISPRAVNIHSESTOMESECNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('2', "Ne", listaPUTNICKOVOZILO);
            int brISPRAVNIHVANREDOVNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('4', "Da", listaPUTNICKOVOZILO);
            int brNEISPRAVNIHVANREDOVNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('4', "Ne", listaPUTNICKOVOZILO);
            int brISPRAVNIHKONTROLNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('3', "Da", listaPUTNICKOVOZILO);
            int brNEISPRAVNIHKONTROLNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('3', "Ne", listaPUTNICKOVOZILO);
            int brPREKINUTIHREDOVNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('1', "Prekid", listaPUTNICKOVOZILO);
            int brPREKINUTIHSESTOMESECNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('2', "Prekid", listaPUTNICKOVOZILO);
            int brPREKINUTIHVANREDOVNIHPUTNICKOVOZILO = nadjiBrojPoVrsti('4', "Prekid", listaPUTNICKOVOZILO);
            // PUTNICKOVOZILO - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHPUTNICKOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaPUTNICKOVOZILO);
            int brNEISPRAVNIHREDOVNIHPUTNICKOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaPUTNICKOVOZILO);
            int brISPRAVNIHSESTOMESECNIHPUTNICKOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaPUTNICKOVOZILO);
            int brNEISPRAVNIHSESTOMESECNIHPUTNICKOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaPUTNICKOVOZILO);
            int brISPRAVNIHVANREDOVNIHPUTNICKOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaPUTNICKOVOZILO);
            int brNEISPRAVNIHVANREDOVNIHPUTNICKOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaPUTNICKOVOZILO);
            #endregion
            #region AUTOBUS
            int brISPRAVNIHREDOVNIHAUTOBUS = nadjiBrojPoVrsti('1', "Da", listaAUTOBUS);
            int brNEISPRAVNIHREDOVNIHAUTOBUS = nadjiBrojPoVrsti('1', "Ne", listaAUTOBUS);
            int brISPRAVNIHSESTOMESECNIHAUTOBUS = nadjiBrojPoVrsti('2', "Da", listaAUTOBUS);
            int brNEISPRAVNIHSESTOMESECNIHAUTOBUS = nadjiBrojPoVrsti('2', "Ne", listaAUTOBUS);
            int brISPRAVNIHVANREDOVNIHAUTOBUS = nadjiBrojPoVrsti('4', "Da", listaAUTOBUS);
            int brNEISPRAVNIHVANREDOVNIHAUTOBUS = nadjiBrojPoVrsti('4', "Ne", listaAUTOBUS);
            int brISPRAVNIHKONTROLNIHAUTOBUS = nadjiBrojPoVrsti('3', "Da", listaAUTOBUS);
            int brNEISPRAVNIHKONTROLNIHAUTOBUS = nadjiBrojPoVrsti('3', "Ne", listaAUTOBUS);
            int brPREKINUTIHREDOVNIHAUTOBUS = nadjiBrojPoVrsti('1', "Prekid", listaAUTOBUS);
            int brPREKINUTIHSESTOMESECNIHAUTOBUS = nadjiBrojPoVrsti('2', "Prekid", listaAUTOBUS);
            int brPREKINUTIHVANREDOVNIHAUTOBUS = nadjiBrojPoVrsti('4', "Prekid", listaAUTOBUS);
            // AUTOBUS - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHAUTOBUSPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaAUTOBUS);
            int brNEISPRAVNIHREDOVNIHAUTOBUSPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaAUTOBUS);
            int brISPRAVNIHSESTOMESECNIHAUTOBUSPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaAUTOBUS);
            int brNEISPRAVNIHSESTOMESECNIHAUTOBUSPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaAUTOBUS);
            int brISPRAVNIHVANREDOVNIHAUTOBUSPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaAUTOBUS);
            int brNEISPRAVNIHVANREDOVNIHAUTOBUSPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaAUTOBUS);
            #endregion
            #region TERETNOVOZILO
            int brISPRAVNIHREDOVNIHTERETNOVOZILO = nadjiBrojPoVrsti('1', "Da", listaTERETNOVOZILO);
            int brNEISPRAVNIHREDOVNIHTERETNOVOZILO = nadjiBrojPoVrsti('1', "Ne", listaTERETNOVOZILO);
            int brISPRAVNIHSESTOMESECNIHTERETNOVOZILO = nadjiBrojPoVrsti('2', "Da", listaTERETNOVOZILO);
            int brNEISPRAVNIHSESTOMESECNIHTERETNOVOZILO = nadjiBrojPoVrsti('2', "Ne", listaTERETNOVOZILO);
            int brISPRAVNIHVANREDOVNIHTERETNOVOZILO = nadjiBrojPoVrsti('4', "Da", listaTERETNOVOZILO);
            int brNEISPRAVNIHVANREDOVNIHTERETNOVOZILO = nadjiBrojPoVrsti('4', "Ne", listaTERETNOVOZILO);
            int brISPRAVNIHKONTROLNIHTERETNOVOZILO = nadjiBrojPoVrsti('3', "Da", listaTERETNOVOZILO);
            int brNEISPRAVNIHKONTROLNIHTERETNOVOZILO = nadjiBrojPoVrsti('3', "Ne", listaTERETNOVOZILO);
            int brPREKINUTIHREDOVNIHTERETNOVOZILO = nadjiBrojPoVrsti('1', "Prekid", listaTERETNOVOZILO);
            int brPREKINUTIHSESTOMESECNIHTERETNOVOZILO = nadjiBrojPoVrsti('2', "Prekid", listaTERETNOVOZILO);
            int brPREKINUTIHVANREDOVNIHTERETNOVOZILO = nadjiBrojPoVrsti('4', "Prekid", listaTERETNOVOZILO);
            // TERETNOVOZILO - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHTERETNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaTERETNOVOZILO);
            int brNEISPRAVNIHREDOVNIHTERETNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaTERETNOVOZILO);
            int brISPRAVNIHSESTOMESECNIHTERETNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaTERETNOVOZILO);
            int brNEISPRAVNIHSESTOMESECNIHTERETNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaTERETNOVOZILO);
            int brISPRAVNIHVANREDOVNIHTERETNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaTERETNOVOZILO);
            int brNEISPRAVNIHVANREDOVNIHTERETNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaTERETNOVOZILO);
            #endregion
            #region RADNAMASINA
            int brISPRAVNIHREDOVNIHRADNAMASINA = nadjiBrojPoVrsti('1', "Da", listaRADNAMASINA);
            int brNEISPRAVNIHREDOVNIHRADNAMASINA = nadjiBrojPoVrsti('1', "Ne", listaRADNAMASINA);
            int brISPRAVNIHSESTOMESECNIHRADNAMASINA = nadjiBrojPoVrsti('2', "Da", listaRADNAMASINA);
            int brNEISPRAVNIHSESTOMESECNIHRADNAMASINA = nadjiBrojPoVrsti('2', "Ne", listaRADNAMASINA);
            int brISPRAVNIHVANREDOVNIHRADNAMASINA = nadjiBrojPoVrsti('4', "Da", listaRADNAMASINA);
            int brNEISPRAVNIHVANREDOVNIHRADNAMASINA = nadjiBrojPoVrsti('4', "Ne", listaRADNAMASINA);
            int brISPRAVNIHKONTROLNIHRADNAMASINA = nadjiBrojPoVrsti('3', "Da", listaRADNAMASINA);
            int brNEISPRAVNIHKONTROLNIHRADNAMASINA = nadjiBrojPoVrsti('3', "Ne", listaRADNAMASINA);
            int brPREKINUTIHREDOVNIHRADNAMASINA = nadjiBrojPoVrsti('1', "Prekid", listaRADNAMASINA);
            int brPREKINUTIHSESTOMESECNIHRADNAMASINA = nadjiBrojPoVrsti('2', "Prekid", listaRADNAMASINA);
            int brPREKINUTIHVANREDOVNIHRADNAMASINA = nadjiBrojPoVrsti('4', "Prekid", listaRADNAMASINA);
            // RADNAMASINA - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHRADNAMASINAPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaRADNAMASINA);
            int brNEISPRAVNIHREDOVNIHRADNAMASINAPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaRADNAMASINA);
            int brISPRAVNIHSESTOMESECNIHRADNAMASINAPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaRADNAMASINA);
            int brNEISPRAVNIHSESTOMESECNIHRADNAMASINAPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaRADNAMASINA);
            int brISPRAVNIHVANREDOVNIHRADNAMASINAPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaRADNAMASINA);
            int brNEISPRAVNIHVANREDOVNIHRADNAMASINAPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaRADNAMASINA);
            #endregion
            #region TRAKTOR
            int brISPRAVNIHREDOVNIHTRAKTOR = nadjiBrojPoVrsti('1', "Da", listaTRAKTOR);
            int brNEISPRAVNIHREDOVNIHTRAKTOR = nadjiBrojPoVrsti('1', "Ne", listaTRAKTOR);
            int brISPRAVNIHSESTOMESECNIHTRAKTOR = nadjiBrojPoVrsti('2', "Da", listaTRAKTOR);
            int brNEISPRAVNIHSESTOMESECNIHTRAKTOR = nadjiBrojPoVrsti('2', "Ne", listaTRAKTOR);
            int brISPRAVNIHVANREDOVNIHTRAKTOR = nadjiBrojPoVrsti('4', "Da", listaTRAKTOR);
            int brNEISPRAVNIHVANREDOVNIHTRAKTOR = nadjiBrojPoVrsti('4', "Ne", listaTRAKTOR);
            int brISPRAVNIHKONTROLNIHTRAKTOR = nadjiBrojPoVrsti('3', "Da", listaTRAKTOR);
            int brNEISPRAVNIHKONTROLNIHTRAKTOR = nadjiBrojPoVrsti('3', "Ne", listaTRAKTOR);
            int brPREKINUTIHREDOVNIHTRAKTOR = nadjiBrojPoVrsti('1', "Prekid", listaTRAKTOR);
            int brPREKINUTIHSESTOMESECNIHTRAKTOR = nadjiBrojPoVrsti('2', "Prekid", listaTRAKTOR);
            int brPREKINUTIHVANREDOVNIHTRAKTOR = nadjiBrojPoVrsti('4', "Prekid", listaTRAKTOR);
            // TRAKTOR - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHTRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaTRAKTOR);
            int brNEISPRAVNIHREDOVNIHTRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaTRAKTOR);
            int brISPRAVNIHSESTOMESECNIHTRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaTRAKTOR);
            int brNEISPRAVNIHSESTOMESECNIHTRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaTRAKTOR);
            int brISPRAVNIHVANREDOVNIHTRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaTRAKTOR);
            int brNEISPRAVNIHVANREDOVNIHTRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaTRAKTOR);
            #endregion
            #region MOTOKULTIVATOR
            int brISPRAVNIHREDOVNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('1', "Da", listaMOTOKULTIVATOR);
            int brNEISPRAVNIHREDOVNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('1', "Ne", listaMOTOKULTIVATOR);
            int brISPRAVNIHSESTOMESECNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('2', "Da", listaMOTOKULTIVATOR);
            int brNEISPRAVNIHSESTOMESECNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('2', "Ne", listaMOTOKULTIVATOR);
            int brISPRAVNIHVANREDOVNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('4', "Da", listaMOTOKULTIVATOR);
            int brNEISPRAVNIHVANREDOVNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('4', "Ne", listaMOTOKULTIVATOR);
            int brISPRAVNIHKONTROLNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('3', "Da", listaMOTOKULTIVATOR);
            int brNEISPRAVNIHKONTROLNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('3', "Ne", listaMOTOKULTIVATOR);
            int brPREKINUTIHREDOVNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('1', "Prekid", listaMOTOKULTIVATOR);
            int brPREKINUTIHSESTOMESECNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('2', "Prekid", listaMOTOKULTIVATOR);
            int brPREKINUTIHVANREDOVNIHMOTOKULTIVATOR = nadjiBrojPoVrsti('4', "Prekid", listaMOTOKULTIVATOR);
            // MOTOKULTIVATOR - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHMOTOKULTIVATORPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaMOTOKULTIVATOR);
            int brNEISPRAVNIHREDOVNIHMOTOKULTIVATORPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaMOTOKULTIVATOR);
            int brISPRAVNIHSESTOMESECNIHMOTOKULTIVATORPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaMOTOKULTIVATOR);
            int brNEISPRAVNIHSESTOMESECNIHMOTOKULTIVATORPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaMOTOKULTIVATOR);
            int brISPRAVNIHVANREDOVNIHMOTOKULTIVATORPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaMOTOKULTIVATOR);
            int brNEISPRAVNIHVANREDOVNIHMOTOKULTIVATORPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaMOTOKULTIVATOR);
            #endregion
            #region PRIKLJUCNOVOZILO
            int brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('1', "Da", listaPRIKLJUCNOVOZILO);
            int brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('1', "Ne", listaPRIKLJUCNOVOZILO);
            int brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('2', "Da", listaPRIKLJUCNOVOZILO);
            int brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('2', "Ne", listaPRIKLJUCNOVOZILO);
            int brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('4', "Da", listaPRIKLJUCNOVOZILO);
            int brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('4', "Ne", listaPRIKLJUCNOVOZILO);
            int brISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('3', "Da", listaPRIKLJUCNOVOZILO);
            int brNEISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('3', "Ne", listaPRIKLJUCNOVOZILO);
            int brPREKINUTIHREDOVNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('1', "Prekid", listaPRIKLJUCNOVOZILO);
            int brPREKINUTIHSESTOMESECNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('2', "Prekid", listaPRIKLJUCNOVOZILO);
            int brPREKINUTIHVANREDOVNIHPRIKLJUCNOVOZILO = nadjiBrojPoVrsti('4', "Prekid", listaPRIKLJUCNOVOZILO);
            // PRIKLJUCNOVOZILO - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaPRIKLJUCNOVOZILO);
            int brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaPRIKLJUCNOVOZILO);
            int brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaPRIKLJUCNOVOZILO);
            int brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaPRIKLJUCNOVOZILO);
            int brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaPRIKLJUCNOVOZILO);
            int brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaPRIKLJUCNOVOZILO);
            #endregion
            #region PRIKLJUCNOVOZILOZATRAKTOR
            int brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR = nadjiBrojPoVrsti('1', "Da", listaPRIKLJUCNOZATRAKTOR);
            int brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR = nadjiBrojPoVrsti('1', "Ne", listaPRIKLJUCNOZATRAKTOR);
            int brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTOR = nadjiBrojPoVrsti('2', "Da", listaPRIKLJUCNOZATRAKTOR);
            int brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTOR = nadjiBrojPoVrsti('2', "Ne", listaPRIKLJUCNOZATRAKTOR);
            int brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR = nadjiBrojPoVrsti('4', "Da", listaPRIKLJUCNOZATRAKTOR);
            int brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR = nadjiBrojPoVrsti('4', "Ne", listaPRIKLJUCNOZATRAKTOR);
            int brISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILOZATRAKTOR = nadjiBrojPoVrsti('3', "Da", listaPRIKLJUCNOZATRAKTOR);
            int brNEISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILOZATRAKTOR = nadjiBrojPoVrsti('3', "Ne", listaPRIKLJUCNOZATRAKTOR);
            int brPREKINUTIHREDOVNIHPRIKLJUCNOZATRAKTOR = nadjiBrojPoVrsti('1', "Prekid", listaPRIKLJUCNOZATRAKTOR);
            int brPREKINUTIHSESTOMESECNIHPRIKLJUCNOZATRAKTOR = nadjiBrojPoVrsti('2', "Prekid", listaPRIKLJUCNOZATRAKTOR);
            int brPREKINUTIHVANREDOVNIHPRIKLJUCNOZATRAKTOR = nadjiBrojPoVrsti('4', "Prekid", listaPRIKLJUCNOZATRAKTOR);
            // PRIKLJUCNOVOZILOZATRAKTOR - PONOVO DOSLI
            int brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('1', "Da", listaPRIKLJUCNOZATRAKTOR);
            int brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('1', "Ne", listaPRIKLJUCNOZATRAKTOR);
            int brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('2', "Da", listaPRIKLJUCNOZATRAKTOR);
            int brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('2', "Ne", listaPRIKLJUCNOZATRAKTOR);
            int brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('4', "Da", listaPRIKLJUCNOZATRAKTOR);
            int brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO = nadjiBrojPoVrstiPonovnih('4', "Ne", listaPRIKLJUCNOZATRAKTOR);
            #endregion

            Excel.Application xlApp = new Excel.Application();
            string sFileName = Application.StartupPath + "\\BS-13-1.xlsx";
            Excel.Workbook xlBook = (Excel.Workbook)xlApp.Workbooks.Open(sFileName);
            Excel.Worksheet xlSheet = (Excel.Worksheet)xlBook.Worksheets[1];

            #region UKUPNOMOPED
            if (brPREKINUTIHREDOVNIHMOPED > 0)
                xlSheet.Range["D6"].Value = brPREKINUTIHREDOVNIHMOPED;
            if (brISPRAVNIHREDOVNIHMOPED > 0)
                xlSheet.Range["E6"].Value = brISPRAVNIHREDOVNIHMOPED;
            if (brNEISPRAVNIHREDOVNIHMOPED > 0)
                xlSheet.Range["F6"].Value = brNEISPRAVNIHREDOVNIHMOPED;
            if (brISPRAVNIHREDOVNIHMOPEDPONOVO > 0)
                xlSheet.Range["G6"].Value = brISPRAVNIHREDOVNIHMOPEDPONOVO;
            if (brNEISPRAVNIHREDOVNIHMOPEDPONOVO > 0)
                xlSheet.Range["H6"].Value = brNEISPRAVNIHREDOVNIHMOPEDPONOVO;
            int UKUPNOREDOVNIMOPED = brPREKINUTIHREDOVNIHMOPED + brISPRAVNIHREDOVNIHMOPED + brNEISPRAVNIHREDOVNIHMOPED + brISPRAVNIHREDOVNIHMOPEDPONOVO + brNEISPRAVNIHREDOVNIHMOPEDPONOVO;
            if (UKUPNOREDOVNIMOPED > 0)
                xlSheet.Range["C6"].Value = UKUPNOREDOVNIMOPED;
            if (brPREKINUTIHSESTOMESECNIHMOPED > 0)
                xlSheet.Range["J6"].Value = brPREKINUTIHSESTOMESECNIHMOPED;
            if (brISPRAVNIHSESTOMESECNIHMOPED > 0)
                xlSheet.Range["K6"].Value = brISPRAVNIHSESTOMESECNIHMOPED;
            if (brNEISPRAVNIHSESTOMESECNIHMOPED > 0)
                xlSheet.Range["L6"].Value = brNEISPRAVNIHSESTOMESECNIHMOPED;
            if (brISPRAVNIHSESTOMESECNIHMOPEDPONOVO > 0)
                xlSheet.Range["M6"].Value = brISPRAVNIHSESTOMESECNIHMOPEDPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHMOPEDPONOVO > 0)
                xlSheet.Range["N6"].Value = brNEISPRAVNIHSESTOMESECNIHMOPEDPONOVO;
            int UKUPNOSESTOMESECNIMOPED = brPREKINUTIHSESTOMESECNIHMOPED + brISPRAVNIHSESTOMESECNIHMOPED + brNEISPRAVNIHSESTOMESECNIHMOPED + brISPRAVNIHSESTOMESECNIHMOPEDPONOVO + brNEISPRAVNIHSESTOMESECNIHMOPEDPONOVO;
            if (UKUPNOSESTOMESECNIMOPED > 0)
                xlSheet.Range["I6"].Value = UKUPNOSESTOMESECNIMOPED;
            if (brPREKINUTIHVANREDOVNIHMOPED > 0)
                xlSheet.Range["P6"].Value = brPREKINUTIHVANREDOVNIHMOPED;
            if (brISPRAVNIHVANREDOVNIHMOPED > 0)
                xlSheet.Range["Q6"].Value = brISPRAVNIHVANREDOVNIHMOPED;
            if (brNEISPRAVNIHVANREDOVNIHMOPED > 0)
                xlSheet.Range["R6"].Value = brNEISPRAVNIHVANREDOVNIHMOPED;
            if (brISPRAVNIHVANREDOVNIHMOPEDPONOVO > 0)
                xlSheet.Range["S6"].Value = brISPRAVNIHVANREDOVNIHMOPEDPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHMOPEDPONOVO > 0)
                xlSheet.Range["T6"].Value = brNEISPRAVNIHVANREDOVNIHMOPEDPONOVO;
            int UKUPNOVANREDNIMOPED = brPREKINUTIHVANREDOVNIHMOPED + brISPRAVNIHVANREDOVNIHMOPED + brNEISPRAVNIHVANREDOVNIHMOPED + brISPRAVNIHVANREDOVNIHMOPEDPONOVO + brNEISPRAVNIHVANREDOVNIHMOPEDPONOVO;
            if (UKUPNOVANREDNIMOPED > 0)
                xlSheet.Range["O6"].Value = UKUPNOVANREDNIMOPED;
            if (brISPRAVNIHKONTROLNIHMOPED > 0)
                xlSheet.Range["V6"].Value = brISPRAVNIHKONTROLNIHMOPED;
            if (brNEISPRAVNIHKONTROLNIHMOPED > 0)
                xlSheet.Range["W6"].Value = brNEISPRAVNIHKONTROLNIHMOPED;
            int UKUPNOKONTROLNIMOPED = brISPRAVNIHKONTROLNIHMOPED + brNEISPRAVNIHKONTROLNIHMOPED;
            if (UKUPNOKONTROLNIMOPED > 0)
                xlSheet.Range["U6"].Value = UKUPNOKONTROLNIMOPED;
            int UKUPNOMOPED = UKUPNOREDOVNIMOPED + UKUPNOSESTOMESECNIMOPED + UKUPNOVANREDNIMOPED + UKUPNOKONTROLNIMOPED;
            if (UKUPNOMOPED > 0)
                xlSheet.Range["B6"].Value = UKUPNOMOPED;
            #endregion
            #region UKUPNOLAKITRICIKL
            if (brPREKINUTIHREDOVNIHLAKITRICIKL > 0)
                xlSheet.Range["D7"].Value = brPREKINUTIHREDOVNIHLAKITRICIKL;
            if (brISPRAVNIHREDOVNIHLAKITRICIKL > 0)
                xlSheet.Range["E7"].Value = brISPRAVNIHREDOVNIHLAKITRICIKL;
            if (brNEISPRAVNIHREDOVNIHLAKITRICIKL > 0)
                xlSheet.Range["F7"].Value = brNEISPRAVNIHREDOVNIHLAKITRICIKL;
            if (brISPRAVNIHREDOVNIHLAKITRICIKLPONOVO > 0)
                xlSheet.Range["G7"].Value = brISPRAVNIHREDOVNIHLAKITRICIKLPONOVO;
            if (brNEISPRAVNIHREDOVNIHLAKITRICIKLPONOVO > 0)
                xlSheet.Range["H7"].Value = brNEISPRAVNIHREDOVNIHLAKITRICIKLPONOVO;
            int UKUPNOREDOVNILAKITRICIKL = brPREKINUTIHREDOVNIHLAKITRICIKL + brISPRAVNIHREDOVNIHLAKITRICIKL + brNEISPRAVNIHREDOVNIHLAKITRICIKL + brISPRAVNIHREDOVNIHLAKITRICIKLPONOVO + brNEISPRAVNIHREDOVNIHLAKITRICIKLPONOVO;
            if (UKUPNOREDOVNILAKITRICIKL > 0)
                xlSheet.Range["C7"].Value = UKUPNOREDOVNILAKITRICIKL;
            if (brPREKINUTIHSESTOMESECNIHLAKITRICIKL > 0)
                xlSheet.Range["J7"].Value = brPREKINUTIHSESTOMESECNIHLAKITRICIKL;
            if (brISPRAVNIHSESTOMESECNIHLAKITRICIKL > 0)
                xlSheet.Range["K7"].Value = brISPRAVNIHSESTOMESECNIHLAKITRICIKL;
            if (brNEISPRAVNIHSESTOMESECNIHLAKITRICIKL > 0)
                xlSheet.Range["L7"].Value = brNEISPRAVNIHSESTOMESECNIHLAKITRICIKL;
            if (brISPRAVNIHSESTOMESECNIHLAKITRICIKLPONOVO > 0)
                xlSheet.Range["M7"].Value = brISPRAVNIHSESTOMESECNIHLAKITRICIKLPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHLAKITRICIKLPONOVO > 0)
                xlSheet.Range["N7"].Value = brNEISPRAVNIHSESTOMESECNIHLAKITRICIKLPONOVO;
            int UKUPNOSESTOMESECNILAKITRICIKL = brPREKINUTIHSESTOMESECNIHLAKITRICIKL + brISPRAVNIHSESTOMESECNIHLAKITRICIKL + brNEISPRAVNIHSESTOMESECNIHLAKITRICIKL + brISPRAVNIHSESTOMESECNIHLAKITRICIKLPONOVO + brNEISPRAVNIHSESTOMESECNIHLAKITRICIKLPONOVO;
            if (UKUPNOSESTOMESECNILAKITRICIKL > 0)
                xlSheet.Range["I7"].Value = UKUPNOSESTOMESECNILAKITRICIKL;
            if (brPREKINUTIHVANREDOVNIHLAKITRICIKL > 0)
                xlSheet.Range["P7"].Value = brPREKINUTIHVANREDOVNIHLAKITRICIKL;
            if (brISPRAVNIHVANREDOVNIHLAKITRICIKL > 0)
                xlSheet.Range["Q7"].Value = brISPRAVNIHVANREDOVNIHLAKITRICIKL;
            if (brNEISPRAVNIHVANREDOVNIHLAKITRICIKL > 0)
                xlSheet.Range["R7"].Value = brNEISPRAVNIHVANREDOVNIHLAKITRICIKL;
            if (brISPRAVNIHVANREDOVNIHLAKITRICIKLPONOVO > 0)
                xlSheet.Range["S7"].Value = brISPRAVNIHVANREDOVNIHLAKITRICIKLPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHLAKITRICIKLPONOVO > 0)
                xlSheet.Range["T7"].Value = brNEISPRAVNIHVANREDOVNIHLAKITRICIKLPONOVO;
            int UKUPNOVANREDNILAKITRICIKL = brPREKINUTIHVANREDOVNIHLAKITRICIKL + brISPRAVNIHVANREDOVNIHLAKITRICIKL + brNEISPRAVNIHVANREDOVNIHLAKITRICIKL + brISPRAVNIHVANREDOVNIHLAKITRICIKLPONOVO + brNEISPRAVNIHVANREDOVNIHLAKITRICIKLPONOVO;
            if (UKUPNOVANREDNILAKITRICIKL > 0)
                xlSheet.Range["O7"].Value = UKUPNOVANREDNILAKITRICIKL;
            if (brISPRAVNIHKONTROLNIHLAKITRICIKL > 0)
                xlSheet.Range["V7"].Value = brISPRAVNIHKONTROLNIHLAKITRICIKL;
            if (brNEISPRAVNIHKONTROLNIHLAKITRICIKL > 0)
                xlSheet.Range["W7"].Value = brNEISPRAVNIHKONTROLNIHLAKITRICIKL;
            int UKUPNOKONTROLNILAKITRICIKL = brISPRAVNIHKONTROLNIHLAKITRICIKL + brNEISPRAVNIHKONTROLNIHLAKITRICIKL;
            if (UKUPNOKONTROLNILAKITRICIKL > 0)
                xlSheet.Range["U7"].Value = UKUPNOKONTROLNILAKITRICIKL;
            int UKUPNOLAKITRICIKL = UKUPNOREDOVNILAKITRICIKL + UKUPNOSESTOMESECNILAKITRICIKL + UKUPNOVANREDNILAKITRICIKL + UKUPNOKONTROLNILAKITRICIKL;
            if (UKUPNOLAKITRICIKL > 0)
                xlSheet.Range["B7"].Value = UKUPNOLAKITRICIKL;
            #endregion
            #region UKUPNOMOTOCIKL
            if (brPREKINUTIHREDOVNIHMOTOCIKL > 0)
                xlSheet.Range["D8"].Value = brPREKINUTIHREDOVNIHMOTOCIKL;
            if (brISPRAVNIHREDOVNIHMOTOCIKL > 0)
                xlSheet.Range["E8"].Value = brISPRAVNIHREDOVNIHMOTOCIKL;
            if (brNEISPRAVNIHREDOVNIHMOTOCIKL > 0)
                xlSheet.Range["F8"].Value = brNEISPRAVNIHREDOVNIHMOTOCIKL;
            if (brISPRAVNIHREDOVNIHMOTOCIKLPONOVO > 0)
                xlSheet.Range["G8"].Value = brISPRAVNIHREDOVNIHMOTOCIKLPONOVO;
            if (brNEISPRAVNIHREDOVNIHMOTOCIKLPONOVO > 0)
                xlSheet.Range["H8"].Value = brNEISPRAVNIHREDOVNIHMOTOCIKLPONOVO;
            int UKUPNOREDOVNIMOTOCIKL = brPREKINUTIHREDOVNIHMOTOCIKL + brISPRAVNIHREDOVNIHMOTOCIKL + brNEISPRAVNIHREDOVNIHMOTOCIKL + brISPRAVNIHREDOVNIHMOTOCIKLPONOVO + brNEISPRAVNIHREDOVNIHMOTOCIKLPONOVO;
            if (UKUPNOREDOVNIMOTOCIKL > 0)
                xlSheet.Range["C8"].Value = UKUPNOREDOVNIMOTOCIKL;
            if (brPREKINUTIHSESTOMESECNIHMOTOCIKL > 0)
                xlSheet.Range["J8"].Value = brPREKINUTIHSESTOMESECNIHMOTOCIKL;
            if (brISPRAVNIHSESTOMESECNIHMOTOCIKL > 0)
                xlSheet.Range["K8"].Value = brISPRAVNIHSESTOMESECNIHMOTOCIKL;
            if (brNEISPRAVNIHSESTOMESECNIHMOTOCIKL > 0)
                xlSheet.Range["L8"].Value = brNEISPRAVNIHSESTOMESECNIHMOTOCIKL;
            if (brISPRAVNIHSESTOMESECNIHMOTOCIKLPONOVO > 0)
                xlSheet.Range["M8"].Value = brISPRAVNIHSESTOMESECNIHMOTOCIKLPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHMOTOCIKLPONOVO > 0)
                xlSheet.Range["N8"].Value = brNEISPRAVNIHSESTOMESECNIHMOTOCIKLPONOVO;
            int UKUPNOSESTOMESECNIMOTOCIKL = brPREKINUTIHSESTOMESECNIHMOTOCIKL + brISPRAVNIHSESTOMESECNIHMOTOCIKL + brNEISPRAVNIHSESTOMESECNIHMOTOCIKL + brISPRAVNIHSESTOMESECNIHMOTOCIKLPONOVO + brNEISPRAVNIHSESTOMESECNIHMOTOCIKLPONOVO;
            if (UKUPNOSESTOMESECNIMOTOCIKL > 0)
                xlSheet.Range["I8"].Value = UKUPNOSESTOMESECNIMOTOCIKL;
            if (brPREKINUTIHVANREDOVNIHMOTOCIKL > 0)
                xlSheet.Range["P8"].Value = brPREKINUTIHVANREDOVNIHMOTOCIKL;
            if (brISPRAVNIHVANREDOVNIHMOTOCIKL > 0)
                xlSheet.Range["Q8"].Value = brISPRAVNIHVANREDOVNIHMOTOCIKL;
            if (brNEISPRAVNIHVANREDOVNIHMOTOCIKL > 0)
                xlSheet.Range["R8"].Value = brNEISPRAVNIHVANREDOVNIHMOTOCIKL;
            if (brISPRAVNIHVANREDOVNIHMOTOCIKLPONOVO > 0)
                xlSheet.Range["S8"].Value = brISPRAVNIHVANREDOVNIHMOTOCIKLPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHMOTOCIKLPONOVO > 0)
                xlSheet.Range["T8"].Value = brNEISPRAVNIHVANREDOVNIHMOTOCIKLPONOVO;
            int UKUPNOVANREDNIMOTOCIKL = brPREKINUTIHVANREDOVNIHMOTOCIKL + brISPRAVNIHVANREDOVNIHMOTOCIKL + brNEISPRAVNIHVANREDOVNIHMOTOCIKL + brISPRAVNIHVANREDOVNIHMOTOCIKLPONOVO + brNEISPRAVNIHVANREDOVNIHMOTOCIKLPONOVO;
            if (UKUPNOVANREDNIMOTOCIKL > 0)
                xlSheet.Range["O8"].Value = UKUPNOVANREDNIMOTOCIKL;
            if (brISPRAVNIHKONTROLNIHMOTOCIKL > 0)
                xlSheet.Range["V8"].Value = brISPRAVNIHKONTROLNIHMOTOCIKL;
            if (brNEISPRAVNIHKONTROLNIHMOTOCIKL > 0)
                xlSheet.Range["W8"].Value = brNEISPRAVNIHKONTROLNIHMOTOCIKL;
            int UKUPNOKONTROLNIMOTOCIKL = brISPRAVNIHKONTROLNIHMOTOCIKL + brNEISPRAVNIHKONTROLNIHMOTOCIKL;
            if (UKUPNOKONTROLNIMOTOCIKL > 0)
                xlSheet.Range["U8"].Value = UKUPNOKONTROLNIMOTOCIKL;
            int UKUPNOMOTOCIKL = UKUPNOREDOVNIMOTOCIKL + UKUPNOSESTOMESECNIMOTOCIKL + UKUPNOVANREDNIMOTOCIKL + UKUPNOKONTROLNIMOTOCIKL;
            if (UKUPNOMOTOCIKL > 0)
                xlSheet.Range["B8"].Value = UKUPNOMOTOCIKL;
            #endregion
            #region UKUPNOTESKITRICIKL
            if (brPREKINUTIHREDOVNIHTESKITRICIKL > 0)
                xlSheet.Range["D9"].Value = brPREKINUTIHREDOVNIHTESKITRICIKL;
            if (brISPRAVNIHREDOVNIHTESKITRICIKL > 0)
                xlSheet.Range["E9"].Value = brISPRAVNIHREDOVNIHTESKITRICIKL;
            if (brNEISPRAVNIHREDOVNIHTESKITRICIKL > 0)
                xlSheet.Range["F9"].Value = brNEISPRAVNIHREDOVNIHTESKITRICIKL;
            if (brISPRAVNIHREDOVNIHTESKITRICIKLPONOVO > 0)
                xlSheet.Range["G9"].Value = brISPRAVNIHREDOVNIHTESKITRICIKLPONOVO;
            if (brNEISPRAVNIHREDOVNIHTESKITRICIKLPONOVO > 0)
                xlSheet.Range["H9"].Value = brNEISPRAVNIHREDOVNIHTESKITRICIKLPONOVO;
            int UKUPNOREDOVNITESKITRICIKL = brPREKINUTIHREDOVNIHTESKITRICIKL + brISPRAVNIHREDOVNIHTESKITRICIKL + brNEISPRAVNIHREDOVNIHTESKITRICIKL + brISPRAVNIHREDOVNIHTESKITRICIKLPONOVO + brNEISPRAVNIHREDOVNIHTESKITRICIKLPONOVO;
            if (UKUPNOREDOVNITESKITRICIKL > 0)
                xlSheet.Range["C9"].Value = UKUPNOREDOVNITESKITRICIKL;
            if (brPREKINUTIHSESTOMESECNIHTESKITRICIKL > 0)
                xlSheet.Range["J9"].Value = brPREKINUTIHSESTOMESECNIHTESKITRICIKL;
            if (brISPRAVNIHSESTOMESECNIHTESKITRICIKL > 0)
                xlSheet.Range["K9"].Value = brISPRAVNIHSESTOMESECNIHTESKITRICIKL;
            if (brNEISPRAVNIHSESTOMESECNIHTESKITRICIKL > 0)
                xlSheet.Range["L9"].Value = brNEISPRAVNIHSESTOMESECNIHTESKITRICIKL;
            if (brISPRAVNIHSESTOMESECNIHTESKITRICIKLPONOVO > 0)
                xlSheet.Range["M9"].Value = brISPRAVNIHSESTOMESECNIHTESKITRICIKLPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHTESKITRICIKLPONOVO > 0)
                xlSheet.Range["N9"].Value = brNEISPRAVNIHSESTOMESECNIHTESKITRICIKLPONOVO;
            int UKUPNOSESTOMESECNITESKITRICIKL = brPREKINUTIHSESTOMESECNIHTESKITRICIKL + brISPRAVNIHSESTOMESECNIHTESKITRICIKL + brNEISPRAVNIHSESTOMESECNIHTESKITRICIKL + brISPRAVNIHSESTOMESECNIHTESKITRICIKLPONOVO + brNEISPRAVNIHSESTOMESECNIHTESKITRICIKLPONOVO;
            if (UKUPNOSESTOMESECNITESKITRICIKL > 0)
                xlSheet.Range["I9"].Value = UKUPNOSESTOMESECNITESKITRICIKL;
            if (brPREKINUTIHVANREDOVNIHTESKITRICIKL > 0)
                xlSheet.Range["P9"].Value = brPREKINUTIHVANREDOVNIHTESKITRICIKL;
            if (brISPRAVNIHVANREDOVNIHTESKITRICIKL > 0)
                xlSheet.Range["Q9"].Value = brISPRAVNIHVANREDOVNIHTESKITRICIKL;
            if (brNEISPRAVNIHVANREDOVNIHTESKITRICIKL > 0)
                xlSheet.Range["R9"].Value = brNEISPRAVNIHVANREDOVNIHTESKITRICIKL;
            if (brISPRAVNIHVANREDOVNIHTESKITRICIKLPONOVO > 0)
                xlSheet.Range["S9"].Value = brISPRAVNIHVANREDOVNIHTESKITRICIKLPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHTESKITRICIKLPONOVO > 0)
                xlSheet.Range["T9"].Value = brNEISPRAVNIHVANREDOVNIHTESKITRICIKLPONOVO;
            int UKUPNOVANREDNITESKITRICIKL = brPREKINUTIHVANREDOVNIHTESKITRICIKL + brISPRAVNIHVANREDOVNIHTESKITRICIKL + brNEISPRAVNIHVANREDOVNIHTESKITRICIKL + brISPRAVNIHVANREDOVNIHTESKITRICIKLPONOVO + brNEISPRAVNIHVANREDOVNIHTESKITRICIKLPONOVO;
            if (UKUPNOVANREDNITESKITRICIKL > 0)
                xlSheet.Range["O9"].Value = UKUPNOVANREDNITESKITRICIKL;
            if (brISPRAVNIHKONTROLNIHTESKITRICIKL > 0)
                xlSheet.Range["V9"].Value = brISPRAVNIHKONTROLNIHTESKITRICIKL;
            if (brNEISPRAVNIHKONTROLNIHTESKITRICIKL > 0)
                xlSheet.Range["W9"].Value = brNEISPRAVNIHKONTROLNIHTESKITRICIKL;
            int UKUPNOKONTROLNITESKITRICIKL = brISPRAVNIHKONTROLNIHTESKITRICIKL + brNEISPRAVNIHKONTROLNIHTESKITRICIKL;
            if (UKUPNOKONTROLNITESKITRICIKL > 0)
                xlSheet.Range["U9"].Value = UKUPNOKONTROLNITESKITRICIKL;
            int UKUPNOTESKITRICIKL = UKUPNOREDOVNITESKITRICIKL + UKUPNOSESTOMESECNITESKITRICIKL + UKUPNOVANREDNITESKITRICIKL + UKUPNOKONTROLNITESKITRICIKL;
            if (UKUPNOTESKITRICIKL > 0)
                xlSheet.Range["B9"].Value = UKUPNOTESKITRICIKL;
            #endregion
            #region UKUPNOLAKICETVOROCIKL
            if (brPREKINUTIHREDOVNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["D10"].Value = brPREKINUTIHREDOVNIHLAKICETVOROCIKL;
            if (brISPRAVNIHREDOVNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["E10"].Value = brISPRAVNIHREDOVNIHLAKICETVOROCIKL;
            if (brNEISPRAVNIHREDOVNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["F10"].Value = brNEISPRAVNIHREDOVNIHLAKICETVOROCIKL;
            if (brISPRAVNIHREDOVNIHLAKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["G10"].Value = brISPRAVNIHREDOVNIHLAKICETVOROCIKLPONOVO;
            if (brNEISPRAVNIHREDOVNIHLAKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["H10"].Value = brNEISPRAVNIHREDOVNIHLAKICETVOROCIKLPONOVO;
            int UKUPNOREDOVNILAKICETVOROCIKL = brPREKINUTIHREDOVNIHLAKICETVOROCIKL + brISPRAVNIHREDOVNIHLAKICETVOROCIKL + brNEISPRAVNIHREDOVNIHLAKICETVOROCIKL + brISPRAVNIHREDOVNIHLAKICETVOROCIKLPONOVO + brNEISPRAVNIHREDOVNIHLAKICETVOROCIKLPONOVO;
            if (UKUPNOREDOVNILAKICETVOROCIKL > 0)
                xlSheet.Range["C10"].Value = UKUPNOREDOVNILAKICETVOROCIKL;
            if (brPREKINUTIHSESTOMESECNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["J10"].Value = brPREKINUTIHSESTOMESECNIHLAKICETVOROCIKL;
            if (brISPRAVNIHSESTOMESECNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["K10"].Value = brISPRAVNIHSESTOMESECNIHLAKICETVOROCIKL;
            if (brNEISPRAVNIHSESTOMESECNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["L10"].Value = brNEISPRAVNIHSESTOMESECNIHLAKICETVOROCIKL;
            if (brISPRAVNIHSESTOMESECNIHLAKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["M10"].Value = brISPRAVNIHSESTOMESECNIHLAKICETVOROCIKLPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHLAKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["N10"].Value = brNEISPRAVNIHSESTOMESECNIHLAKICETVOROCIKLPONOVO;
            int UKUPNOSESTOMESECNILAKICETVOROCIKL = brPREKINUTIHSESTOMESECNIHLAKICETVOROCIKL + brISPRAVNIHSESTOMESECNIHLAKICETVOROCIKL + brNEISPRAVNIHSESTOMESECNIHLAKICETVOROCIKL + brISPRAVNIHSESTOMESECNIHLAKICETVOROCIKLPONOVO + brNEISPRAVNIHSESTOMESECNIHLAKICETVOROCIKLPONOVO;
            if (UKUPNOSESTOMESECNILAKICETVOROCIKL > 0)
                xlSheet.Range["I10"].Value = UKUPNOSESTOMESECNILAKICETVOROCIKL;
            if (brPREKINUTIHVANREDOVNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["P10"].Value = brPREKINUTIHVANREDOVNIHLAKICETVOROCIKL;
            if (brISPRAVNIHVANREDOVNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["Q10"].Value = brISPRAVNIHVANREDOVNIHLAKICETVOROCIKL;
            if (brNEISPRAVNIHVANREDOVNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["R10"].Value = brNEISPRAVNIHVANREDOVNIHLAKICETVOROCIKL;
            if (brISPRAVNIHVANREDOVNIHLAKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["S10"].Value = brISPRAVNIHVANREDOVNIHLAKICETVOROCIKLPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHLAKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["T10"].Value = brNEISPRAVNIHVANREDOVNIHLAKICETVOROCIKLPONOVO;
            int UKUPNOVANREDNILAKICETVOROCIKL = brPREKINUTIHVANREDOVNIHLAKICETVOROCIKL + brISPRAVNIHVANREDOVNIHLAKICETVOROCIKL + brNEISPRAVNIHVANREDOVNIHLAKICETVOROCIKL + brISPRAVNIHVANREDOVNIHLAKICETVOROCIKLPONOVO + brNEISPRAVNIHVANREDOVNIHLAKICETVOROCIKLPONOVO;
            if (UKUPNOVANREDNILAKICETVOROCIKL > 0)
                xlSheet.Range["O10"].Value = UKUPNOVANREDNILAKICETVOROCIKL;
            if (brISPRAVNIHKONTROLNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["V10"].Value = brISPRAVNIHKONTROLNIHLAKICETVOROCIKL;
            if (brNEISPRAVNIHKONTROLNIHLAKICETVOROCIKL > 0)
                xlSheet.Range["W10"].Value = brNEISPRAVNIHKONTROLNIHLAKICETVOROCIKL;
            int UKUPNOKONTROLNILAKICETVOROCIKL = brISPRAVNIHKONTROLNIHLAKICETVOROCIKL + brNEISPRAVNIHKONTROLNIHLAKICETVOROCIKL;
            if (UKUPNOKONTROLNILAKICETVOROCIKL > 0)
                xlSheet.Range["U10"].Value = UKUPNOKONTROLNILAKICETVOROCIKL;
            int UKUPNOLAKICETVOROCIKL = UKUPNOREDOVNILAKICETVOROCIKL + UKUPNOSESTOMESECNILAKICETVOROCIKL + UKUPNOVANREDNILAKICETVOROCIKL + UKUPNOKONTROLNILAKICETVOROCIKL;
            if (UKUPNOLAKICETVOROCIKL > 0)
                xlSheet.Range["B10"].Value = UKUPNOLAKICETVOROCIKL;
            #endregion
            #region UKUPNOTESKICETVOROCIKL
            if (brPREKINUTIHREDOVNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["D11"].Value = brPREKINUTIHREDOVNIHTESKICETVOROCIKL;
            if (brISPRAVNIHREDOVNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["E11"].Value = brISPRAVNIHREDOVNIHTESKICETVOROCIKL;
            if (brNEISPRAVNIHREDOVNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["F11"].Value = brNEISPRAVNIHREDOVNIHTESKICETVOROCIKL;
            if (brISPRAVNIHREDOVNIHTESKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["G11"].Value = brISPRAVNIHREDOVNIHTESKICETVOROCIKLPONOVO;
            if (brNEISPRAVNIHREDOVNIHTESKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["H11"].Value = brNEISPRAVNIHREDOVNIHTESKICETVOROCIKLPONOVO;
            int UKUPNOREDOVNITESKICETVOROCIKL = brPREKINUTIHREDOVNIHTESKICETVOROCIKL + brISPRAVNIHREDOVNIHTESKICETVOROCIKL + brNEISPRAVNIHREDOVNIHTESKICETVOROCIKL + brISPRAVNIHREDOVNIHTESKICETVOROCIKLPONOVO + brNEISPRAVNIHREDOVNIHTESKICETVOROCIKLPONOVO;
            if (UKUPNOREDOVNITESKICETVOROCIKL > 0)
                xlSheet.Range["C11"].Value = UKUPNOREDOVNITESKICETVOROCIKL;
            if (brPREKINUTIHSESTOMESECNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["J11"].Value = brPREKINUTIHSESTOMESECNIHTESKICETVOROCIKL;
            if (brISPRAVNIHSESTOMESECNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["K11"].Value = brISPRAVNIHSESTOMESECNIHTESKICETVOROCIKL;
            if (brNEISPRAVNIHSESTOMESECNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["L11"].Value = brNEISPRAVNIHSESTOMESECNIHTESKICETVOROCIKL;
            if (brISPRAVNIHSESTOMESECNIHTESKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["M11"].Value = brISPRAVNIHSESTOMESECNIHTESKICETVOROCIKLPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHTESKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["N11"].Value = brNEISPRAVNIHSESTOMESECNIHTESKICETVOROCIKLPONOVO;
            int UKUPNOSESTOMESECNITESKICETVOROCIKL = brPREKINUTIHSESTOMESECNIHTESKICETVOROCIKL + brISPRAVNIHSESTOMESECNIHTESKICETVOROCIKL + brNEISPRAVNIHSESTOMESECNIHTESKICETVOROCIKL + brISPRAVNIHSESTOMESECNIHTESKICETVOROCIKLPONOVO + brNEISPRAVNIHSESTOMESECNIHTESKICETVOROCIKLPONOVO;
            if (UKUPNOSESTOMESECNITESKICETVOROCIKL > 0)
                xlSheet.Range["I11"].Value = UKUPNOSESTOMESECNITESKICETVOROCIKL;
            if (brPREKINUTIHVANREDOVNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["P11"].Value = brPREKINUTIHVANREDOVNIHTESKICETVOROCIKL;
            if (brISPRAVNIHVANREDOVNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["Q11"].Value = brISPRAVNIHVANREDOVNIHTESKICETVOROCIKL;
            if (brNEISPRAVNIHVANREDOVNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["R11"].Value = brNEISPRAVNIHVANREDOVNIHTESKICETVOROCIKL;
            if (brISPRAVNIHVANREDOVNIHTESKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["S11"].Value = brISPRAVNIHVANREDOVNIHTESKICETVOROCIKLPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHTESKICETVOROCIKLPONOVO > 0)
                xlSheet.Range["T11"].Value = brNEISPRAVNIHVANREDOVNIHTESKICETVOROCIKLPONOVO;
            int UKUPNOVANREDNITESKICETVOROCIKL = brPREKINUTIHVANREDOVNIHTESKICETVOROCIKL + brISPRAVNIHVANREDOVNIHTESKICETVOROCIKL + brNEISPRAVNIHVANREDOVNIHTESKICETVOROCIKL + brISPRAVNIHVANREDOVNIHTESKICETVOROCIKLPONOVO + brNEISPRAVNIHVANREDOVNIHTESKICETVOROCIKLPONOVO;
            if (UKUPNOVANREDNITESKICETVOROCIKL > 0)
                xlSheet.Range["O11"].Value = UKUPNOVANREDNITESKICETVOROCIKL;
            if (brISPRAVNIHKONTROLNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["V11"].Value = brISPRAVNIHKONTROLNIHTESKICETVOROCIKL;
            if (brNEISPRAVNIHKONTROLNIHTESKICETVOROCIKL > 0)
                xlSheet.Range["W11"].Value = brNEISPRAVNIHKONTROLNIHTESKICETVOROCIKL;
            int UKUPNOKONTROLNITESKICETVOROCIKL = brISPRAVNIHKONTROLNIHTESKICETVOROCIKL + brNEISPRAVNIHKONTROLNIHTESKICETVOROCIKL;
            if (UKUPNOKONTROLNITESKICETVOROCIKL > 0)
                xlSheet.Range["U11"].Value = UKUPNOKONTROLNITESKICETVOROCIKL;
            int UKUPNOTESKICETVOROCIKL = UKUPNOREDOVNITESKICETVOROCIKL + UKUPNOSESTOMESECNITESKICETVOROCIKL + UKUPNOVANREDNITESKICETVOROCIKL + UKUPNOKONTROLNITESKICETVOROCIKL;
            if (UKUPNOTESKICETVOROCIKL > 0)
                xlSheet.Range["B11"].Value = UKUPNOTESKICETVOROCIKL;
            #endregion
            #region UKUPNOPUTNICKOVOZILO
            if (brPREKINUTIHREDOVNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["D12"].Value = brPREKINUTIHREDOVNIHPUTNICKOVOZILO;
            if (brISPRAVNIHREDOVNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["E12"].Value = brISPRAVNIHREDOVNIHPUTNICKOVOZILO;
            if (brNEISPRAVNIHREDOVNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["F12"].Value = brNEISPRAVNIHREDOVNIHPUTNICKOVOZILO;
            if (brISPRAVNIHREDOVNIHPUTNICKOVOZILOPONOVO > 0)
                xlSheet.Range["G12"].Value = brISPRAVNIHREDOVNIHPUTNICKOVOZILOPONOVO;
            if (brNEISPRAVNIHREDOVNIHPUTNICKOVOZILOPONOVO > 0)
                xlSheet.Range["H12"].Value = brNEISPRAVNIHREDOVNIHPUTNICKOVOZILOPONOVO;
            int UKUPNOREDOVNIPUTNICKOVOZILO = brPREKINUTIHREDOVNIHPUTNICKOVOZILO + brISPRAVNIHREDOVNIHPUTNICKOVOZILO + brNEISPRAVNIHREDOVNIHPUTNICKOVOZILO + brISPRAVNIHREDOVNIHPUTNICKOVOZILOPONOVO + brNEISPRAVNIHREDOVNIHPUTNICKOVOZILOPONOVO;
            if (UKUPNOREDOVNIPUTNICKOVOZILO > 0)
                xlSheet.Range["C12"].Value = UKUPNOREDOVNIPUTNICKOVOZILO;
            if (brPREKINUTIHSESTOMESECNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["J12"].Value = brPREKINUTIHSESTOMESECNIHPUTNICKOVOZILO;
            if (brISPRAVNIHSESTOMESECNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["K12"].Value = brISPRAVNIHSESTOMESECNIHPUTNICKOVOZILO;
            if (brNEISPRAVNIHSESTOMESECNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["L12"].Value = brNEISPRAVNIHSESTOMESECNIHPUTNICKOVOZILO;
            if (brISPRAVNIHSESTOMESECNIHPUTNICKOVOZILOPONOVO > 0)
                xlSheet.Range["M12"].Value = brISPRAVNIHSESTOMESECNIHPUTNICKOVOZILOPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHPUTNICKOVOZILOPONOVO > 0)
                xlSheet.Range["N12"].Value = brNEISPRAVNIHSESTOMESECNIHPUTNICKOVOZILOPONOVO;
            int UKUPNOSESTOMESECNIPUTNICKOVOZILO = brPREKINUTIHSESTOMESECNIHPUTNICKOVOZILO + brISPRAVNIHSESTOMESECNIHPUTNICKOVOZILO + brNEISPRAVNIHSESTOMESECNIHPUTNICKOVOZILO + brISPRAVNIHSESTOMESECNIHPUTNICKOVOZILOPONOVO + brNEISPRAVNIHSESTOMESECNIHPUTNICKOVOZILOPONOVO;
            if (UKUPNOSESTOMESECNIPUTNICKOVOZILO > 0)
                xlSheet.Range["I12"].Value = UKUPNOSESTOMESECNIPUTNICKOVOZILO;
            if (brPREKINUTIHVANREDOVNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["P12"].Value = brPREKINUTIHVANREDOVNIHPUTNICKOVOZILO;
            if (brISPRAVNIHVANREDOVNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["Q12"].Value = brISPRAVNIHVANREDOVNIHPUTNICKOVOZILO;
            if (brNEISPRAVNIHVANREDOVNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["R12"].Value = brNEISPRAVNIHVANREDOVNIHPUTNICKOVOZILO;
            if (brISPRAVNIHVANREDOVNIHPUTNICKOVOZILOPONOVO > 0)
                xlSheet.Range["S12"].Value = brISPRAVNIHVANREDOVNIHPUTNICKOVOZILOPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHPUTNICKOVOZILOPONOVO > 0)
                xlSheet.Range["T12"].Value = brNEISPRAVNIHVANREDOVNIHPUTNICKOVOZILOPONOVO;
            int UKUPNOVANREDNIPUTNICKOVOZILO = brPREKINUTIHVANREDOVNIHPUTNICKOVOZILO + brISPRAVNIHVANREDOVNIHPUTNICKOVOZILO + brNEISPRAVNIHVANREDOVNIHPUTNICKOVOZILO + brISPRAVNIHVANREDOVNIHPUTNICKOVOZILOPONOVO + brNEISPRAVNIHVANREDOVNIHPUTNICKOVOZILOPONOVO;
            if (UKUPNOVANREDNIPUTNICKOVOZILO > 0)
                xlSheet.Range["O12"].Value = UKUPNOVANREDNIPUTNICKOVOZILO;
            if (brISPRAVNIHKONTROLNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["V12"].Value = brISPRAVNIHKONTROLNIHPUTNICKOVOZILO;
            if (brNEISPRAVNIHKONTROLNIHPUTNICKOVOZILO > 0)
                xlSheet.Range["W12"].Value = brNEISPRAVNIHKONTROLNIHPUTNICKOVOZILO;
            int UKUPNOKONTROLNIPUTNICKOVOZILO = brISPRAVNIHKONTROLNIHPUTNICKOVOZILO + brNEISPRAVNIHKONTROLNIHPUTNICKOVOZILO;
            if (UKUPNOKONTROLNIPUTNICKOVOZILO > 0)
                xlSheet.Range["U12"].Value = UKUPNOKONTROLNIPUTNICKOVOZILO;
            int UKUPNOPUTNICKOVOZILO = UKUPNOREDOVNIPUTNICKOVOZILO + UKUPNOSESTOMESECNIPUTNICKOVOZILO + UKUPNOVANREDNIPUTNICKOVOZILO + UKUPNOKONTROLNIPUTNICKOVOZILO;
            if (UKUPNOPUTNICKOVOZILO > 0)
                xlSheet.Range["B12"].Value = UKUPNOPUTNICKOVOZILO;
            #endregion
            #region UKUPNOAUTOBUS
            if (brPREKINUTIHREDOVNIHAUTOBUS > 0)
                xlSheet.Range["D13"].Value = brPREKINUTIHREDOVNIHAUTOBUS;
            if (brISPRAVNIHREDOVNIHAUTOBUS > 0)
                xlSheet.Range["E13"].Value = brISPRAVNIHREDOVNIHAUTOBUS;
            if (brNEISPRAVNIHREDOVNIHAUTOBUS > 0)
                xlSheet.Range["F13"].Value = brNEISPRAVNIHREDOVNIHAUTOBUS;
            if (brISPRAVNIHREDOVNIHAUTOBUSPONOVO > 0)
                xlSheet.Range["G13"].Value = brISPRAVNIHREDOVNIHAUTOBUSPONOVO;
            if (brNEISPRAVNIHREDOVNIHAUTOBUSPONOVO > 0)
                xlSheet.Range["H13"].Value = brNEISPRAVNIHREDOVNIHAUTOBUSPONOVO;
            int UKUPNOREDOVNIAUTOBUS = brPREKINUTIHREDOVNIHAUTOBUS + brISPRAVNIHREDOVNIHAUTOBUS + brNEISPRAVNIHREDOVNIHAUTOBUS + brISPRAVNIHREDOVNIHAUTOBUSPONOVO + brNEISPRAVNIHREDOVNIHAUTOBUSPONOVO;
            if (UKUPNOREDOVNIAUTOBUS > 0)
                xlSheet.Range["C13"].Value = UKUPNOREDOVNIAUTOBUS;
            if (brPREKINUTIHSESTOMESECNIHAUTOBUS > 0)
                xlSheet.Range["J13"].Value = brPREKINUTIHSESTOMESECNIHAUTOBUS;
            if (brISPRAVNIHSESTOMESECNIHAUTOBUS > 0)
                xlSheet.Range["K13"].Value = brISPRAVNIHSESTOMESECNIHAUTOBUS;
            if (brNEISPRAVNIHSESTOMESECNIHAUTOBUS > 0)
                xlSheet.Range["L13"].Value = brNEISPRAVNIHSESTOMESECNIHAUTOBUS;
            if (brISPRAVNIHSESTOMESECNIHAUTOBUSPONOVO > 0)
                xlSheet.Range["M13"].Value = brISPRAVNIHSESTOMESECNIHAUTOBUSPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHAUTOBUSPONOVO > 0)
                xlSheet.Range["N13"].Value = brNEISPRAVNIHSESTOMESECNIHAUTOBUSPONOVO;
            int UKUPNOSESTOMESECNIAUTOBUS = brPREKINUTIHSESTOMESECNIHAUTOBUS + brISPRAVNIHSESTOMESECNIHAUTOBUS + brNEISPRAVNIHSESTOMESECNIHAUTOBUS + brISPRAVNIHSESTOMESECNIHAUTOBUSPONOVO + brNEISPRAVNIHSESTOMESECNIHAUTOBUSPONOVO;
            if (UKUPNOSESTOMESECNIAUTOBUS > 0)
                xlSheet.Range["I13"].Value = UKUPNOSESTOMESECNIAUTOBUS;
            if (brPREKINUTIHVANREDOVNIHAUTOBUS > 0)
                xlSheet.Range["P13"].Value = brPREKINUTIHVANREDOVNIHAUTOBUS;
            if (brISPRAVNIHVANREDOVNIHAUTOBUS > 0)
                xlSheet.Range["Q13"].Value = brISPRAVNIHVANREDOVNIHAUTOBUS;
            if (brNEISPRAVNIHVANREDOVNIHAUTOBUS > 0)
                xlSheet.Range["R13"].Value = brNEISPRAVNIHVANREDOVNIHAUTOBUS;
            if (brISPRAVNIHVANREDOVNIHAUTOBUSPONOVO > 0)
                xlSheet.Range["S13"].Value = brISPRAVNIHVANREDOVNIHAUTOBUSPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHAUTOBUSPONOVO > 0)
                xlSheet.Range["T13"].Value = brNEISPRAVNIHVANREDOVNIHAUTOBUSPONOVO;
            int UKUPNOVANREDNIAUTOBUS = brPREKINUTIHVANREDOVNIHAUTOBUS + brISPRAVNIHVANREDOVNIHAUTOBUS + brNEISPRAVNIHVANREDOVNIHAUTOBUS + brISPRAVNIHVANREDOVNIHAUTOBUSPONOVO + brNEISPRAVNIHVANREDOVNIHAUTOBUSPONOVO;
            if (UKUPNOVANREDNIAUTOBUS > 0)
                xlSheet.Range["O13"].Value = UKUPNOVANREDNIAUTOBUS;
            if (brISPRAVNIHKONTROLNIHAUTOBUS > 0)
                xlSheet.Range["V13"].Value = brISPRAVNIHKONTROLNIHAUTOBUS;
            if (brNEISPRAVNIHKONTROLNIHAUTOBUS > 0)
                xlSheet.Range["W13"].Value = brNEISPRAVNIHKONTROLNIHAUTOBUS;
            int UKUPNOKONTROLNIAUTOBUS = brISPRAVNIHKONTROLNIHAUTOBUS + brNEISPRAVNIHKONTROLNIHAUTOBUS;
            if (UKUPNOKONTROLNIAUTOBUS > 0)
                xlSheet.Range["U13"].Value = UKUPNOKONTROLNIAUTOBUS;
            int UKUPNOAUTOBUS = UKUPNOREDOVNIAUTOBUS + UKUPNOSESTOMESECNIAUTOBUS + UKUPNOVANREDNIAUTOBUS + UKUPNOKONTROLNIAUTOBUS;
            if (UKUPNOAUTOBUS > 0)
                xlSheet.Range["B13"].Value = UKUPNOAUTOBUS;
            #endregion
            #region UKUPNOTERETNOVOZILO
            if (brPREKINUTIHREDOVNIHTERETNOVOZILO > 0)
                xlSheet.Range["D14"].Value = brPREKINUTIHREDOVNIHTERETNOVOZILO;
            if (brISPRAVNIHREDOVNIHTERETNOVOZILO > 0)
                xlSheet.Range["E14"].Value = brISPRAVNIHREDOVNIHTERETNOVOZILO;
            if (brNEISPRAVNIHREDOVNIHTERETNOVOZILO > 0)
                xlSheet.Range["F14"].Value = brNEISPRAVNIHREDOVNIHTERETNOVOZILO;
            if (brISPRAVNIHREDOVNIHTERETNOVOZILOPONOVO > 0)
                xlSheet.Range["G14"].Value = brISPRAVNIHREDOVNIHTERETNOVOZILOPONOVO;
            if (brNEISPRAVNIHREDOVNIHTERETNOVOZILOPONOVO > 0)
                xlSheet.Range["H14"].Value = brNEISPRAVNIHREDOVNIHTERETNOVOZILOPONOVO;
            int UKUPNOREDOVNITERETNOVOZILO = brPREKINUTIHREDOVNIHTERETNOVOZILO + brISPRAVNIHREDOVNIHTERETNOVOZILO + brNEISPRAVNIHREDOVNIHTERETNOVOZILO + brISPRAVNIHREDOVNIHTERETNOVOZILOPONOVO + brNEISPRAVNIHREDOVNIHTERETNOVOZILOPONOVO;
            if (UKUPNOREDOVNITERETNOVOZILO > 0)
                xlSheet.Range["C14"].Value = UKUPNOREDOVNITERETNOVOZILO;
            if (brPREKINUTIHSESTOMESECNIHTERETNOVOZILO > 0)
                xlSheet.Range["J14"].Value = brPREKINUTIHSESTOMESECNIHTERETNOVOZILO;
            if (brISPRAVNIHSESTOMESECNIHTERETNOVOZILO > 0)
                xlSheet.Range["K14"].Value = brISPRAVNIHSESTOMESECNIHTERETNOVOZILO;
            if (brNEISPRAVNIHSESTOMESECNIHTERETNOVOZILO > 0)
                xlSheet.Range["L14"].Value = brNEISPRAVNIHSESTOMESECNIHTERETNOVOZILO;
            if (brISPRAVNIHSESTOMESECNIHTERETNOVOZILOPONOVO > 0)
                xlSheet.Range["M14"].Value = brISPRAVNIHSESTOMESECNIHTERETNOVOZILOPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHTERETNOVOZILOPONOVO > 0)
                xlSheet.Range["N14"].Value = brNEISPRAVNIHSESTOMESECNIHTERETNOVOZILOPONOVO;
            int UKUPNOSESTOMESECNITERETNOVOZILO = brPREKINUTIHSESTOMESECNIHTERETNOVOZILO + brISPRAVNIHSESTOMESECNIHTERETNOVOZILO + brNEISPRAVNIHSESTOMESECNIHTERETNOVOZILO + brISPRAVNIHSESTOMESECNIHTERETNOVOZILOPONOVO + brNEISPRAVNIHSESTOMESECNIHTERETNOVOZILOPONOVO;
            if (UKUPNOSESTOMESECNITERETNOVOZILO > 0)
                xlSheet.Range["I14"].Value = UKUPNOSESTOMESECNITERETNOVOZILO;
            if (brPREKINUTIHVANREDOVNIHTERETNOVOZILO > 0)
                xlSheet.Range["P14"].Value = brPREKINUTIHVANREDOVNIHTERETNOVOZILO;
            if (brISPRAVNIHVANREDOVNIHTERETNOVOZILO > 0)
                xlSheet.Range["Q14"].Value = brISPRAVNIHVANREDOVNIHTERETNOVOZILO;
            if (brNEISPRAVNIHVANREDOVNIHTERETNOVOZILO > 0)
                xlSheet.Range["R14"].Value = brNEISPRAVNIHVANREDOVNIHTERETNOVOZILO;
            if (brISPRAVNIHVANREDOVNIHTERETNOVOZILOPONOVO > 0)
                xlSheet.Range["S14"].Value = brISPRAVNIHVANREDOVNIHTERETNOVOZILOPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHTERETNOVOZILOPONOVO > 0)
                xlSheet.Range["T14"].Value = brNEISPRAVNIHVANREDOVNIHTERETNOVOZILOPONOVO;
            int UKUPNOVANREDNITERETNOVOZILO = brPREKINUTIHVANREDOVNIHTERETNOVOZILO + brISPRAVNIHVANREDOVNIHTERETNOVOZILO + brNEISPRAVNIHVANREDOVNIHTERETNOVOZILO + brISPRAVNIHVANREDOVNIHTERETNOVOZILOPONOVO + brNEISPRAVNIHVANREDOVNIHTERETNOVOZILOPONOVO;
            if (UKUPNOVANREDNITERETNOVOZILO > 0)
                xlSheet.Range["O14"].Value = UKUPNOVANREDNITERETNOVOZILO;
            if (brISPRAVNIHKONTROLNIHTERETNOVOZILO > 0)
                xlSheet.Range["V14"].Value = brISPRAVNIHKONTROLNIHTERETNOVOZILO;
            if (brNEISPRAVNIHKONTROLNIHTERETNOVOZILO > 0)
                xlSheet.Range["W14"].Value = brNEISPRAVNIHKONTROLNIHTERETNOVOZILO;
            int UKUPNOKONTROLNITERETNOVOZILO = brISPRAVNIHKONTROLNIHTERETNOVOZILO + brNEISPRAVNIHKONTROLNIHTERETNOVOZILO;
            if (UKUPNOKONTROLNITERETNOVOZILO > 0)
                xlSheet.Range["U14"].Value = UKUPNOKONTROLNITERETNOVOZILO;
            int UKUPNOTERETNOVOZILO = UKUPNOREDOVNITERETNOVOZILO + UKUPNOSESTOMESECNITERETNOVOZILO + UKUPNOVANREDNITERETNOVOZILO + UKUPNOKONTROLNITERETNOVOZILO;
            if (UKUPNOTERETNOVOZILO > 0)
                xlSheet.Range["B14"].Value = UKUPNOTERETNOVOZILO;
            #endregion
            #region UKUPNORADNAMASINA
            if (brPREKINUTIHREDOVNIHRADNAMASINA > 0)
                xlSheet.Range["D15"].Value = brPREKINUTIHREDOVNIHRADNAMASINA;
            if (brISPRAVNIHREDOVNIHRADNAMASINA > 0)
                xlSheet.Range["E15"].Value = brISPRAVNIHREDOVNIHRADNAMASINA;
            if (brNEISPRAVNIHREDOVNIHRADNAMASINA > 0)
                xlSheet.Range["F15"].Value = brNEISPRAVNIHREDOVNIHRADNAMASINA;
            if (brISPRAVNIHREDOVNIHRADNAMASINAPONOVO > 0)
                xlSheet.Range["G15"].Value = brISPRAVNIHREDOVNIHRADNAMASINAPONOVO;
            if (brNEISPRAVNIHREDOVNIHRADNAMASINAPONOVO > 0)
                xlSheet.Range["H15"].Value = brNEISPRAVNIHREDOVNIHRADNAMASINAPONOVO;
            int UKUPNOREDOVNIRADNAMASINA = brPREKINUTIHREDOVNIHRADNAMASINA + brISPRAVNIHREDOVNIHRADNAMASINA + brNEISPRAVNIHREDOVNIHRADNAMASINA + brISPRAVNIHREDOVNIHRADNAMASINAPONOVO + brNEISPRAVNIHREDOVNIHRADNAMASINAPONOVO;
            if (UKUPNOREDOVNIRADNAMASINA > 0)
                xlSheet.Range["C15"].Value = UKUPNOREDOVNIRADNAMASINA;
            if (brPREKINUTIHSESTOMESECNIHRADNAMASINA > 0)
                xlSheet.Range["J15"].Value = brPREKINUTIHSESTOMESECNIHRADNAMASINA;
            if (brISPRAVNIHSESTOMESECNIHRADNAMASINA > 0)
                xlSheet.Range["K15"].Value = brISPRAVNIHSESTOMESECNIHRADNAMASINA;
            if (brNEISPRAVNIHSESTOMESECNIHRADNAMASINA > 0)
                xlSheet.Range["L15"].Value = brNEISPRAVNIHSESTOMESECNIHRADNAMASINA;
            if (brISPRAVNIHSESTOMESECNIHRADNAMASINAPONOVO > 0)
                xlSheet.Range["M15"].Value = brISPRAVNIHSESTOMESECNIHRADNAMASINAPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHRADNAMASINAPONOVO > 0)
                xlSheet.Range["N15"].Value = brNEISPRAVNIHSESTOMESECNIHRADNAMASINAPONOVO;
            int UKUPNOSESTOMESECNIRADNAMASINA = brPREKINUTIHSESTOMESECNIHRADNAMASINA + brISPRAVNIHSESTOMESECNIHRADNAMASINA + brNEISPRAVNIHSESTOMESECNIHRADNAMASINA + brISPRAVNIHSESTOMESECNIHRADNAMASINAPONOVO + brNEISPRAVNIHSESTOMESECNIHRADNAMASINAPONOVO;
            if (UKUPNOSESTOMESECNIRADNAMASINA > 0)
                xlSheet.Range["I15"].Value = UKUPNOSESTOMESECNIRADNAMASINA;
            if (brPREKINUTIHVANREDOVNIHRADNAMASINA > 0)
                xlSheet.Range["P15"].Value = brPREKINUTIHVANREDOVNIHRADNAMASINA;
            if (brISPRAVNIHVANREDOVNIHRADNAMASINA > 0)
                xlSheet.Range["Q15"].Value = brISPRAVNIHVANREDOVNIHRADNAMASINA;
            if (brNEISPRAVNIHVANREDOVNIHRADNAMASINA > 0)
                xlSheet.Range["R15"].Value = brNEISPRAVNIHVANREDOVNIHRADNAMASINA;
            if (brISPRAVNIHVANREDOVNIHRADNAMASINAPONOVO > 0)
                xlSheet.Range["S15"].Value = brISPRAVNIHVANREDOVNIHRADNAMASINAPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHRADNAMASINAPONOVO > 0)
                xlSheet.Range["T15"].Value = brNEISPRAVNIHVANREDOVNIHRADNAMASINAPONOVO;
            int UKUPNOVANREDNIRADNAMASINA = brPREKINUTIHVANREDOVNIHRADNAMASINA + brISPRAVNIHVANREDOVNIHRADNAMASINA + brNEISPRAVNIHVANREDOVNIHRADNAMASINA + brISPRAVNIHVANREDOVNIHRADNAMASINAPONOVO + brNEISPRAVNIHVANREDOVNIHRADNAMASINAPONOVO;
            if (UKUPNOVANREDNIRADNAMASINA > 0)
                xlSheet.Range["O15"].Value = UKUPNOVANREDNIRADNAMASINA;
            if (brISPRAVNIHKONTROLNIHRADNAMASINA > 0)
                xlSheet.Range["V15"].Value = brISPRAVNIHKONTROLNIHRADNAMASINA;
            if (brNEISPRAVNIHKONTROLNIHRADNAMASINA > 0)
                xlSheet.Range["W15"].Value = brNEISPRAVNIHKONTROLNIHRADNAMASINA;
            int UKUPNOKONTROLNIRADNAMASINA = brISPRAVNIHKONTROLNIHRADNAMASINA + brNEISPRAVNIHKONTROLNIHRADNAMASINA;
            if (UKUPNOKONTROLNIRADNAMASINA > 0)
                xlSheet.Range["U15"].Value = UKUPNOKONTROLNIRADNAMASINA;
            int UKUPNORADNAMASINA = UKUPNOREDOVNIRADNAMASINA + UKUPNOSESTOMESECNIRADNAMASINA + UKUPNOVANREDNIRADNAMASINA + UKUPNOKONTROLNIRADNAMASINA;
            if (UKUPNORADNAMASINA > 0)
                xlSheet.Range["B15"].Value = UKUPNORADNAMASINA;
            #endregion
            #region UKUPNOTRAKTOR
            if (brPREKINUTIHREDOVNIHTRAKTOR > 0)
                xlSheet.Range["D16"].Value = brPREKINUTIHREDOVNIHTRAKTOR;
            if (brISPRAVNIHREDOVNIHTRAKTOR > 0)
                xlSheet.Range["E16"].Value = brISPRAVNIHREDOVNIHTRAKTOR;
            if (brNEISPRAVNIHREDOVNIHTRAKTOR > 0)
                xlSheet.Range["F16"].Value = brNEISPRAVNIHREDOVNIHTRAKTOR;
            if (brISPRAVNIHREDOVNIHTRAKTORPONOVO > 0)
                xlSheet.Range["G16"].Value = brISPRAVNIHREDOVNIHTRAKTORPONOVO;
            if (brNEISPRAVNIHREDOVNIHTRAKTORPONOVO > 0)
                xlSheet.Range["H16"].Value = brNEISPRAVNIHREDOVNIHTRAKTORPONOVO;
            int UKUPNOREDOVNITRAKTOR = brPREKINUTIHREDOVNIHTRAKTOR + brISPRAVNIHREDOVNIHTRAKTOR + brNEISPRAVNIHREDOVNIHTRAKTOR + brISPRAVNIHREDOVNIHTRAKTORPONOVO + brNEISPRAVNIHREDOVNIHTRAKTORPONOVO;
            if (UKUPNOREDOVNITRAKTOR > 0)
                xlSheet.Range["C16"].Value = UKUPNOREDOVNITRAKTOR;
            if (brPREKINUTIHSESTOMESECNIHTRAKTOR > 0)
                xlSheet.Range["J16"].Value = brPREKINUTIHSESTOMESECNIHTRAKTOR;
            if (brISPRAVNIHSESTOMESECNIHTRAKTOR > 0)
                xlSheet.Range["K16"].Value = brISPRAVNIHSESTOMESECNIHTRAKTOR;
            if (brNEISPRAVNIHSESTOMESECNIHTRAKTOR > 0)
                xlSheet.Range["L16"].Value = brNEISPRAVNIHSESTOMESECNIHTRAKTOR;
            if (brISPRAVNIHSESTOMESECNIHTRAKTORPONOVO > 0)
                xlSheet.Range["M16"].Value = brISPRAVNIHSESTOMESECNIHTRAKTORPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHTRAKTORPONOVO > 0)
                xlSheet.Range["N16"].Value = brNEISPRAVNIHSESTOMESECNIHTRAKTORPONOVO;
            int UKUPNOSESTOMESECNITRAKTOR = brPREKINUTIHSESTOMESECNIHTRAKTOR + brISPRAVNIHSESTOMESECNIHTRAKTOR + brNEISPRAVNIHSESTOMESECNIHTRAKTOR + brISPRAVNIHSESTOMESECNIHTRAKTORPONOVO + brNEISPRAVNIHSESTOMESECNIHTRAKTORPONOVO;
            if (UKUPNOSESTOMESECNITRAKTOR > 0)
                xlSheet.Range["I16"].Value = UKUPNOSESTOMESECNITRAKTOR;
            if (brPREKINUTIHVANREDOVNIHTRAKTOR > 0)
                xlSheet.Range["P16"].Value = brPREKINUTIHVANREDOVNIHTRAKTOR;
            if (brISPRAVNIHVANREDOVNIHTRAKTOR > 0)
                xlSheet.Range["Q16"].Value = brISPRAVNIHVANREDOVNIHTRAKTOR;
            if (brNEISPRAVNIHVANREDOVNIHTRAKTOR > 0)
                xlSheet.Range["R16"].Value = brNEISPRAVNIHVANREDOVNIHTRAKTOR;
            if (brISPRAVNIHVANREDOVNIHTRAKTORPONOVO > 0)
                xlSheet.Range["S16"].Value = brISPRAVNIHVANREDOVNIHTRAKTORPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHTRAKTORPONOVO > 0)
                xlSheet.Range["T16"].Value = brNEISPRAVNIHVANREDOVNIHTRAKTORPONOVO;
            int UKUPNOVANREDNITRAKTOR = brPREKINUTIHVANREDOVNIHTRAKTOR + brISPRAVNIHVANREDOVNIHTRAKTOR + brNEISPRAVNIHVANREDOVNIHTRAKTOR + brISPRAVNIHVANREDOVNIHTRAKTORPONOVO + brNEISPRAVNIHVANREDOVNIHTRAKTORPONOVO;
            if (UKUPNOVANREDNITRAKTOR > 0)
                xlSheet.Range["O16"].Value = UKUPNOVANREDNITRAKTOR;
            if (brISPRAVNIHKONTROLNIHTRAKTOR > 0)
                xlSheet.Range["V16"].Value = brISPRAVNIHKONTROLNIHTRAKTOR;
            if (brNEISPRAVNIHKONTROLNIHTRAKTOR > 0)
                xlSheet.Range["W16"].Value = brNEISPRAVNIHKONTROLNIHTRAKTOR;
            int UKUPNOKONTROLNITRAKTOR = brISPRAVNIHKONTROLNIHTRAKTOR + brNEISPRAVNIHKONTROLNIHTRAKTOR;
            if (UKUPNOKONTROLNITRAKTOR > 0)
                xlSheet.Range["U16"].Value = UKUPNOKONTROLNITRAKTOR;
            int UKUPNOTRAKTOR = UKUPNOREDOVNITRAKTOR + UKUPNOSESTOMESECNITRAKTOR + UKUPNOVANREDNITRAKTOR + UKUPNOKONTROLNITRAKTOR;
            if (UKUPNOTRAKTOR > 0)
                xlSheet.Range["B16"].Value = UKUPNOTRAKTOR;
            #endregion
            #region UKUPNOMOTOKULTIVATOR
            if (brPREKINUTIHREDOVNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["D17"].Value = brPREKINUTIHREDOVNIHMOTOKULTIVATOR;
            if (brISPRAVNIHREDOVNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["E17"].Value = brISPRAVNIHREDOVNIHMOTOKULTIVATOR;
            if (brNEISPRAVNIHREDOVNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["F17"].Value = brNEISPRAVNIHREDOVNIHMOTOKULTIVATOR;
            if (brISPRAVNIHREDOVNIHMOTOKULTIVATORPONOVO > 0)
                xlSheet.Range["G17"].Value = brISPRAVNIHREDOVNIHMOTOKULTIVATORPONOVO;
            if (brNEISPRAVNIHREDOVNIHMOTOKULTIVATORPONOVO > 0)
                xlSheet.Range["H17"].Value = brNEISPRAVNIHREDOVNIHMOTOKULTIVATORPONOVO;
            int UKUPNOREDOVNIMOTOKULTIVATOR = brPREKINUTIHREDOVNIHMOTOKULTIVATOR + brISPRAVNIHREDOVNIHMOTOKULTIVATOR + brNEISPRAVNIHREDOVNIHMOTOKULTIVATOR + brISPRAVNIHREDOVNIHMOTOKULTIVATORPONOVO + brNEISPRAVNIHREDOVNIHMOTOKULTIVATORPONOVO;
            if (UKUPNOREDOVNIMOTOKULTIVATOR > 0)
                xlSheet.Range["C17"].Value = UKUPNOREDOVNIMOTOKULTIVATOR;
            if (brPREKINUTIHSESTOMESECNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["J17"].Value = brPREKINUTIHSESTOMESECNIHMOTOKULTIVATOR;
            if (brISPRAVNIHSESTOMESECNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["K17"].Value = brISPRAVNIHSESTOMESECNIHMOTOKULTIVATOR;
            if (brNEISPRAVNIHSESTOMESECNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["L17"].Value = brNEISPRAVNIHSESTOMESECNIHMOTOKULTIVATOR;
            if (brISPRAVNIHSESTOMESECNIHMOTOKULTIVATORPONOVO > 0)
                xlSheet.Range["M17"].Value = brISPRAVNIHSESTOMESECNIHMOTOKULTIVATORPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHMOTOKULTIVATORPONOVO > 0)
                xlSheet.Range["N17"].Value = brNEISPRAVNIHSESTOMESECNIHMOTOKULTIVATORPONOVO;
            int UKUPNOSESTOMESECNIMOTOKULTIVATOR = brPREKINUTIHSESTOMESECNIHMOTOKULTIVATOR + brISPRAVNIHSESTOMESECNIHMOTOKULTIVATOR + brNEISPRAVNIHSESTOMESECNIHMOTOKULTIVATOR + brISPRAVNIHSESTOMESECNIHMOTOKULTIVATORPONOVO + brNEISPRAVNIHSESTOMESECNIHMOTOKULTIVATORPONOVO;
            if (UKUPNOSESTOMESECNIMOTOKULTIVATOR > 0)
                xlSheet.Range["I17"].Value = UKUPNOSESTOMESECNIMOTOKULTIVATOR;
            if (brPREKINUTIHVANREDOVNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["P17"].Value = brPREKINUTIHVANREDOVNIHMOTOKULTIVATOR;
            if (brISPRAVNIHVANREDOVNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["Q17"].Value = brISPRAVNIHVANREDOVNIHMOTOKULTIVATOR;
            if (brNEISPRAVNIHVANREDOVNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["R17"].Value = brNEISPRAVNIHVANREDOVNIHMOTOKULTIVATOR;
            if (brISPRAVNIHVANREDOVNIHMOTOKULTIVATORPONOVO > 0)
                xlSheet.Range["S17"].Value = brISPRAVNIHVANREDOVNIHMOTOKULTIVATORPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHMOTOKULTIVATORPONOVO > 0)
                xlSheet.Range["T17"].Value = brNEISPRAVNIHVANREDOVNIHMOTOKULTIVATORPONOVO;
            int UKUPNOVANREDNIMOTOKULTIVATOR = brPREKINUTIHVANREDOVNIHMOTOKULTIVATOR + brISPRAVNIHVANREDOVNIHMOTOKULTIVATOR + brNEISPRAVNIHVANREDOVNIHMOTOKULTIVATOR + brISPRAVNIHVANREDOVNIHMOTOKULTIVATORPONOVO + brNEISPRAVNIHVANREDOVNIHMOTOKULTIVATORPONOVO;
            if (UKUPNOVANREDNIMOTOKULTIVATOR > 0)
                xlSheet.Range["O17"].Value = UKUPNOVANREDNIMOTOKULTIVATOR;
            if (brISPRAVNIHKONTROLNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["V17"].Value = brISPRAVNIHKONTROLNIHMOTOKULTIVATOR;
            if (brNEISPRAVNIHKONTROLNIHMOTOKULTIVATOR > 0)
                xlSheet.Range["W17"].Value = brNEISPRAVNIHKONTROLNIHMOTOKULTIVATOR;
            int UKUPNOKONTROLNIMOTOKULTIVATOR = brISPRAVNIHKONTROLNIHMOTOKULTIVATOR + brNEISPRAVNIHKONTROLNIHMOTOKULTIVATOR;
            if (UKUPNOKONTROLNIMOTOKULTIVATOR > 0)
                xlSheet.Range["U17"].Value = UKUPNOKONTROLNIMOTOKULTIVATOR;
            int UKUPNOMOTOKULTIVATOR = UKUPNOREDOVNIMOTOKULTIVATOR + UKUPNOSESTOMESECNIMOTOKULTIVATOR + UKUPNOVANREDNIMOTOKULTIVATOR + UKUPNOKONTROLNIMOTOKULTIVATOR;
            if (UKUPNOMOTOKULTIVATOR > 0)
                xlSheet.Range["B17"].Value = UKUPNOMOTOKULTIVATOR;
            #endregion
            #region UKUPNOPRIKLJUCNOVOZILO
            if (brPREKINUTIHREDOVNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["D18"].Value = brPREKINUTIHREDOVNIHPRIKLJUCNOVOZILO;
            if (brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["E18"].Value = brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILO;
            if (brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["F18"].Value = brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILO;
            if (brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOPONOVO > 0)
                xlSheet.Range["G18"].Value = brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOPONOVO;
            if (brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOPONOVO > 0)
                xlSheet.Range["H18"].Value = brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOPONOVO;
            int UKUPNOREDOVNIPRIKLJUCNOVOZILO = brPREKINUTIHREDOVNIHPRIKLJUCNOVOZILO + brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILO + brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILO + brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOPONOVO + brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOPONOVO;
            if (UKUPNOREDOVNIPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["C18"].Value = UKUPNOREDOVNIPRIKLJUCNOVOZILO;
            if (brPREKINUTIHSESTOMESECNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["J18"].Value = brPREKINUTIHSESTOMESECNIHPRIKLJUCNOVOZILO;
            if (brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["K18"].Value = brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILO;
            if (brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["L18"].Value = brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILO;
            if (brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOPONOVO > 0)
                xlSheet.Range["M18"].Value = brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOPONOVO > 0)
                xlSheet.Range["N18"].Value = brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOPONOVO;
            int UKUPNOSESTOMESECNIPRIKLJUCNOVOZILO = brPREKINUTIHSESTOMESECNIHPRIKLJUCNOVOZILO + brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILO + brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILO + brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOPONOVO + brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOPONOVO;
            if (UKUPNOSESTOMESECNIPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["I18"].Value = UKUPNOSESTOMESECNIPRIKLJUCNOVOZILO;
            if (brPREKINUTIHVANREDOVNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["P18"].Value = brPREKINUTIHVANREDOVNIHPRIKLJUCNOVOZILO;
            if (brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["Q18"].Value = brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILO;
            if (brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["R18"].Value = brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILO;
            if (brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOPONOVO > 0)
                xlSheet.Range["S18"].Value = brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOPONOVO > 0)
                xlSheet.Range["T18"].Value = brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOPONOVO;
            int UKUPNOVANREDNIPRIKLJUCNOVOZILO = brPREKINUTIHVANREDOVNIHPRIKLJUCNOVOZILO + brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILO + brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILO + brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOPONOVO + brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOPONOVO;
            if (UKUPNOVANREDNIPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["O18"].Value = UKUPNOVANREDNIPRIKLJUCNOVOZILO;
            if (brISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["V18"].Value = brISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILO;
            if (brNEISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["W18"].Value = brNEISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILO;
            int UKUPNOKONTROLNIPRIKLJUCNOVOZILO = brISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILO + brNEISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILO;
            if (UKUPNOKONTROLNIPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["U18"].Value = UKUPNOKONTROLNIPRIKLJUCNOVOZILO;
            int UKUPNOPRIKLJUCNOVOZILO = UKUPNOREDOVNIPRIKLJUCNOVOZILO + UKUPNOSESTOMESECNIPRIKLJUCNOVOZILO + UKUPNOVANREDNIPRIKLJUCNOVOZILO + UKUPNOKONTROLNIPRIKLJUCNOVOZILO;
            if (UKUPNOPRIKLJUCNOVOZILO > 0)
                xlSheet.Range["B18"].Value = UKUPNOPRIKLJUCNOVOZILO;
            #endregion
            #region UKUPNOPRIKLJUCNOVOZILOZATRAKTOR
            if (brPREKINUTIHREDOVNIHPRIKLJUCNOZATRAKTOR > 0)
                xlSheet.Range["D19"].Value = brPREKINUTIHREDOVNIHPRIKLJUCNOZATRAKTOR;
            if (brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["E19"].Value = brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["F19"].Value = brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO > 0)
                xlSheet.Range["G19"].Value = brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            if (brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO > 0)
                xlSheet.Range["H19"].Value = brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            int UKUPNOREDOVNIPRIKLJUCNOVOZILOZATRAKTOR = brPREKINUTIHREDOVNIHPRIKLJUCNOZATRAKTOR + brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR + brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR + brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO + brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            if (UKUPNOREDOVNIPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["C19"].Value = UKUPNOREDOVNIPRIKLJUCNOVOZILOZATRAKTOR;
            if (brPREKINUTIHSESTOMESECNIHPRIKLJUCNOZATRAKTOR > 0)
                xlSheet.Range["J19"].Value = brPREKINUTIHSESTOMESECNIHPRIKLJUCNOZATRAKTOR;
            if (brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["K19"].Value = brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["L19"].Value = brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO > 0)
                xlSheet.Range["M19"].Value = brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            if (brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO > 0)
                xlSheet.Range["N19"].Value = brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            int UKUPNOSESTOMESECNIPRIKLJUCNOVOZILOZATRAKTOR = brPREKINUTIHSESTOMESECNIHPRIKLJUCNOZATRAKTOR + brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTOR + brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTOR + brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO + brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            if (UKUPNOSESTOMESECNIPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["I19"].Value = UKUPNOSESTOMESECNIPRIKLJUCNOVOZILOZATRAKTOR;
            if (brPREKINUTIHVANREDOVNIHPRIKLJUCNOZATRAKTOR > 0)
                xlSheet.Range["P19"].Value = brPREKINUTIHVANREDOVNIHPRIKLJUCNOZATRAKTOR;
            if (brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["Q19"].Value = brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["R19"].Value = brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO > 0)
                xlSheet.Range["S19"].Value = brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            if (brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO > 0)
                xlSheet.Range["T19"].Value = brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            int UKUPNOVANREDNIPRIKLJUCNOVOZILOZATRAKTOR = brPREKINUTIHVANREDOVNIHPRIKLJUCNOZATRAKTOR + brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR + brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR + brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO + brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            if (UKUPNOVANREDNIPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["O19"].Value = UKUPNOVANREDNIPRIKLJUCNOVOZILOZATRAKTOR;
            if (brISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["V19"].Value = brISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (brNEISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["W19"].Value = brNEISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILOZATRAKTOR;
            int UKUPNOKONTROLNIPRIKLJUCNOVOZILOZATRAKTOR = brISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILOZATRAKTOR + brNEISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (UKUPNOKONTROLNIPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["U19"].Value = UKUPNOKONTROLNIPRIKLJUCNOVOZILOZATRAKTOR;
            int UKUPNOPRIKLJUCNOVOZILOZATRAKTOR = UKUPNOREDOVNIPRIKLJUCNOVOZILOZATRAKTOR + UKUPNOSESTOMESECNIPRIKLJUCNOVOZILOZATRAKTOR + UKUPNOVANREDNIPRIKLJUCNOVOZILOZATRAKTOR + UKUPNOKONTROLNIPRIKLJUCNOVOZILOZATRAKTOR;
            if (UKUPNOPRIKLJUCNOVOZILOZATRAKTOR > 0)
                xlSheet.Range["B19"].Value = UKUPNOPRIKLJUCNOVOZILOZATRAKTOR;
            #endregion
            #region UKUPNOSVE
            int UKUPNO = UKUPNOPRIKLJUCNOVOZILOZATRAKTOR + UKUPNOMOPED + UKUPNOLAKITRICIKL + UKUPNOMOTOCIKL + UKUPNOTESKITRICIKL +
                UKUPNOLAKICETVOROCIKL + UKUPNOTESKICETVOROCIKL + UKUPNOPUTNICKOVOZILO + UKUPNOAUTOBUS + UKUPNOTERETNOVOZILO +
                UKUPNORADNAMASINA + UKUPNOMOTOKULTIVATOR + UKUPNOPRIKLJUCNOVOZILO + UKUPNOTRAKTOR;
            if (UKUPNO > 0)
                xlSheet.Range["B20"].Value = UKUPNO;

            int UKUPNOREDOVNIH = UKUPNOREDOVNIPRIKLJUCNOVOZILOZATRAKTOR + UKUPNOREDOVNIMOPED + UKUPNOREDOVNILAKITRICIKL + UKUPNOREDOVNIMOTOCIKL + UKUPNOREDOVNITESKITRICIKL +
                UKUPNOREDOVNILAKICETVOROCIKL + UKUPNOREDOVNITESKICETVOROCIKL + UKUPNOREDOVNIPUTNICKOVOZILO + UKUPNOREDOVNIAUTOBUS + UKUPNOREDOVNITERETNOVOZILO +
                UKUPNOREDOVNIRADNAMASINA + UKUPNOREDOVNIMOTOKULTIVATOR + UKUPNOREDOVNIPRIKLJUCNOVOZILO + UKUPNOREDOVNITRAKTOR;
            if (UKUPNOREDOVNIH > 0)
                xlSheet.Range["C20"].Value = UKUPNOREDOVNIH;
            int UKUPNOPREKID = brPREKINUTIHREDOVNIHMOPED + brPREKINUTIHREDOVNIHLAKITRICIKL + brPREKINUTIHREDOVNIHMOTOCIKL + brPREKINUTIHREDOVNIHTESKITRICIKL +
                brPREKINUTIHREDOVNIHLAKICETVOROCIKL + brPREKINUTIHREDOVNIHTESKICETVOROCIKL + brPREKINUTIHREDOVNIHPUTNICKOVOZILO + brPREKINUTIHREDOVNIHAUTOBUS + brPREKINUTIHREDOVNIHTERETNOVOZILO +
                brPREKINUTIHREDOVNIHRADNAMASINA + brPREKINUTIHREDOVNIHTRAKTOR + brPREKINUTIHREDOVNIHMOTOKULTIVATOR + brPREKINUTIHREDOVNIHPRIKLJUCNOVOZILO + brPREKINUTIHREDOVNIHPRIKLJUCNOZATRAKTOR;
            if (UKUPNOPREKID > 0)
                xlSheet.Range["D20"].Value = UKUPNOPREKID;
            int UKUPNOISPRAVNIHREDOVNIH = brISPRAVNIHREDOVNIHAUTOBUS + brISPRAVNIHREDOVNIHMOPED + brISPRAVNIHREDOVNIHLAKITRICIKL + brISPRAVNIHREDOVNIHMOTOCIKL + brISPRAVNIHREDOVNIHTESKITRICIKL +
                brISPRAVNIHREDOVNIHLAKICETVOROCIKL + brISPRAVNIHREDOVNIHTESKICETVOROCIKL + brISPRAVNIHREDOVNIHPUTNICKOVOZILO + brISPRAVNIHREDOVNIHTERETNOVOZILO + brISPRAVNIHREDOVNIHRADNAMASINA +
                brISPRAVNIHREDOVNIHTRAKTOR + brISPRAVNIHREDOVNIHMOTOKULTIVATOR + brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILO + brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (UKUPNOISPRAVNIHREDOVNIH > 0)
                xlSheet.Range["E20"].Value = UKUPNOISPRAVNIHREDOVNIH;
            int UKUPNONEISPRAVNIHREDOVNIH = brNEISPRAVNIHREDOVNIHAUTOBUS + brNEISPRAVNIHREDOVNIHMOPED + brNEISPRAVNIHREDOVNIHLAKITRICIKL + brNEISPRAVNIHREDOVNIHMOTOCIKL + brNEISPRAVNIHREDOVNIHTESKITRICIKL +
                brNEISPRAVNIHREDOVNIHLAKICETVOROCIKL + brNEISPRAVNIHREDOVNIHTESKICETVOROCIKL + brNEISPRAVNIHREDOVNIHPUTNICKOVOZILO + brNEISPRAVNIHREDOVNIHTERETNOVOZILO + brNEISPRAVNIHREDOVNIHRADNAMASINA +
                brNEISPRAVNIHREDOVNIHTRAKTOR + brNEISPRAVNIHREDOVNIHMOTOKULTIVATOR + brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILO + brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (UKUPNONEISPRAVNIHREDOVNIH > 0)
                xlSheet.Range["F20"].Value = UKUPNONEISPRAVNIHREDOVNIH;
            int UKUPNOISPRAVNIHREDOVNIHNOVO = brISPRAVNIHREDOVNIHAUTOBUSPONOVO + brISPRAVNIHREDOVNIHMOPEDPONOVO + brISPRAVNIHREDOVNIHLAKITRICIKLPONOVO + brISPRAVNIHREDOVNIHMOTOCIKLPONOVO + brISPRAVNIHREDOVNIHTESKITRICIKLPONOVO +
                brISPRAVNIHREDOVNIHLAKICETVOROCIKLPONOVO + brISPRAVNIHREDOVNIHTESKICETVOROCIKLPONOVO + brISPRAVNIHREDOVNIHPUTNICKOVOZILOPONOVO + brISPRAVNIHREDOVNIHTERETNOVOZILOPONOVO + brISPRAVNIHREDOVNIHRADNAMASINAPONOVO +
                brISPRAVNIHREDOVNIHTRAKTORPONOVO + brISPRAVNIHREDOVNIHMOTOKULTIVATORPONOVO + brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOPONOVO + brISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            if (UKUPNOISPRAVNIHREDOVNIHNOVO > 0)
                xlSheet.Range["G20"].Value = UKUPNOISPRAVNIHREDOVNIHNOVO;
            int UKUPNONEISPRAVNIHREDOVNIHNOVO = brNEISPRAVNIHREDOVNIHAUTOBUSPONOVO + brNEISPRAVNIHREDOVNIHMOPEDPONOVO + brNEISPRAVNIHREDOVNIHLAKITRICIKLPONOVO + brNEISPRAVNIHREDOVNIHMOTOCIKLPONOVO + brNEISPRAVNIHREDOVNIHTESKITRICIKLPONOVO +
                brNEISPRAVNIHREDOVNIHLAKICETVOROCIKLPONOVO + brNEISPRAVNIHREDOVNIHTESKICETVOROCIKLPONOVO + brNEISPRAVNIHREDOVNIHPUTNICKOVOZILOPONOVO + brNEISPRAVNIHREDOVNIHTERETNOVOZILOPONOVO + brNEISPRAVNIHREDOVNIHRADNAMASINAPONOVO +
                brNEISPRAVNIHREDOVNIHTRAKTORPONOVO + brNEISPRAVNIHREDOVNIHMOTOKULTIVATORPONOVO + brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOPONOVO + brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;

            if (UKUPNONEISPRAVNIHREDOVNIHNOVO > 0)
                xlSheet.Range["H20"].Value = UKUPNONEISPRAVNIHREDOVNIHNOVO;
            int UKUPNOSESTOMESECNIH = UKUPNOSESTOMESECNIPRIKLJUCNOVOZILOZATRAKTOR + UKUPNOSESTOMESECNIMOPED + UKUPNOSESTOMESECNILAKITRICIKL + UKUPNOSESTOMESECNIMOTOCIKL + UKUPNOSESTOMESECNITESKITRICIKL +
                UKUPNOSESTOMESECNILAKICETVOROCIKL + UKUPNOSESTOMESECNITESKICETVOROCIKL + UKUPNOSESTOMESECNIPUTNICKOVOZILO + UKUPNOSESTOMESECNIAUTOBUS + UKUPNOSESTOMESECNITERETNOVOZILO +
                UKUPNOSESTOMESECNIRADNAMASINA + UKUPNOSESTOMESECNIMOTOKULTIVATOR + UKUPNOSESTOMESECNIPRIKLJUCNOVOZILO + UKUPNOSESTOMESECNITRAKTOR;
            if (UKUPNOSESTOMESECNIH > 0)
                xlSheet.Range["I20"].Value = UKUPNOSESTOMESECNIH;
            int UKUPNOPREKIDSESTOMESECNIH = brPREKINUTIHSESTOMESECNIHMOPED + brPREKINUTIHSESTOMESECNIHLAKITRICIKL + brPREKINUTIHSESTOMESECNIHMOTOCIKL + brPREKINUTIHSESTOMESECNIHTESKITRICIKL +
                brPREKINUTIHSESTOMESECNIHLAKICETVOROCIKL + brPREKINUTIHSESTOMESECNIHTESKICETVOROCIKL + brPREKINUTIHSESTOMESECNIHPUTNICKOVOZILO + brPREKINUTIHSESTOMESECNIHAUTOBUS + brPREKINUTIHSESTOMESECNIHTERETNOVOZILO +
                brPREKINUTIHSESTOMESECNIHRADNAMASINA + brPREKINUTIHSESTOMESECNIHTRAKTOR + brPREKINUTIHSESTOMESECNIHMOTOKULTIVATOR + brPREKINUTIHSESTOMESECNIHPRIKLJUCNOVOZILO + brPREKINUTIHSESTOMESECNIHPRIKLJUCNOZATRAKTOR;
            if (UKUPNOPREKIDSESTOMESECNIH > 0)
                xlSheet.Range["J20"].Value = UKUPNOPREKIDSESTOMESECNIH;
            int UKUPNOISPRAVNIHSESTOMESECNIH = brISPRAVNIHSESTOMESECNIHAUTOBUS + brISPRAVNIHSESTOMESECNIHMOPED + brISPRAVNIHSESTOMESECNIHLAKITRICIKL + brISPRAVNIHSESTOMESECNIHMOTOCIKL + brISPRAVNIHSESTOMESECNIHTESKITRICIKL +
                brISPRAVNIHSESTOMESECNIHLAKICETVOROCIKL + brISPRAVNIHSESTOMESECNIHTESKICETVOROCIKL + brISPRAVNIHSESTOMESECNIHPUTNICKOVOZILO + brISPRAVNIHSESTOMESECNIHTERETNOVOZILO + brISPRAVNIHSESTOMESECNIHRADNAMASINA +
                brISPRAVNIHSESTOMESECNIHTRAKTOR + brISPRAVNIHSESTOMESECNIHMOTOKULTIVATOR + brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILO + brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (UKUPNOISPRAVNIHSESTOMESECNIH > 0)
                xlSheet.Range["K20"].Value = UKUPNOISPRAVNIHSESTOMESECNIH;
            int UKUPNONEISPRAVNIHSESTOMESECNIH = brNEISPRAVNIHSESTOMESECNIHAUTOBUS + brNEISPRAVNIHSESTOMESECNIHMOPED + brNEISPRAVNIHSESTOMESECNIHLAKITRICIKL + brNEISPRAVNIHSESTOMESECNIHMOTOCIKL + brNEISPRAVNIHSESTOMESECNIHTESKITRICIKL +
                brNEISPRAVNIHSESTOMESECNIHLAKICETVOROCIKL + brNEISPRAVNIHSESTOMESECNIHTESKICETVOROCIKL + brNEISPRAVNIHSESTOMESECNIHPUTNICKOVOZILO + brNEISPRAVNIHSESTOMESECNIHTERETNOVOZILO + brNEISPRAVNIHSESTOMESECNIHRADNAMASINA +
                brNEISPRAVNIHSESTOMESECNIHTRAKTOR + brNEISPRAVNIHSESTOMESECNIHMOTOKULTIVATOR + brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILO + brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (UKUPNONEISPRAVNIHSESTOMESECNIH > 0)
                xlSheet.Range["L20"].Value = UKUPNONEISPRAVNIHSESTOMESECNIH;
            int UKUPNOISPRAVNIHSESTOMESECNIHNOVO = brISPRAVNIHSESTOMESECNIHAUTOBUSPONOVO + brISPRAVNIHSESTOMESECNIHMOPEDPONOVO + brISPRAVNIHSESTOMESECNIHLAKITRICIKLPONOVO + brISPRAVNIHSESTOMESECNIHMOTOCIKLPONOVO + brISPRAVNIHSESTOMESECNIHTESKITRICIKLPONOVO +
                brISPRAVNIHSESTOMESECNIHLAKICETVOROCIKLPONOVO + brISPRAVNIHSESTOMESECNIHTESKICETVOROCIKLPONOVO + brISPRAVNIHSESTOMESECNIHPUTNICKOVOZILOPONOVO + brISPRAVNIHSESTOMESECNIHTERETNOVOZILOPONOVO + brISPRAVNIHSESTOMESECNIHRADNAMASINAPONOVO +
                brISPRAVNIHSESTOMESECNIHTRAKTORPONOVO + brISPRAVNIHSESTOMESECNIHMOTOKULTIVATORPONOVO + brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOPONOVO + brISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            if (UKUPNOISPRAVNIHSESTOMESECNIHNOVO > 0)
                xlSheet.Range["M20"].Value = UKUPNOISPRAVNIHSESTOMESECNIHNOVO;
            int UKUPNONEISPRAVNIHSESTOMESECNIHNOVO = brNEISPRAVNIHSESTOMESECNIHAUTOBUSPONOVO + brNEISPRAVNIHSESTOMESECNIHMOPEDPONOVO + brNEISPRAVNIHSESTOMESECNIHLAKITRICIKLPONOVO + brNEISPRAVNIHSESTOMESECNIHMOTOCIKLPONOVO + brNEISPRAVNIHSESTOMESECNIHTESKITRICIKLPONOVO +
                brNEISPRAVNIHSESTOMESECNIHLAKICETVOROCIKLPONOVO + brNEISPRAVNIHSESTOMESECNIHTESKICETVOROCIKLPONOVO + brNEISPRAVNIHSESTOMESECNIHPUTNICKOVOZILOPONOVO + brNEISPRAVNIHSESTOMESECNIHTERETNOVOZILOPONOVO + brNEISPRAVNIHSESTOMESECNIHRADNAMASINAPONOVO +
                brNEISPRAVNIHSESTOMESECNIHTRAKTORPONOVO + brNEISPRAVNIHSESTOMESECNIHMOTOKULTIVATORPONOVO + brNEISPRAVNIHSESTOMESECNIHPRIKLJUCNOVOZILOPONOVO + brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;

            if (UKUPNONEISPRAVNIHSESTOMESECNIHNOVO > 0)
                xlSheet.Range["N20"].Value = UKUPNONEISPRAVNIHSESTOMESECNIHNOVO;
            int UKUPNOVANREDNIH = UKUPNOVANREDNIPRIKLJUCNOVOZILOZATRAKTOR + UKUPNOVANREDNIMOPED + UKUPNOVANREDNILAKITRICIKL + UKUPNOVANREDNIMOTOCIKL + UKUPNOVANREDNITESKITRICIKL +
                UKUPNOVANREDNILAKICETVOROCIKL + UKUPNOVANREDNITESKICETVOROCIKL + UKUPNOVANREDNIPUTNICKOVOZILO + UKUPNOVANREDNIAUTOBUS + UKUPNOVANREDNITERETNOVOZILO +
                UKUPNOVANREDNIRADNAMASINA + UKUPNOVANREDNIMOTOKULTIVATOR + UKUPNOVANREDNIPRIKLJUCNOVOZILO + UKUPNOVANREDNITRAKTOR;
            if (UKUPNOVANREDNIH > 0)
                xlSheet.Range["O20"].Value = UKUPNOVANREDNIH;
            int UKUPNOPREKIDVANREDNIH = brPREKINUTIHVANREDOVNIHMOPED + brPREKINUTIHVANREDOVNIHLAKITRICIKL + brPREKINUTIHVANREDOVNIHMOTOCIKL + brPREKINUTIHVANREDOVNIHTESKITRICIKL +
                brPREKINUTIHVANREDOVNIHLAKICETVOROCIKL + brPREKINUTIHVANREDOVNIHTESKICETVOROCIKL + brPREKINUTIHVANREDOVNIHPUTNICKOVOZILO + brPREKINUTIHVANREDOVNIHAUTOBUS + brPREKINUTIHVANREDOVNIHTERETNOVOZILO +
                brPREKINUTIHVANREDOVNIHRADNAMASINA + brPREKINUTIHVANREDOVNIHTRAKTOR + brPREKINUTIHVANREDOVNIHMOTOKULTIVATOR + brPREKINUTIHVANREDOVNIHPRIKLJUCNOVOZILO + brPREKINUTIHVANREDOVNIHPRIKLJUCNOZATRAKTOR; ;
            if (UKUPNOPREKIDVANREDNIH > 0)
                xlSheet.Range["P20"].Value = UKUPNOPREKIDVANREDNIH;
            int UKUPNOISPRAVNIHVANREDNIH = brISPRAVNIHVANREDOVNIHAUTOBUS + brISPRAVNIHVANREDOVNIHMOPED + brISPRAVNIHVANREDOVNIHLAKITRICIKL + brISPRAVNIHVANREDOVNIHMOTOCIKL + brISPRAVNIHVANREDOVNIHTESKITRICIKL +
                brISPRAVNIHVANREDOVNIHLAKICETVOROCIKL + brISPRAVNIHVANREDOVNIHTESKICETVOROCIKL + brISPRAVNIHVANREDOVNIHPUTNICKOVOZILO + brISPRAVNIHVANREDOVNIHTERETNOVOZILO + brISPRAVNIHVANREDOVNIHRADNAMASINA +
                brISPRAVNIHVANREDOVNIHTRAKTOR + brISPRAVNIHVANREDOVNIHMOTOKULTIVATOR + brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILO + brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (UKUPNOISPRAVNIHVANREDNIH > 0)
                xlSheet.Range["Q20"].Value = UKUPNOISPRAVNIHVANREDNIH;
            int UKUPNONEISPRAVNIHVANREDOVNIH = brNEISPRAVNIHVANREDOVNIHAUTOBUS + brNEISPRAVNIHVANREDOVNIHMOPED + brNEISPRAVNIHVANREDOVNIHLAKITRICIKL + brNEISPRAVNIHVANREDOVNIHMOTOCIKL + brNEISPRAVNIHVANREDOVNIHTESKITRICIKL +
                brNEISPRAVNIHVANREDOVNIHLAKICETVOROCIKL + brNEISPRAVNIHVANREDOVNIHTESKICETVOROCIKL + brNEISPRAVNIHVANREDOVNIHPUTNICKOVOZILO + brNEISPRAVNIHVANREDOVNIHTERETNOVOZILO + brNEISPRAVNIHVANREDOVNIHRADNAMASINA +
                brNEISPRAVNIHVANREDOVNIHTRAKTOR + brNEISPRAVNIHVANREDOVNIHMOTOKULTIVATOR + brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILO + brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTOR;
            if (UKUPNONEISPRAVNIHVANREDOVNIH > 0)
                xlSheet.Range["R20"].Value = UKUPNONEISPRAVNIHVANREDOVNIH;
            int UKUPNOISPRAVNIHVANREDOVNIHNOVO = brISPRAVNIHVANREDOVNIHAUTOBUSPONOVO + brISPRAVNIHVANREDOVNIHMOPEDPONOVO + brISPRAVNIHVANREDOVNIHLAKITRICIKLPONOVO + brISPRAVNIHVANREDOVNIHMOTOCIKLPONOVO + brISPRAVNIHVANREDOVNIHTESKITRICIKLPONOVO +
                brISPRAVNIHVANREDOVNIHLAKICETVOROCIKLPONOVO + brISPRAVNIHVANREDOVNIHTESKICETVOROCIKLPONOVO + brISPRAVNIHVANREDOVNIHPUTNICKOVOZILOPONOVO + brISPRAVNIHVANREDOVNIHTERETNOVOZILOPONOVO + brISPRAVNIHVANREDOVNIHRADNAMASINAPONOVO +
                brISPRAVNIHVANREDOVNIHTRAKTORPONOVO + brISPRAVNIHVANREDOVNIHMOTOKULTIVATORPONOVO + brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOPONOVO + brISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;
            if (UKUPNOISPRAVNIHVANREDOVNIHNOVO > 0)
                xlSheet.Range["S20"].Value = UKUPNOISPRAVNIHVANREDOVNIHNOVO;
            int UKUPNONEISPRAVNIHVANREDOVNIHNOVO = brNEISPRAVNIHVANREDOVNIHAUTOBUSPONOVO + brNEISPRAVNIHVANREDOVNIHMOPEDPONOVO + brNEISPRAVNIHVANREDOVNIHLAKITRICIKLPONOVO + brNEISPRAVNIHVANREDOVNIHMOTOCIKLPONOVO + brNEISPRAVNIHVANREDOVNIHTESKITRICIKLPONOVO +
                brNEISPRAVNIHVANREDOVNIHLAKICETVOROCIKLPONOVO + brNEISPRAVNIHVANREDOVNIHTESKICETVOROCIKLPONOVO + brNEISPRAVNIHVANREDOVNIHPUTNICKOVOZILOPONOVO + brNEISPRAVNIHVANREDOVNIHTERETNOVOZILOPONOVO + brNEISPRAVNIHVANREDOVNIHRADNAMASINAPONOVO +
                brNEISPRAVNIHVANREDOVNIHTRAKTORPONOVO + brNEISPRAVNIHVANREDOVNIHMOTOKULTIVATORPONOVO + brNEISPRAVNIHVANREDOVNIHPRIKLJUCNOVOZILOPONOVO + brNEISPRAVNIHREDOVNIHPRIKLJUCNOVOZILOZATRAKTORPONOVO;

            if (UKUPNONEISPRAVNIHVANREDOVNIHNOVO > 0)
                xlSheet.Range["T20"].Value = UKUPNONEISPRAVNIHVANREDOVNIHNOVO;
            int UKUPNOKONTROLNIH = UKUPNOKONTROLNIPRIKLJUCNOVOZILOZATRAKTOR + UKUPNOKONTROLNIMOPED + UKUPNOKONTROLNILAKITRICIKL + UKUPNOKONTROLNIMOTOCIKL + UKUPNOKONTROLNITESKITRICIKL +
                UKUPNOKONTROLNILAKICETVOROCIKL + UKUPNOKONTROLNITESKICETVOROCIKL + UKUPNOKONTROLNIPUTNICKOVOZILO + UKUPNOKONTROLNIAUTOBUS + UKUPNOKONTROLNITERETNOVOZILO +
                UKUPNOKONTROLNIRADNAMASINA + UKUPNOKONTROLNIMOTOKULTIVATOR + UKUPNOKONTROLNIPRIKLJUCNOVOZILO + UKUPNOKONTROLNITRAKTOR;

            if (UKUPNOKONTROLNIH > 0)
                xlSheet.Range["U20"].Value = UKUPNOKONTROLNIH;
            int UKUPNOISPRAVNIHKONTROLNIH = brISPRAVNIHKONTROLNIHAUTOBUS + brISPRAVNIHKONTROLNIHMOPED + brISPRAVNIHKONTROLNIHLAKITRICIKL + brISPRAVNIHKONTROLNIHMOTOCIKL + brISPRAVNIHKONTROLNIHTESKITRICIKL +
                brISPRAVNIHKONTROLNIHLAKICETVOROCIKL + brISPRAVNIHKONTROLNIHTESKICETVOROCIKL + brISPRAVNIHKONTROLNIHPUTNICKOVOZILO + brISPRAVNIHKONTROLNIHTERETNOVOZILO + brISPRAVNIHKONTROLNIHRADNAMASINA +
                brISPRAVNIHKONTROLNIHTRAKTOR + brISPRAVNIHKONTROLNIHMOTOKULTIVATOR + brISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILO + brISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILOZATRAKTOR;

            if (UKUPNOISPRAVNIHKONTROLNIH > 0)
                xlSheet.Range["V20"].Value = UKUPNOISPRAVNIHKONTROLNIH;
            int UKUPNONEISPRAVNIHKONTROLNIH = brNEISPRAVNIHKONTROLNIHAUTOBUS + brNEISPRAVNIHKONTROLNIHMOPED + brNEISPRAVNIHKONTROLNIHLAKITRICIKL + brNEISPRAVNIHKONTROLNIHMOTOCIKL + brNEISPRAVNIHKONTROLNIHTESKITRICIKL +
                brNEISPRAVNIHKONTROLNIHLAKICETVOROCIKL + brNEISPRAVNIHKONTROLNIHTESKICETVOROCIKL + brNEISPRAVNIHKONTROLNIHPUTNICKOVOZILO + brNEISPRAVNIHKONTROLNIHTERETNOVOZILO + brNEISPRAVNIHKONTROLNIHRADNAMASINA +
                brNEISPRAVNIHKONTROLNIHTRAKTOR + brNEISPRAVNIHKONTROLNIHMOTOKULTIVATOR + brNEISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILO + brNEISPRAVNIHKONTROLNIHPRIKLJUCNOVOZILOZATRAKTOR;

            if (UKUPNONEISPRAVNIHKONTROLNIH > 0)
                xlSheet.Range["W20"].Value = UKUPNONEISPRAVNIHKONTROLNIH;
            #endregion

            xlApp.Visible = true;

            //clean up
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlSheet);
            Marshal.ReleaseComObject(xlBook);
            Marshal.ReleaseComObject(xlApp);
        }
        private int nadjiBrojPoVrstiRedovnih(char tip, string isp, List<vozilo> lista)
        {
            var listaPoTipu = lista.Where(x => x.IDBroj[8] == tip).ToList();
            int br = 0;
            for (int i = 0; i < listaPoTipu.Count; i++)
            {
                if (listaPoTipu[i].Ispravnost == isp)
                {
                    br++;
                }
            }
            return br;
        }
        private int nadjiBrojPoVrsti(char tip, string isp, List<vozilo> lista)
        {
            var listaPoTipu = lista.Where(x => x.IDBroj[8] == tip).ToList();
            int br = 0;
            for (int i = 0; i < listaPoTipu.Count; i++)
            {
                bool sem = false;
                if (listaPoTipu[i].Ispravnost == isp)
                {
                    for (int j = 0; j < i; j++)
                        if (listaPoTipu[i].BrojSasije == listaPoTipu[j].BrojSasije && listaPoTipu[j].Ispravnost == "Ne")
                            sem = true;
                    if (!sem)
                        br++;
                }
            }
            return br;
        }
        private int nadjiBrojPoVrstiPonovnih(char tip, string isp, List<vozilo> lista)
        {
            var listaPoTipu = lista.Where(x => x.IDBroj[8] == tip).ToList();
            int br = 0;
            for (int i = 0; i < listaPoTipu.Count; i++)
            {
                bool sem = false;
                if (listaPoTipu[i].Ispravnost == isp)
                {
                    for (int j = 0; j < i; j++)
                        if (listaPoTipu[i].BrojSasije == listaPoTipu[j].BrojSasije && listaPoTipu[j].Ispravnost == "Ne")
                            sem = true;
                    if (sem)
                        br++;
                }
            }
            return br;
        }

        private void rbTehIspPrekid_CheckedChanged(object sender, EventArgs e)
        {
            if (rbTehIspPrekid.Checked)
            {
                ispravnost = "Prekid";
                UkljuciIskljuciNeispravnosti();
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (prvoUcitavanje == false && /*((ComboBox)sender).SelectedItem !=null*/ cbSasijeEnter == true)
            {
                txtTPNap.Text = "";
                txtVezaID.Text = "";
                txtVezaGod.Text = "";
                vozilo v = cbSasije.SelectedItem as voziloSasije;

                txtTPID.Text = v.IDBroj;
                txtRegBr.Text = v.RegistarskiBroj;
                cbVrstaVoz.Text = v.VrstaVozila;
                cbVPV.Text = v.PodvrstaVozila;
                cbVKlasa.Text = v.klasa;
                cbVOblikKar.Text = v.oblikkaroserije;
                cbUredjajZaZaust.Text = v.zaustavljanje;
                cbVOKS.Text = v.OblikKaroserijeSkraceno;
                cbVSpec.Text = v.specijalno;
                if (cbVSpec.Text != "")
                    checkBox12.Checked = true;
                else
                    checkBox12.Checked = false;
                cbVMarka.Text = v.Marka;
                txtVModel.Text = v.Model;
                txtVTip.Text = v.Tip;
                txtVGodinaProizvodnje.Text = v.GodinaProizvodnje;
                txtVBrojSasije.Text = v.BrojSasije;
                txtVBrojMotora.Text = v.BrojMotora;
                cbVBoja.Text = v.Boja;
                txtVPrvaReg.Text = v.DatumPrveRegostracije;
                txtVOdnos.Text = v.Odnos;
                txtVSnaga.Text = v.Snaga;
                txtVZapremina.Text = v.Zapremina;
                txtVNajMasa.Text = v.NajvecaMasa;
                txtVTezina.Text = v.Tezina;
                txtVNosivost.Text = v.Nosivost;
                txtVSedenje1.Text = v.Sedenje;
                txtVStajanje.Text = v.Stajanje;
                txtVBrojOsovina.Text = v.BrojOsovina;
                cbVGorivo.Text = v.Gorivo;
                txtVHomologacija.Text = v.Homologacija;
                txtVIme.Text = v.ImeVL;
                txtVPrezime.Text = v.PrezimeVL;
                txtVJMBG.Text = v.MaticniBrojVL;
                txtVBRLK.Text = v.BRLKVL;
                txtVBrPoste.Text = v.BrojPosteVL;
                cbVMesto.Text = v.MestoVL;
                txtVUlica.Text = v.UlicaVL;
                txtVBroj.Text = v.BrojVL;
                txtVBrTelefona.Text = v.TelefonaVL;
                txtVZiroRacun.Text = v.ZiroRacunVL;
                txtDLIme.Text = v.ImeDL;
                txtDLPrezime.Text = v.PrezimeDL;
                txtDLJMBG.Text = v.MaticniBrojDL;
                txtDLBRLK.Text = v.BRLKDL;
                txtDLBrPoste.Text = v.BrojPosteDL;
                cbDLMesto.Text = v.MestoDL;
                txtDLUlica.Text = v.UlicaDL;
                txtDLBroj.Text = v.BrojDL;
                txtDLBrTelefona.Text = v.TelefonaDL;
                txtTPNap.Text = v.Napomena;
                txtTPDatum.Text = v.Datum;
                txtTPVremeOd.CustomFormat = "HH:mm";
                txtTPVremeDo.CustomFormat = "HH:mm";
                txtTPVremeDo.Format = DateTimePickerFormat.Custom;
                txtTPVremeOd.Format = DateTimePickerFormat.Custom;
                txtTPVremeOd.Text = v.VremeOd;
                txtTPVremeDo.Text = v.VremeDo;
                cbTPKontr1.Text = v.Kontrolor1;
                cbTPKontr2.Text = v.Kontrolor2;
                txtVKM.Text = v.Predjenikm;
                txtPoslednjeUverenjeBroj.Text = v.PoslednjeUverenjeBroj;
                txtPoslednjeUverenjeDatum.Text = v.PoslednjeUverenjeDatum;
                cbDecOzn.Text = v.DecimalnaOznaka;
                cbBojaL.Text = v.BojaL;
                cbBojaD.Text = v.BojaD;
                txtVD.Text = v.duzina;
                txtVS.Text = v.sirina;
                txtVV.Text = v.visina;
                txtVBrDokumenta.Text = v.BrojDokumenta;
                txtVSerijskiBroj.Text = v.SerijskiBroj;
                txtVBrSati.Text = v.BrSati;
                cbVKuka.Text = v.Kuka;
                txtVOpterecenjeN1.Text = v.OpterecenjeN1; txtVOpterecenjeN2.Text = v.OpterecenjeN2;
                txtVOpterecenjeN3.Text = v.OpterecenjeN3;
                txtVOpterecenjeN4.Text = v.OpterecenjeN4;
                txtVOpterecenjeN5.Text = v.OpterecenjeN5;
                txtVPneumatikL1.Text = v.PneumatikL1;
                txtVPneumatikL2.Text = v.PneumatikL2;
                txtVPneumatikL3.Text = v.PneumatikL3;
                txtVPneumatikL4.Text = v.PneumatikL4;
                txtVPneumatikL5.Text = v.PneumatikL5;
                txtVPneumatikD1.Text = v.PneumatikD1;
                txtVPneumatikD2.Text = v.PneumatikD2;
                txtVPneumatikD3.Text = v.PneumatikD3;
                txtVPneumatikD4.Text = v.PneumatikD4;
                txtVPneumatikD5.Text = v.PneumatikD5;
                cbUredjajZaZaust.Text = v.zaustavljanje;
                txtVKomentar.Text = v.komentar;
                txtVrstaPlusDecOznk.Text = v.decplusoznaka;
                txtVVlaznostKT.Text = v.VlaznostKT;
                datumIsteka.Text = v.datumIsteka;
                txtDLZiroRacun.Text = v.DLZiroRacun;
                tbDatumIzdavanja.Text = v.DatumIzdavanjaSaobracajne;
                //if (txtDLIme.Text != "")
                //    rbDL.Checked = true;
                //else
                //    rbVL.Checked = true;
                if (v.Prioritet == "DRUGO LICE")
                    rbDL.Checked = true;
                else
                    rbVL.Checked = true;

                radioButton2_CheckedChanged(sender, e);
                int a = 1;
                try
                {
                    a = int.Parse(v.rejting);
                }
                catch
                {
                    a = 1;
                }
                txtVRejting.Value = a;

                if (v.Ispravnost == "Da")
                {
                    rbTehIsp.Checked = true;
                    ispravnost = "Da";
                }
                else if (v.Ispravnost == "Ne")
                {
                    rbTehNeis.Checked = true;
                    ispravnost = "Ne";
                }
                else if (v.Ispravnost == "Prekid")
                {
                    rbTehIspPrekid.Checked = true;
                    ispravnost = "Prekid";
                }
                if (v.terensko == "Da")
                {
                    cbTerensko.Checked = true;
                    terensko = "Da";
                }
                else if (v.terensko == "Ne")
                {
                    cbTerensko.Checked = false;
                    terensko = "Ne";
                }
                popuniTabeluNeispravnosti(txtTPID.Text);
                cbVPV.Text = v.PodvrstaVozila;
                btnIzv.Visible = true;
                btnPot2.Visible = true;
                btnReg2.Visible = true;
                btnZap1.Visible = true;
                txtImePrezimePretraga.Text = "";
                if (v.ZiroRacunVL != "")
                    checkBox2.Checked = true;
                else
                    checkBox4.Checked = true;
                if (v.DLZiroRacun != "")
                    cbDLPravno.Checked = true;
                else
                    cbDLFizicko.Checked = true;
                popuniFakture();
                string datum = txtTPDatum.Text;
                proveriSlike(datum);
                oznaciVrstuPregleda(txtTPID.Text[8]);
                txtTPID.Text = v.IDBroj;
                cbIdjevi.SelectedIndex = -1;
                cbIdjevi.Text = "";
                cbTablice.SelectedIndex = -1;
                cbTablice.Text = "";

            }
        }

        private void cbSasije_Leave(object sender, EventArgs e)
        {
            cbSasijeEnter = false;
        }

        private void cbSasije_Enter(object sender, EventArgs e)
        {
            cbSasijeEnter = true;
        }

        private void cbIdjevi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (prvoUcitavanje == false && /*((ComboBox)sender).SelectedItem !=null*/ cbIdjeviEnter == true)
            {
                txtVezaID.Text = "";
                txtVezaGod.Text = "";
                txtTPNap.Text = "";
                voziloId v = cbIdjevi.SelectedItem as voziloId;

                txtTPID.Text = v.IDBroj;
                txtRegBr.Text = v.RegistarskiBroj;
                cbVrstaVoz.Text = v.VrstaVozila;
                cbVPV.Text = v.PodvrstaVozila;
                cbVKlasa.Text = v.klasa;
                cbVOblikKar.Text = v.oblikkaroserije;
                cbUredjajZaZaust.Text = v.zaustavljanje;
                cbVOKS.Text = v.OblikKaroserijeSkraceno;
                cbVSpec.Text = v.specijalno;
                if (cbVSpec.Text != "")
                    checkBox12.Checked = true;
                else
                    checkBox12.Checked = false;
                cbVMarka.Text = v.Marka;
                txtVModel.Text = v.Model;
                txtVTip.Text = v.Tip;
                txtVGodinaProizvodnje.Text = v.GodinaProizvodnje;
                txtVBrojSasije.Text = v.BrojSasije;
                txtVBrojMotora.Text = v.BrojMotora;
                cbVBoja.Text = v.Boja;
                txtVPrvaReg.Text = v.DatumPrveRegostracije;
                tbDatumIzdavanja.Text = v.DatumIzdavanjaSaobracajne;
                txtVOdnos.Text = v.Odnos;
                txtVSnaga.Text = v.Snaga;
                txtVZapremina.Text = v.Zapremina;
                txtVNajMasa.Text = v.NajvecaMasa;
                txtVTezina.Text = v.Tezina;
                txtVNosivost.Text = v.Nosivost;
                txtVSedenje1.Text = v.Sedenje;
                txtVStajanje.Text = v.Stajanje;
                txtVBrojOsovina.Text = v.BrojOsovina;
                cbVGorivo.Text = v.Gorivo;
                txtVHomologacija.Text = v.Homologacija;
                txtVIme.Text = v.ImeVL;
                txtVPrezime.Text = v.PrezimeVL;
                txtVJMBG.Text = v.MaticniBrojVL;
                txtVBRLK.Text = v.BRLKVL;
                txtVBrPoste.Text = v.BrojPosteVL;
                cbVMesto.Text = v.MestoVL;
                txtVUlica.Text = v.UlicaVL;
                txtVBroj.Text = v.BrojVL;
                txtVBrTelefona.Text = v.TelefonaVL;
                txtVZiroRacun.Text = v.ZiroRacunVL;
                txtDLIme.Text = v.ImeDL;
                txtDLPrezime.Text = v.PrezimeDL;
                txtDLJMBG.Text = v.MaticniBrojDL;
                txtDLBRLK.Text = v.BRLKDL;
                txtDLBrPoste.Text = v.BrojPosteDL;
                cbDLMesto.Text = v.MestoDL;
                txtDLUlica.Text = v.UlicaDL;
                txtDLBroj.Text = v.BrojDL;
                txtDLBrTelefona.Text = v.TelefonaDL;
                txtTPNap.Text = v.Napomena;
                txtTPDatum.Text = v.Datum;
                txtTPVremeOd.CustomFormat = "HH:mm";
                txtTPVremeDo.CustomFormat = "HH:mm";
                txtTPVremeDo.Format = DateTimePickerFormat.Custom;
                txtTPVremeOd.Format = DateTimePickerFormat.Custom;
                txtTPVremeOd.Text = v.VremeOd;
                txtTPVremeDo.Text = v.VremeDo;
                cbTPKontr1.Text = v.Kontrolor1;
                cbTPKontr2.Text = v.Kontrolor2;
                txtVKM.Text = v.Predjenikm;
                txtPoslednjeUverenjeBroj.Text = v.PoslednjeUverenjeBroj;
                txtPoslednjeUverenjeDatum.Text = v.PoslednjeUverenjeDatum;
                cbDecOzn.Text = v.DecimalnaOznaka;
                cbBojaL.Text = v.BojaL;
                cbBojaD.Text = v.BojaD;
                txtVD.Text = v.duzina;
                txtVS.Text = v.sirina;
                txtVV.Text = v.visina;
                txtVBrDokumenta.Text = v.BrojDokumenta;
                txtVSerijskiBroj.Text = v.SerijskiBroj;
                txtVBrSati.Text = v.BrSati;
                cbVKuka.Text = v.Kuka;
                txtVOpterecenjeN1.Text = v.OpterecenjeN1; txtVOpterecenjeN2.Text = v.OpterecenjeN2;
                txtVOpterecenjeN3.Text = v.OpterecenjeN3;
                txtVOpterecenjeN4.Text = v.OpterecenjeN4;
                txtVOpterecenjeN5.Text = v.OpterecenjeN5;
                txtVPneumatikL1.Text = v.PneumatikL1;
                txtVPneumatikL2.Text = v.PneumatikL2;
                txtVPneumatikL3.Text = v.PneumatikL3;
                txtVPneumatikL4.Text = v.PneumatikL4;
                txtVPneumatikL5.Text = v.PneumatikL5;
                txtVPneumatikD1.Text = v.PneumatikD1;
                txtVPneumatikD2.Text = v.PneumatikD2;
                txtVPneumatikD3.Text = v.PneumatikD3;
                txtVPneumatikD4.Text = v.PneumatikD4;
                txtVPneumatikD5.Text = v.PneumatikD5;
                cbUredjajZaZaust.Text = v.zaustavljanje;
                txtVVlaznostKT.Text = v.VlaznostKT;
                txtVKomentar.Text = v.komentar;
                txtVrstaPlusDecOznk.Text = v.decplusoznaka;
                datumIsteka.Text = v.datumIsteka;
                cbOsiguranje.Text = v.Osiguranje;

                //if (txtDLIme.Text != "")
                //    rbDL.Checked = true;
                //else
                //    rbVL.Checked = true;
                if (v.Prioritet == "DRUGO LICE")
                    rbDL.Checked = true;
                else
                    rbVL.Checked = true;
                radioButton2_CheckedChanged(sender, e);
                int a = 1;
                try
                {
                    a = int.Parse(v.rejting);
                }
                catch
                {
                    a = 1;
                }
                txtVRejting.Value = a;

                if (v.Ispravnost == "Da")
                {
                    rbTehIsp.Checked = true;
                    ispravnost = "Da";
                }
                else if (v.Ispravnost == "Ne")
                {
                    rbTehNeis.Checked = true;
                    ispravnost = "Ne";
                }
                else if (v.Ispravnost == "Prekid")
                {
                    rbTehIspPrekid.Checked = true;
                    ispravnost = "Prekid";
                }
                if (v.terensko == "Da")
                {
                    cbTerensko.Checked = true;
                    terensko = "Da";
                }
                else if (v.terensko == "Ne")
                {
                    cbTerensko.Checked = false;
                    terensko = "Ne";
                }
                popuniTabeluNeispravnosti(txtTPID.Text);
                cbVPV.Text = v.PodvrstaVozila;
                btnIzv.Visible = true;
                btnPot2.Visible = true;
                btnReg2.Visible = true;
                btnZap1.Visible = true;
                txtImePrezimePretraga.Text = "";
                txtDLZiroRacun.Text = v.DLZiroRacun;
                if (v.ZiroRacunVL != "")
                    checkBox2.Checked = true;
                else
                    checkBox4.Checked = true;
                if (v.DLZiroRacun != "")
                    cbDLPravno.Checked = true;
                else
                    cbDLFizicko.Checked = true;
                string datum = txtTPDatum.Text;
                proveriSlike(datum);
                popuniFakture();
                oznaciVrstuPregleda(txtTPID.Text[8]);
                txtTPID.Text = v.IDBroj;
                cbTablice.SelectedIndex = -1;
                cbTablice.Text = "";
                cbSasije.SelectedIndex = -1;
                cbSasije.Text = "";

            }
        }

        private void proveriSlike(string datum)
        {
            try
            {
                string putanja = @"D:\Slike\" + datum.Substring(6) + @"\" + datum.Substring(0, 2) + ". " + nadjiMesec(datum.Substring(3, 2)) + @"\" + int.Parse(cbIdjevi.Text.Substring(9));
                if (Directory.Exists(putanja))
                    btnSlike.Visible = true;
                else
                    btnSlike.Visible = false;
            }
            catch { }
        }

        private void cbIdjevi_Leave(object sender, EventArgs e)
        {
            cbIdjeviEnter = false;
        }

        private void cbIdjevi_Enter(object sender, EventArgs e)
        {
            cbIdjeviEnter = true;
        }

        private void dgwRacun_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            var index = e.RowIndex;
            string idbroj = dgwRacun.Rows[index].Cells[7].FormattedValue.ToString();
            try
            {
                vozilo v = b.id(idbroj.Trim());

                txtTPID.Text = v.IDBroj;
                txtRegBr.Text = v.RegistarskiBroj;
                cbVrstaVoz.Text = v.VrstaVozila;
                cbVPV.Text = v.PodvrstaVozila;
                cbVKlasa.Text = v.klasa;
                cbVOblikKar.Text = v.oblikkaroserije;
                cbUredjajZaZaust.Text = v.zaustavljanje;
                cbVOKS.Text = v.OblikKaroserijeSkraceno;
                cbVSpec.Text = v.specijalno;
                if (cbVSpec.Text != "")
                    checkBox12.Checked = true;
                else
                    checkBox12.Checked = false;
                cbVMarka.Text = v.Marka;
                txtVModel.Text = v.Model;
                txtVTip.Text = v.Tip;
                txtVGodinaProizvodnje.Text = v.GodinaProizvodnje;
                txtVBrojSasije.Text = v.BrojSasije;
                txtVBrojMotora.Text = v.BrojMotora;
                cbVBoja.Text = v.Boja;
                txtVPrvaReg.Text = v.DatumPrveRegostracije;
                txtVOdnos.Text = v.Odnos;
                txtVSnaga.Text = v.Snaga;
                txtVZapremina.Text = v.Zapremina;
                txtVNajMasa.Text = v.NajvecaMasa;
                txtVTezina.Text = v.Tezina;
                txtVNosivost.Text = v.Nosivost;
                txtVSedenje1.Text = v.Sedenje;
                txtVStajanje.Text = v.Stajanje;
                txtVBrojOsovina.Text = v.BrojOsovina;
                cbVGorivo.Text = v.Gorivo;
                txtVHomologacija.Text = v.Homologacija;
                txtVIme.Text = v.ImeVL;
                txtVPrezime.Text = v.PrezimeVL;
                txtVJMBG.Text = v.MaticniBrojVL;
                txtVBRLK.Text = v.BRLKVL;
                txtVBrPoste.Text = v.BrojPosteVL;
                cbVMesto.Text = v.MestoVL;
                txtVUlica.Text = v.UlicaVL;
                txtVBroj.Text = v.BrojVL;
                txtVBrTelefona.Text = v.TelefonaVL;
                txtVZiroRacun.Text = v.ZiroRacunVL;
                txtDLIme.Text = v.ImeDL;
                txtDLPrezime.Text = v.PrezimeDL;
                txtDLJMBG.Text = v.MaticniBrojDL;
                txtDLBRLK.Text = v.BRLKDL;
                txtDLBrPoste.Text = v.BrojPosteDL;
                cbDLMesto.Text = v.MestoDL;
                txtDLUlica.Text = v.UlicaDL;
                txtDLBroj.Text = v.BrojDL;
                txtDLBrTelefona.Text = v.TelefonaDL;
                txtTPNap.Text = v.Napomena;
                txtTPDatum.Text = v.Datum;
                txtTPVremeOd.CustomFormat = "HH:mm";
                txtTPVremeDo.CustomFormat = "HH:mm";
                txtTPVremeOd.Text = v.VremeOd;
                txtTPVremeDo.Text = v.VremeDo;
                cbTPKontr1.Text = v.Kontrolor1;
                cbTPKontr2.Text = v.Kontrolor2;
                txtVKM.Text = v.Predjenikm;
                txtPoslednjeUverenjeBroj.Text = v.PoslednjeUverenjeBroj;
                txtPoslednjeUverenjeDatum.Text = v.PoslednjeUverenjeDatum;
                cbDecOzn.Text = v.DecimalnaOznaka;
                cbBojaL.Text = v.BojaL;
                cbBojaD.Text = v.BojaD;

                txtVBrSati.Text = v.BrSati;
                cbVKuka.Text = v.Kuka;
                txtVOpterecenjeN1.Text = v.OpterecenjeN1; txtVOpterecenjeN2.Text = v.OpterecenjeN2;
                txtVOpterecenjeN3.Text = v.OpterecenjeN3;
                txtVOpterecenjeN4.Text = v.OpterecenjeN4;
                txtVOpterecenjeN5.Text = v.OpterecenjeN5;
                txtVPneumatikL1.Text = v.PneumatikL1;
                txtVPneumatikL2.Text = v.PneumatikL2;
                txtVPneumatikL3.Text = v.PneumatikL3;
                txtVPneumatikL4.Text = v.PneumatikL4;
                txtVPneumatikL5.Text = v.PneumatikL5;
                txtVPneumatikD1.Text = v.PneumatikD1;
                txtVPneumatikD2.Text = v.PneumatikD2;
                txtVPneumatikD3.Text = v.PneumatikD3;
                txtVPneumatikD4.Text = v.PneumatikD4;
                txtVPneumatikD5.Text = v.PneumatikD5;
                txtVVlaznostKT.Text = v.VlaznostKT;
                txtVBrDokumenta.Text = v.BrojDokumenta;
                cbUredjajZaZaust.Text = v.zaustavljanje;
                txtVKomentar.Text = v.komentar;
                txtVrstaPlusDecOznk.Text = v.decplusoznaka;
                datumIsteka.Text = v.datumIsteka;
                tbDatumIzdavanja.Text = v.DatumIzdavanjaSaobracajne;
                txtVVlaznostKT.Text = v.VlaznostKT;
                txtVBrDokumenta.Text = v.BrojDokumenta;
                txtVSerijskiBroj.Text = v.SerijskiBroj;
                
                if (txtDLIme.Text != "")
                    rbDL.Checked = true;
                else
                    rbVL.Checked = true;
                radioButton2_CheckedChanged(sender, e);
                int a = 1;
                try
                {
                    a = int.Parse(v.rejting);
                }
                catch
                {
                    a = 1;
                }
                txtVRejting.Value = a;

                if (v.Ispravnost == "Da")
                {
                    rbTehIsp.Checked = true;
                    ispravnost = "Da";
                }
                else if (v.Ispravnost == "Ne")
                {
                    rbTehNeis.Checked = true;
                    ispravnost = "Ne";
                }
                else if (v.Ispravnost == "Prekid")
                {
                    rbTehIspPrekid.Checked = true;
                    ispravnost = "Prekid";
                }
                if (v.terensko == "Da")
                {
                    cbTerensko.Checked = true;
                    terensko = "Da";
                }
                else if (v.terensko == "Ne")
                {
                    cbTerensko.Checked = false;
                    terensko = "Ne";
                }
                popuniTabeluNeispravnosti(txtTPID.Text);
                cbVPV.Text = v.PodvrstaVozila;
                btnIzv.Visible = true;
                btnPot2.Visible = true;
                btnReg2.Visible = true;
                btnZap1.Visible = true;
                txtImePrezimePretraga.Text = "";
                txtsasija.Clear();
                if (v.ZiroRacunVL != "")
                    checkBox2.Checked = true;
                else
                    checkBox4.Checked = true;
                popuniFakture();
                oznaciVrstuPregleda(txtTPID.Text[8]);
                txtTPID.Text = v.IDBroj;
                tabControl1.SelectedIndex = 0;
            }
            catch
            {
                MessageBox.Show("Nije pronadjen broj sasije");
                bunifuThinButton219_Click(sender, e);
            }
        }

        private void cbUredjajZaZaust_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
                cbTablice.Focus();
        }
        private static void PosaljiSmsCom(string broj, string reg, SmsComService sms)
        {
            if (string.IsNullOrWhiteSpace(broj))
                return;

            reg = reg
                .Replace('Š', 'S')
                .Replace('Č', 'C')
                .Replace('Ć', 'C')
                .Replace('Ž', 'Z')
                .Replace('Đ', 'D');

            if (broj.StartsWith("0"))
                broj = "+381" + broj.Substring(1);

            string poruka =
                $"Postovani, " +
                $"uskoro Vam istice registracija za vozilo {reg}. " +
                $"Posetite nas kako biste produzili registraciju na vreme. " +
                $"VAAS DOO SJENICA.";

            sms.SendSms(broj, poruka);
        }

        private static async Task Kredit()
        {
            try
            {
                HttpResponseMessage response = await client.GetAsync("http://login.balkan-cash.com/smsbulk/api/?user=vaasdoo&pass=J5kR7uQ4&action=credits");
                response.EnsureSuccessStatusCode();
                string responseBody = await response.Content.ReadAsStringAsync();
                // Above three lines can be replaced with new helper method below
                // string responseBody = await client.GetStringAsync(uri);

                MessageBox.Show("Trenutno imate : " + responseBody.Split(' ')[1] + " poruka");
            }
            catch (HttpRequestException ex)
            {
                MessageBox.Show("\nException Caught!");
                MessageBox.Show("Message :{0} ", ex.Message);
            }
        }

        static HttpClient client = new HttpClient();



        private async void bunifuThinButton226_Click(object sender, EventArgs e)
        {
            try
            {
                // ⚠️ AKO DOBIJEŠ SAMO "OK" – PROMENI COM PORT
                SmsComService sms = new SmsComService("COM7");

                sms.Open();

                string raw = await sms.SendUssdAsync("*101#");

                // PRIVREMENO – vidi ceo odgovor
                 //MessageBox.Show(raw, "RAW");

                string stanje = sms.ParseUssd(raw);

                MessageBox.Show(
                    stanje,
                    "Stanje prikazano na modulu",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                sms.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Greška",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }



       private async void bunifuThinButton225_Click(object sender, EventArgs e)
{
    // ✅ PROVERA: da li postoji bar jedna čekirana poruka
    bool imaCekiranih = false;

    foreach (DataGridViewRow row in dgwPoruke.Rows)
    {
        if (row.Cells[0].Value != null && Convert.ToBoolean(row.Cells[0].Value))
        {
            imaCekiranih = true;
            break;
        }
    }

    if (!imaCekiranih)
    {
        MessageBox.Show(
            "Molim vas označite poruku da bi je poslali",
            "Upozorenje",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        );
        return; // ⛔ prekid
    }

    // =================== SLANJE ===================

    SmsComService sms = new SmsComService("COM7"); // promeni COM port
    sms.Open();

    try
    {
        foreach (DataGridViewRow row in dgwPoruke.Rows)
        {
            if (row.Cells[0].Value == null) continue;

            bool cekirano = Convert.ToBoolean(row.Cells[0].Value);
            string broj = row.Cells[4].FormattedValue?.ToString();
            string reg = row.Cells[2].FormattedValue?.ToString();

            if (cekirano && !string.IsNullOrEmpty(broj))
            {
                PosaljiSmsCom(broj, reg, sms);
                Thread.Sleep(4000);

                string id = row.Cells[6].FormattedValue.ToString();
                b.posaljiPoruku(id);
            }
        }

        MessageBox.Show(
            "Uspešno slanje poruka",
            "Uspeh",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }
    catch (Exception ex)
    {
        MessageBox.Show(
            "Greška: " + ex.Message,
            "Greška",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        );
    }
    finally
    {
        sms.Close();
    }
}



        private void dgwPoruke_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgwPoruke.Rows[dgwPoruke.CurrentCell.RowIndex].Cells[0];
                if (bool.Parse(chk.Value.ToString()) == true)
                    chk.Value = false;
                else
                    chk.Value = true;
            }
        }
        private void txtDLJMBG_Leave(object sender, EventArgs e)
        {
            radioButton2_CheckedChanged(sender, e);
        }

        private void txtVJMBG_Leave(object sender, EventArgs e)
        {
            radioButton2_CheckedChanged(sender, e);
        }

        private void comboBox3_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //if (!prvoUcitavanje)
            //{
            Faktura f = cbFakture.SelectedItem as Faktura;
            textBox1.Text = f.NazivFirme;
            textBox2.Text = f.Ulica;
            textBox3.Text = f.Grad;
            textBox4.Text = f.PIB;
            textBox6.Text = f.Telefon;
            textBox7.Text = f.TR;
            textBox20.Text = f.MaticniBroj;
            //txtNaziv.Text = f.NazivArtikla;
            //txtSfr.Text = f.Cena;
            //Kom.Value = int.Parse(f.Komada);
            //txtRabat.Text = f.Rabat;
            txtNaziv.Text = f.NazivArtikla;
            txtSfr.Text = f.Cena;

            //}
            //else
            //{
            //    textBox1.Text = "";
            //    textBox2.Text = "";
            //    textBox3.Text = "";
            //    textBox4.Text = "";
            //    textBox6.Text = "";
            //    textBox7.Text = "";
            //    textBox20.Text = "";
            //    //txtNaziv.Text = f.NazivArtikla;
            //    //txtSfr.Text = f.Cena;
            //    //Kom.Value = int.Parse(f.Komada);
            //    //txtRabat.Text = f.Rabat;
            //    txtNaziv.Text = "";
            //    txtSfr.Text = "";
            //}
        }

        private void bunifuThinButton223_Click(object sender, EventArgs e)
        {
            int n;

            if (string.IsNullOrEmpty(txtVBrojSasije.Text.Replace("/", "")))
            {
                MessageBox.Show("Nije unet broj sasije!");
            }
            else
            {
                try
                {
                    Process.Start(@"D:\Slike\Uverenja\" + txtVBrojSasije.Text.Replace("/", "") + ".pdf");
                }
                catch (Exception)
                {
                    MessageBox.Show("Ne postoji uverenje za ovo vozilo!");
                }
            }
        }

        private void bunifuThinButton223_Click_1(object sender, EventArgs e)
        {
            vozilo v = new vozilo
            {
                ImeDL = txtVIme.Text,
                PrezimeDL = txtVPrezime.Text,
                UlicaDL = txtVUlica.Text,
                BrojDL = txtVBroj.Text,
                MestoDL = cbVMesto.Text,
                MaticniBrojDL = txtVJMBG.Text,
                SerijskiBroj = txtVSerijskiBroj.Text,
                BrojDokumenta = txtVBrDokumenta.Text,
                RegistarskiBroj = txtRegBr.Text,
                VrstaVozila = cbVrstaVoz.Text,
                Gorivo = cbVGorivo.Text,
                Zapremina = txtVZapremina.Text,
                Sedenje = txtVSedenje1.Text,
                NajvecaMasa = txtVNajMasa.Text,
                Nosivost = txtVNosivost.Text,
                GodinaProizvodnje = txtVGodinaProizvodnje.Text,
                Snaga = txtVSnaga.Text
            };
            if (checkBox2.Checked)
            {
                v.MaticniBrojDL = txtVBRLK.Text;
                v.ImeDL = v.PrezimeDL;
                v.PrezimeDL = "";
            }
            ePlati ep = new ePlati(v);
            ep.ShowDialog();
        }

        private void cbOsiguranje_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label129_Click(object sender, EventArgs e)
        {
            if (txtVKM.Text != "" && cbVrstaVoz.Text != "" && (checkBox12.Checked && cbVOblikKar.Text == "" ? true : (cbVOblikKar.Text == "" ? false : true)) && cbVMarka.Text != "" && txtVModel.Text != "" && txtVGodinaProizvodnje.Text != ""
                && txtVBrojSasije.Text != "" && txtVBrojMotora.Text != "" && cbVBoja.Text != "" && cbBojaL.Text != "" && cbBojaD.Text != "" && txtVPrvaReg.Text != ""
                && txtVSnaga.Text != "" && txtVZapremina.Text != "" && txtVNajMasa.Text != "" && txtVTezina.Text != "" && txtVSedenje1.Text != "" && txtVStajanje.Text != ""
                && txtVBrojOsovina.Text != "" && cbVGorivo.Text != "" && txtTPID.Text != "" && ispravnost != "" && cbVPV.Text != "" && (checkBox12.Checked && cbVOKS.Text == "" ? true : (cbVOKS.Text == "" ? false : true))
                && txtTPVremeOd.Text != "  :" && txtTPVremeDo.Text != "  :")
            {
                int n;
                if (txtTPID.Text.Length < 14 || txtTPID.Text.Length > 14)
                {
                    MessageBox.Show("ID Broj je nepravilan!");
                }
                else if (int.TryParse(txtTPID.Text.Substring(9), out n) == false)
                {
                    MessageBox.Show("ID Broj je nepravilan!");
                }

                else if (rbTehNeis.Checked == true)
                {
                    MessageBox.Show("Tehnička ispranost vozila je obeležena sa \'Ne\'!");
                }

                else
                {
                    lblVrstaVoz.ForeColor = SystemColors.ControlText; lblVOblikKar.ForeColor = SystemColors.ControlText; lblVMarka.ForeColor = SystemColors.ControlText;
                    lblVModel.ForeColor = SystemColors.ControlText; lblVGodinaProizvodnje.ForeColor = SystemColors.ControlText; lblVBrojSasije.ForeColor = SystemColors.ControlText;
                    lblVBrojMotora.ForeColor = SystemColors.ControlText; lblVBoja.ForeColor = SystemColors.ControlText; lblVPrvaReg.ForeColor = SystemColors.ControlText;
                    lblVSnaga.ForeColor = SystemColors.ControlText; lblVZapremina.ForeColor = SystemColors.ControlText; lblVNajMasa.ForeColor = SystemColors.ControlText;
                    lblVTezina.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText;
                    lblVBrojOsovina.ForeColor = SystemColors.ControlText; lblVGorivo.ForeColor = SystemColors.ControlText; lblTPID.ForeColor = SystemColors.ControlText;
                    lblIspravnost.ForeColor = SystemColors.ControlText; lblPredjenikm.ForeColor = SystemColors.ControlText; lblVremeOd.ForeColor = SystemColors.ControlText;
                    lblVremeDo.ForeColor = SystemColors.ControlText;

                    PrintDialog printDialog = new PrintDialog();

                    PrintDocument printDocument = new PrintDocument();

                    printDialog.Document = printDocument; //add the document to the dialog box...        

                    printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt21); //add an event handler that will do the printing

                    //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

                    DialogResult result = printDialog.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        printDocument.Print();

                    }
                }
            }
            else
            {
                MessageBox.Show("Popunite sva obavezna polja na tabovima \'Podaci o vozilu\' i \'Tehnički pregled\'.");

                lblVrstaVoz.ForeColor = Color.Red; lblVOblikKar.ForeColor = Color.Red; lblVMarka.ForeColor = Color.Red; lblVModel.ForeColor = Color.Red;
                lblVGodinaProizvodnje.ForeColor = Color.Red; lblVBrojSasije.ForeColor = Color.Red; lblVBrojMotora.ForeColor = Color.Red; lblVBoja.ForeColor = Color.Red;
                lblVPrvaReg.ForeColor = Color.Red; lblVSnaga.ForeColor = Color.Red; lblVZapremina.ForeColor = Color.Red; lblVNajMasa.ForeColor = Color.Red;
                lblVTezina.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVBrojOsovina.ForeColor = Color.Red;
                lblVGorivo.ForeColor = Color.Red; lblTPID.ForeColor = Color.Red; lblPredjenikm.ForeColor = Color.Red; lblIspravnost.ForeColor = Color.Red;
                lblVremeOd.ForeColor = Color.Red; lblVremeDo.ForeColor = Color.Red;
            }

        }

        private void cbDLPravno_CheckedChanged(object sender, EventArgs e)
        {
            if (cbDLPravno.Checked)
            {
                txtDLZiroRacun.Enabled = true;
                cbDLFizicko.Checked = false;
                groupBox3.Visible = false;
                groupBox2.Visible = true;
            }
        }

        private void cbDLFizicko_CheckedChanged(object sender, EventArgs e)
        {
            if (cbDLFizicko.Checked)
            {
                txtDLZiroRacun.Enabled = false;
                cbDLPravno.Checked = false;
                groupBox3.Visible = true;
                groupBox2.Visible = false;
            }
        }

        private void txtVZiroRacun_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void txtDLZiroRacun_TextChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
                textBox15.Text = txtVZiroRacun.Text;
            else
                textBox15.Text = txtDLZiroRacun.Text;

            popuniFakture();
        }
        public void CreateReceipt22(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int height = 0;
            int xpos = 30, ypos = 30;
            int i = 0;
            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.5f);
            Font FontNaznacenimanji = new Font("Times New Roman", 11, FontStyle.Bold);
            Font FontNaznaceni = new Font("Times New Roman", 11, FontStyle.Bold);
            Font FontNajjaci = new Font("Times New Roman", 14, FontStyle.Bold | FontStyle.Underline);
            Font Fontoznaka = new Font("Times New Roman", 10);
            Font Fontobicni = new Font("Times New Roman", 10);
            Font Fontobicni2 = new Font("Times New Roman", 10);
            Font fontUpisivanje = new Font("Times New Roman", 10, FontStyle.Italic);
            Font za_licnu_kartu = new Font("Times New Roman", 8);
            Font FontSvoji = new Font("Times New Roman", 14);
            Font fontmali = new Font("Times New Roman", 8);

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;

            StringFormat stringFormat2 = new StringFormat();
            stringFormat2.Alignment = StringAlignment.Near;
            stringFormat2.LineAlignment = StringAlignment.Center;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;

            //stampanje svojih podataka
            //e.Graphics.DrawString("ALSAZE 2016 DOO, Sjenica", FontSvoji, Brushes.Black, new Rectangle(xpos,ypos, 400, 20));
            //e.Graphics.DrawString("1 8 9 0 4 6", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 794, 15), stringFormat);
            e.Graphics.DrawString("Obrazac 1", Fontoznaka, Brushes.Black, new Rectangle(xpos, ypos, 750, 30), stringFormatdesno);

            ypos += 40;
            e.Graphics.DrawString("REGISTRACIONI LIST ZA MOTORNO-PRIKLJUČNO VOZILO", FontSvoji, Brushes.Black, new Rectangle(xpos, ypos, 794, 20), stringFormat);
            ypos += 65;
            //stampanje podataka o kupcu

            e.Graphics.DrawString("PODACI O VOZILU", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            ypos += 50;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(J)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Vrsta vozila", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbVrstaVoz.Text + " " + cbVPV.Text + ((cbTerensko.Checked) ? "G" : ""), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(D.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Marka", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbVMarka.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(D.2)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Tip/varijanta/verzija", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVTip.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 420, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(D.3)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Komercijalna oznaka (model)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVModel.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(E)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Broj šasije", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVBrojSasije.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(F.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Najveća dozvoljena masa (kg)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVNajMasa.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(G)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Masa", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVTezina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(K)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj homologacije", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVHomologacija.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(P.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Radna zapremina motora", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 175, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVZapremina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(P.2)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Snaga motora u kW", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVSnaga.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(P.3)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Vrsta goriva ili pogona", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbVGorivo.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(Q)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Odnos snaga/masa u kw/kg (samo za motocikle)", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVOdnos.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(S.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj mesta za sedenje uključujući i mesto vozača", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVSedenje1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(S.2)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj mesta za stajanje", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVStajanje.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(P.5)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj motora", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVBrojMotora.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(L)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Broj osovina", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVBrojOsovina.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(R)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Boja vozila", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(cbBojaL.Text.Substring(0, 1) + " " + cbVBoja.Text.ToString() + " " + cbBojaD.Text.Substring(0, 1), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 35));
            e.Graphics.DrawString("(B.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 35));
            e.Graphics.DrawString("Godina proizvodnje vozila", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 35));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 35));
            e.Graphics.DrawString(txtVGodinaProizvodnje.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 35));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(J.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Oblik karoserije", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            //string s="";
            //if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
            //{ 
            //    if(!checkBox12.Checked)
            //        s = cbVKlasa.Text + "/" + cbVOKS.Text;
            //    else
            //        s = cbVKlasa.Text + "/" + (cbVSpec.Text.Split(' '))[0].ToString();
            //}

            //else
            //     if (checkBox12.Checked)
            //{
            //    if (cbVrstaVoz.Text == "AUTOBUS" && cbVKlasa.SelectedIndex > -1)
            //        s = cbVKlasa.Text + "/" + (cbVSpec.Text.Split(' '))[0].ToString();
            //    else
            //        if (cbDecOzn.Text != "")
            //        s = (cbVSpec.Text.Split(' '))[0].ToString() + (cbDecOzn.Text.Split(' '))[0].ToString();
            //    else
            //        s = (cbVSpec.Text.Split(' '))[0].ToString();
            //    txtVrstaPlusDecOznk.Text = s;
            //}
            //else
            //    if (cbDecOzn.Text != "")
            //    s = cbVOKS.Text + (cbDecOzn.Text.Split(' '))[0].ToString();

            //else
            //    s = cbVOKS.Text;
            e.Graphics.DrawString(txtVrstaPlusDecOznk.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 25;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("(G.1)", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 25, ypos, 150, 20));
            e.Graphics.DrawString("Nosivost", Fontobicni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos + 175, ypos, 175, 20));
            e.Graphics.DrawString(txtVNosivost.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 30;

            e.Graphics.DrawString("PODACI O TEHNIČKOM PREGLEDU", FontNaznaceni, Brushes.Black, new Rectangle(xpos + 2, ypos, 400, 20));
            ypos += 30;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Naziv i sedište ovlašćenog pravnog lica :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            e.Graphics.DrawString("VAAS d.o.o. Sjenica", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 20;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Šifra :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawString("141022", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 20;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("ID broj :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);

            //e.Graphics.DrawString((int.Parse(txtTPID.Text.Substring(9))).ToString(), Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 30;

            string tekst1 = "VOZILO JE TEHNIČKI ISPRAVNO";
            Rectangle rect1 = new Rectangle(xpos + 2, ypos, 400, 20);

            // ispis teksta
            e.Graphics.DrawString(tekst1, FontNaznaceni, Brushes.Black, rect1);

            // precrtavanje (horizontalna linija kroz sredinu teksta)
            int yLine1 = ypos + (FontNaznaceni.Height / 2);
            e.Graphics.DrawLine(Pens.Black, rect1.Left, yLine1, rect1.Right - 145, yLine1);

            ypos += 30;

            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            string tekst2 = "Datum vršenja tehničkog pregleda :";
            Rectangle rect2 = new Rectangle(xpos + 1, ypos, 400, 20);

            // ispis teksta
            e.Graphics.DrawString(tekst2, Fontobicni, Brushes.Black, rect2, stringFormatdesno);

            // precrtavanje
            int yLine2 = ypos + (Fontobicni.Height / 2);
            e.Graphics.DrawLine(Pens.Black, rect2.Left + 200, yLine2, rect2.Right - 5, yLine2);
            //e.Graphics.DrawString(txtTPDatum.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 30;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Kontrolor :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            //e.Graphics.DrawString(cbTPKontr1.Text, Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 400, 20));
            ypos += 50;
            //e.Graphics.DrawRectangle(pe, new Rectangle(xpos, ypos, 25, 20));
            e.Graphics.DrawString("Napomena :", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);
            e.Graphics.DrawString("ODJAVA VOZILA", Fontobicni, Brushes.Black, new Rectangle(xpos + 2 + 420, ypos, 350, 100));
            ypos += 120;
            e.Graphics.DrawString("MP", Fontobicni, Brushes.Black, new Rectangle(xpos + 1, ypos, 400, 20), stringFormatdesno);


        }
        private void label129_Click_1(object sender, EventArgs e)
        {
            if (cbVrstaVoz.Text != "" && (checkBox12.Checked && cbVOblikKar.Text == "" ? true : (cbVOblikKar.Text == "" ? false : true)) && cbVMarka.Text != "" && txtVModel.Text != "" && txtVGodinaProizvodnje.Text != ""
              && txtVBrojSasije.Text != "" && txtVBrojMotora.Text != "" && cbVBoja.Text != "" && cbBojaL.Text != "" && cbBojaD.Text != "" && txtVPrvaReg.Text != ""
              && txtVSnaga.Text != "" && txtVZapremina.Text != "" && txtVNajMasa.Text != "" && txtVTezina.Text != "" && txtVSedenje1.Text != "" && txtVStajanje.Text != ""
              && txtVBrojOsovina.Text != "" && cbVGorivo.Text != "" && txtTPID.Text != "" && ispravnost != "" && (checkBox12.Checked && cbVOKS.Text == "" ? true : (cbVOKS.Text == "" ? false : true))
             )
            {
                int n;
                if (txtTPID.Text.Length < 14 || txtTPID.Text.Length > 14)
                {
                    MessageBox.Show("ID Broj je nepravilan!");
                }
                else if (int.TryParse(txtTPID.Text.Substring(9), out n) == false)
                {
                    MessageBox.Show("ID Broj je nepravilan!");
                }

                else if (rbTehNeis.Checked == true)
                {
                    MessageBox.Show("Tehnička ispranost vozila je obeležena sa \'Ne\'!");
                }

                else
                {
                    lblVrstaVoz.ForeColor = SystemColors.ControlText; lblVOblikKar.ForeColor = SystemColors.ControlText; lblVMarka.ForeColor = SystemColors.ControlText;
                    lblVModel.ForeColor = SystemColors.ControlText; lblVGodinaProizvodnje.ForeColor = SystemColors.ControlText; lblVBrojSasije.ForeColor = SystemColors.ControlText;
                    lblVBrojMotora.ForeColor = SystemColors.ControlText; lblVBoja.ForeColor = SystemColors.ControlText; lblVPrvaReg.ForeColor = SystemColors.ControlText;
                    lblVSnaga.ForeColor = SystemColors.ControlText; lblVZapremina.ForeColor = SystemColors.ControlText; lblVNajMasa.ForeColor = SystemColors.ControlText;
                    lblVTezina.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText; lblVStajanje.ForeColor = SystemColors.ControlText;
                    lblVBrojOsovina.ForeColor = SystemColors.ControlText; lblVGorivo.ForeColor = SystemColors.ControlText; lblTPID.ForeColor = SystemColors.ControlText;
                    lblIspravnost.ForeColor = SystemColors.ControlText;

                    PrintDialog printDialog = new PrintDialog();

                    PrintDocument printDocument = new PrintDocument();

                    printDialog.Document = printDocument; //add the document to the dialog box...        

                    printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(CreateReceipt22); //add an event handler that will do the printing

                    //on a till you will not want to ask the user where to print but this is fine for the test envoironment.

                    DialogResult result = printDialog.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        printDocument.Print();

                    }
                }
            }
            else
            {
                MessageBox.Show("Popunite sva obavezna polja na tabovima \'Podaci o vozilu\' i \'Tehnički pregled\'.");

                lblVrstaVoz.ForeColor = Color.Red; lblVOblikKar.ForeColor = Color.Red; lblVMarka.ForeColor = Color.Red; lblVModel.ForeColor = Color.Red;
                lblVGodinaProizvodnje.ForeColor = Color.Red; lblVBrojSasije.ForeColor = Color.Red; lblVBrojMotora.ForeColor = Color.Red; lblVBoja.ForeColor = Color.Red;
                lblVPrvaReg.ForeColor = Color.Red; lblVSnaga.ForeColor = Color.Red; lblVZapremina.ForeColor = Color.Red; lblVNajMasa.ForeColor = Color.Red;
                lblVTezina.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVStajanje.ForeColor = Color.Red; lblVBrojOsovina.ForeColor = Color.Red;
                lblVGorivo.ForeColor = Color.Red; lblTPID.ForeColor = Color.Red; lblIspravnost.ForeColor = Color.Red;
                ;
            }
        }

        private void dgwRacun_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        /*
static async Task<DownloadPageAsyncResult> DownloadPageAsync()
{
   var result = new DownloadPageAsyncResult();
   try
   {
       using (HttpResponseMessage response = await client.GetAsync(new Uri("http://login.balkancash.com/smsbulk/api/?user=vaasdoo&pass=J5kR7uQ4&action=send&display=Test&text=tekst % 20poruke & numbers =% 2B381607423480")))
       {
           using (HttpContent content = response.Content)
           {
               // need these to return to Form for display
               string resultString = await content.ReadAsStringAsync();
               string reasonPhrase = response.ReasonPhrase;
               HttpResponseHeaders headers = response.Headers;
               HttpStatusCode code = response.StatusCode;

               result.result = resultString;
               result.reasonPhrase = reasonPhrase;
               result.headers = headers;
               result.code = code;
           }
       }
   }
   catch (Exception ex)
   {
       // need to return ex.message for display.
       result.errorMessage = ex.Message;
   }
   return result;
}
*/
        private void txtVBrojMotora_TextChanged(object sender, EventArgs e)
        {
            lblDuzinaBrmotora.Text = txtVBrojMotora.Text.Length.ToString();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            string jmbg;
            if (rbDL.Checked)
            {
                gbDL.Enabled = true;
                gbVl.Enabled = false;
                //txtVIme.Clear();
                //txtVPrezime.Clear();
                //txtVJMBG.Clear();
                //txtVBRLK.Clear();
                //txtVBrPoste.Clear();
                //cbVMesto.Clear();
                //txtVUlica.Clear();
                //txtVBroj.Clear();
                //txtVBrTelefona.Clear();
                //txtVZiroRacun.Clear();
                jmbg = txtDLJMBG.Text;
            }
            else
            {
                gbVl.Enabled = true;
                gbDL.Enabled = false;
                //txtDLIme.Clear();
                //txtDLPrezime.Clear();
                //txtDLJMBG.Clear();
                //txtDLBRLK.Clear();
                //txtDLBrPoste.Clear();
                //cbDLMesto.Clear();
                //txtDLUlica.Clear();
                //txtDLBroj.Clear();
                //txtDLBrTelefona.Clear();
                jmbg = txtVJMBG.Text;
            }
            var broj = b.izracunajPreglede(jmbg);
            txtBr.Text = (!string.IsNullOrEmpty(jmbg)) ? broj : "0";
            if (!string.IsNullOrEmpty(jmbg) && int.Parse(broj) % 5 == 0)
                txtBr.BackColor = Color.LightPink;
            else
                txtBr.BackColor = Color.FromArgb(240, 240, 240);
        }

        private void bunifuThinButton24_Click_1(object sender, EventArgs e)
        {
            try
            {
                var data = ElectronicIdReader.ReadAll();
                txtVJMBG.Text = data.FixedPersonalData.PersonalNumber;
                txtVIme.Text = data.FixedPersonalData.GivenName;
                txtVPrezime.Text = data.FixedPersonalData.Surname;
                cbVMesto.Text = data.VariablePersonalData.Place + ", " + data.VariablePersonalData.Community;
                txtVUlica.Text = data.VariablePersonalData.Street;
                txtVBroj.Text =
             data.VariablePersonalData.HouseNumber +
             data.VariablePersonalData.HouseLetter +
             (string.IsNullOrWhiteSpace(data.VariablePersonalData.Floor)
                 ? ""
                 : "/" + data.VariablePersonalData.Floor) +
             (string.IsNullOrWhiteSpace(data.VariablePersonalData.ApartmentNumber)
                 ? ""
                 : "/" + data.VariablePersonalData.ApartmentNumber);

                txtVBRLK.Text = data.Document.RegistrationNumber;
                radioButton2_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show("Molim vas ubacite ličnu kartu");

            }
        }

        private void bunifuThinButton25_Click_1(object sender, EventArgs e)
        {
            try
            {
                var data = ElectronicIdReader.ReadAll();
                txtDLJMBG.Text = data.FixedPersonalData.PersonalNumber;
                txtDLIme.Text = data.FixedPersonalData.GivenName;
                txtDLPrezime.Text = data.FixedPersonalData.Surname;
                cbDLMesto.Text = data.VariablePersonalData.Place + ", " + data.VariablePersonalData.Community;
                txtDLUlica.Text = data.VariablePersonalData.Street;
                txtDLBroj.Text =
     data.VariablePersonalData.HouseNumber +
     data.VariablePersonalData.HouseLetter +
     (string.IsNullOrWhiteSpace(data.VariablePersonalData.Floor)
         ? ""
         : "/" + data.VariablePersonalData.Floor) +
     (string.IsNullOrWhiteSpace(data.VariablePersonalData.ApartmentNumber)
         ? ""
         : "/" + data.VariablePersonalData.ApartmentNumber);

                txtDLBRLK.Text = data.Document.RegistrationNumber;
                radioButton2_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show("Molim vas ubacite ličnu kartu");

            }
        }








        //S-OBIČAN
        //    E-METALNI
        //    F-FOLIJA



    }
}