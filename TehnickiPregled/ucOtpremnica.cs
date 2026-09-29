using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Globalization;
using System.Drawing.Drawing2D;

namespace TehnickiPregled
{
    public partial class ucOtpremnica : UserControl
    {
        Broker b = new Broker();
        int broj_optremnice;
        int strana = 0;
        DataTable dt = new DataTable();
        int i = 0, suma = 0;
        double kolicina = 1.0;
        double ukupno = 0, ukupno_bez_rabata = 0, ukupan_rabat = 0, ukupan_pdv = 0;
        

        private void btnPrintReciept_Click(object sender, EventArgs e)
        {
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

        public void CreateReceipt(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            i = 0;
            strana++;
            int pocetak = 10;
            int sredina = 400;
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
            dataGridView1.Font = tekst;

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Near;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;

            //e.Graphics.DrawString("Strana " + strana, dataGridView1.Font, Brushes.Black, new Rectangle(sredina, 1100, 500, 20));

            //stampanje svojih podataka
            if (i < 20)
            {
                e.Graphics.DrawString("\"ALSAZE 2016 \" DOO SJENICA", svojipodaci_font_bold, Brushes.Black, new Rectangle(300, height, 500, 500), stringFormat);
                height += svojipodaci_font_bold.Height + 2;
                e.Graphics.DrawString("Ul. Prijepoljska 100, 36310 Sjenica \nPib 109439308, matični broj 21182532 \nŽiro računi 355-3200492942-51 Vojvodjanska banka\nTelefon 063/619/285", obicni, Brushes.Black, new Rectangle(300, height, 500, 650), stringFormat);
                height += 3 * obicni.Height + 30;
                Point p1 = new Point(0, height);
                Point p2 = new Point(900, height);
                e.Graphics.DrawLine(p, p1, p2);
                height += 5;
                e.Graphics.DrawString("Matični broj: ", za_ukupno, Brushes.Black, new Rectangle(sredina, height, 900, 600));
                e.Graphics.DrawString(textBox4.Text, tekst, Brushes.Black, new Rectangle(sredina + 100, height, 900, 600));
                height += za_ukupno.Height + 10;

                using (GraphicsPath path = RoundedRect(new Rectangle(sredina - 5, height, 355, 170), 20))
                {
                    e.Graphics.DrawPath(pe, path);
                }

                e.Graphics.DrawString("Objekat:", obicni_bold, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString("TEHNIČKI PREGLED", za_ukupno, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));
                height += obicni_bold.Height;
                e.Graphics.DrawString("Mesto promenta:", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString("Sjenica", tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));

                e.Graphics.DrawString("Kupac", obicni_bold, Brushes.Black, new Rectangle(sredina, height, 330, 600));
                e.Graphics.DrawString(textBox1.Text, otpremnica_font, Brushes.Black, new Rectangle(sredina + 100, height, 240, 600));

                height += obicni_bold.Height;
                e.Graphics.DrawString("Datum izdavanja računa:", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString(dateTimePicker1.Value.ToString("dd.MM.yyyy"), tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));
                height += obicni_bold.Height;
                e.Graphics.DrawString("Datum prometa:", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
                e.Graphics.DrawString(dateTimePicker1.Value.ToString("dd.MM.yyyy"), tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));

                e.Graphics.DrawString("Adresa", za_ukupno, Brushes.Black, new Rectangle(sredina, height, 330, 600));
                e.Graphics.DrawString(textBox2.Text + ", " + textBox3.Text, tekst, Brushes.Black, new Rectangle(sredina + 100, height, 240, 600));

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
                e.Graphics.DrawString("", tekst, Brushes.Black, new Rectangle(pocetak + 170, height, 900, 600));

                e.Graphics.DrawString("T.R.", za_ukupno, Brushes.Black, new Rectangle(sredina, height, 900, 600));
                e.Graphics.DrawString(textBox7.Text, tekst, Brushes.Black, new Rectangle(sredina + 100, height, 900, 600));

                height += obicni_bold.Height;

                height += 2 * svojipodaci_font.Height;
                e.Graphics.DrawString("MP Račun-Otpremnica  " + txtOtp.Text, otpremnica_font_opdvuceno, Brushes.Black, new Rectangle(pocetak, height, 900, 600));
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
            e.Graphics.DrawString("PROIZVOD / USLUGA", tabela_font, Brushes.Black, new Rectangle(pocetak, height + 1, 430, dataGridView1.Rows[0].Height), stringFormat);
            #endregion

            #region data2
            e.Graphics.DrawString("RABAT", tabela_font, Brushes.Black, new Rectangle(pocetak + 430, height + 1, 110, dataGridView1.Rows[0].Height), stringFormat);
            #endregion

            #region data3
            e.Graphics.DrawString("PDV", tabela_font, Brushes.Black, new Rectangle(pocetak + 540, height + 1, 110, dataGridView1.Rows[0].Height), stringFormat);
            #endregion

            #region data4
            e.Graphics.DrawString("Cena sa PDV-om", tabela_font, Brushes.Black, new Rectangle(pocetak + 650, height + 8, 60, 2 * dataGridView1.Rows[0].Height), stringFormat);
            #endregion

            #region data5
            e.Graphics.DrawString("Vrednost sa PDV-om", tabela_font, Brushes.Black, new Rectangle(pocetak + 710, height + 8, 70, 2 * dataGridView1.Rows[0].Height), stringFormat);
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
            e.Graphics.DrawString("RB", tabela_font, Brushes.Black, new Rectangle(pocetak, height + 1, 20, dataGridView1.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("Šifra", tabela_font, Brushes.Black, new Rectangle(pocetak + 20, height + 1, 50, dataGridView1.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("Naziv", tabela_font, Brushes.Black, new Rectangle(pocetak + 70, height + 1, 150, dataGridView1.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("JM", tabela_font, Brushes.Black, new Rectangle(pocetak + 220, height + 1, 30, dataGridView1.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("Kol.", tabela_font, Brushes.Black, new Rectangle(pocetak + 250, height + 1, 40, dataGridView1.Rows[0].Height), stringFormatdesno);
            #endregion

            #region data1
            e.Graphics.DrawString("Cena", tabela_font, Brushes.Black, new Rectangle(pocetak + 290, height + 1, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
            #endregion

            #region data1
            e.Graphics.DrawString("Iznos", tabela_font, Brushes.Black, new Rectangle(pocetak + 360, height + 1, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
            #endregion

            #region data1
            e.Graphics.DrawString("%", tabela_font, Brushes.Black, new Rectangle(pocetak + 430, height + 1, 40, dataGridView1.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("Iznos", tabela_font, Brushes.Black, new Rectangle(pocetak + 470, height + 1, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
            #endregion

            #region data1
            e.Graphics.DrawString("%", tabela_font, Brushes.Black, new Rectangle(pocetak + 540, height + 1, 40, dataGridView1.Rows[0].Height), stringFormat);
            #endregion

            #region data1
            e.Graphics.DrawString("Porez", tabela_font, Brushes.Black, new Rectangle(pocetak + 580, height + 1, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
            #endregion

            height += dataGridView1.Rows[0].Height;
            Point prva = new Point(pocetak, height);
            Point zadnja = new Point(790, height);
            e.Graphics.DrawLine(pe, prva, zadnja);

            // Pocetna visina na kojoj se nalaze podaci iz tabele
            height += 2;


            while (i < dataGridView1.Rows.Count - 1)
            {


                e.Graphics.DrawString((i + 1).ToString(), tabela_font, Brushes.Black, new Rectangle(pocetak, height, 20, dataGridView1.Rows[0].Height), stringFormatdesno);
                e.Graphics.DrawString(dataGridView1.Rows[i].Cells[0].FormattedValue.ToString(), tabela_font, Brushes.Black, new Rectangle(pocetak + 22, height, 50, dataGridView1.Rows[0].Height));
                string txt = dataGridView1.Rows[i].Cells[1].FormattedValue.ToString();
                if (txt.Length > 20)
                    e.Graphics.DrawString(txt.Substring(0, 20), tabela_font, Brushes.Black, new Rectangle(pocetak + 70, height, 150, dataGridView1.Rows[0].Height));
                else
                    e.Graphics.DrawString(txt, tabela_font, Brushes.Black, new Rectangle(pocetak + 70, height, 150, dataGridView1.Rows[0].Height));
                e.Graphics.DrawString(dataGridView1.Rows[i].Cells[2].FormattedValue.ToString(), tabela_font, Brushes.Black, new Rectangle(pocetak + 220, height, 30, dataGridView1.Rows[0].Height));

                e.Graphics.DrawString((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 250, height, 40, dataGridView1.Rows[0].Height), stringFormatdesno);

                e.Graphics.DrawString((Convert.ToDouble(dataGridView1.Rows[i].Cells[4].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 290, height, 70, dataGridView1.Rows[0].Height), stringFormatdesno);

                e.Graphics.DrawString((Convert.ToDouble(dataGridView1.Rows[i].Cells[5].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 360, height, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
                ukupno_bez_rabata += Convert.ToDouble(dataGridView1.Rows[i].Cells[5].FormattedValue);
                e.Graphics.DrawString((Convert.ToDouble(dataGridView1.Rows[i].Cells[6].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 430, height, 40, dataGridView1.Rows[0].Height), stringFormatdesno);

                e.Graphics.DrawString((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 470, height, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
                ukupan_rabat += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].FormattedValue);
                e.Graphics.DrawString((Convert.ToInt32(dataGridView1.Rows[i].Cells[8].FormattedValue)).ToString(), tabela_font, Brushes.Black, new Rectangle(pocetak + 540, height, 40, dataGridView1.Rows[0].Height), stringFormatdesno);

                e.Graphics.DrawString((Convert.ToDouble(dataGridView1.Rows[i].Cells[9].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 580, height, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
                ukupan_pdv += Convert.ToDouble(dataGridView1.Rows[i].Cells[9].FormattedValue);
                e.Graphics.DrawString((Convert.ToDouble(dataGridView1.Rows[i].Cells[10].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 650, height, 60, dataGridView1.Rows[0].Height), stringFormatdesno);

                e.Graphics.DrawString((Convert.ToDouble(dataGridView1.Rows[i].Cells[11].FormattedValue)).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 710, height, 70, dataGridView1.Rows[0].Height), stringFormatdesno);

                height += tabela_font.Height + 1;
                prva = new Point(pocetak, height);
                zadnja = new Point(790, height);
                e.Graphics.DrawLine(pe, prva, zadnja);
                if (i == dataGridView1.Rows.Count - 2)
                    e.Graphics.DrawLine(p, prva, zadnja);
                height += 1;
                //e.Graphics.DrawString((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].FormattedValue)).ToString("N", new CultureInfo("is-IS")), dataGridView1.Font, Brushes.Black, new Rectangle(600, height, 100, dataGridView1.Rows[0].Height), stringFormatdesno);
                //potrazuje += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].FormattedValue);
                //duguje += Convert.ToDouble(dataGridView1.Rows[i].Cells[6].FormattedValue);
                //e.Graphics.DrawString((Convert.ToDouble(dataGridView1.Rows[i].Cells[8].FormattedValue)).ToString("N", new CultureInfo("is-IS")), dataGridView1.Font, Brushes.Black, new Rectangle(700, height, 100, dataGridView1.Rows[0].Height), stringFormatdesno);
                //saldo = Convert.ToDouble(dataGridView1.Rows[i].Cells[8].FormattedValue);
                //if (i == 40 || (i > 20 && i+40 % 60 == 0) || (i == dataGridView1.Rows.Count - 2 && dataGridView1.Rows.Count - 1 > 40))
                if (dataGridView1.Rows.Count - 1 > 25 && i == dataGridView1.Rows.Count - 2)
                {
                    i++;

                    e.HasMorePages = true;
                    return;
                }
                i++;
            }

            height += dataGridView1.Rows[0].Height;

            e.Graphics.DrawString("Ukupno bez PDV-a", tekst, Brushes.Black, new Rectangle(pocetak + 470, height, 150, dataGridView1.Rows[0].Height));
            e.Graphics.DrawString(ukupno_bez_rabata.ToString("N", new CultureInfo("is-IS")), tekst, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView1.Rows[0].Height), stringFormatdesno);
            height += tekst.Height;
            e.Graphics.DrawString("PDV ukupno", tekst, Brushes.Black, new Rectangle(pocetak + 470, height, 150, dataGridView1.Rows[0].Height));
            e.Graphics.DrawString(ukupan_pdv.ToString("N", new CultureInfo("is-IS")), tekst, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView1.Rows[0].Height), stringFormatdesno);

            e.Graphics.DrawString("IZNOSI SU ISKAZANI U VALUTI PLAĆANJA - Dinar", tekst, Brushes.Black, new Rectangle(pocetak, height - 5, 500, dataGridView1.Rows[0].Height));

            height += tekst.Height;
            e.Graphics.DrawString("Rabat", tekst, Brushes.Black, new Rectangle(pocetak + 470, height, 150, dataGridView1.Rows[0].Height));
            e.Graphics.DrawString((ukupno_bez_rabata - ukupan_rabat).ToString("N", new CultureInfo("is-IS")), tekst, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView1.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString("NAPOMENA: Nema poreskog oslobađanja", za_ukupno, Brushes.Black, new Rectangle(pocetak, height, 500, dataGridView1.Rows[0].Height));

            height += tekst.Height;
            prva = new Point(pocetak + 470, height);
            zadnja = new Point(790, height);
            e.Graphics.DrawLine(pe, prva, zadnja);
            ukupno += ukupno_bez_rabata - ukupan_rabat + ukupan_pdv;
            height += 2;
            e.Graphics.DrawString("Svega za plaćanje", za_ukupno, Brushes.Black, new Rectangle(pocetak + 470, height, 150, dataGridView1.Rows[0].Height));
            e.Graphics.DrawString(ukupan_rabat.ToString("N", new CultureInfo("is-IS")), za_ukupno, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView1.Rows[0].Height), stringFormatdesno);
            height += za_ukupno.Height;
            e.Graphics.DrawString("Prilikom uplate pozvati se na broj ", tekst, Brushes.Black, new Rectangle(pocetak, height, 500, dataGridView1.Rows[0].Height));
            height += za_ukupno.Height;

            e.Graphics.DrawString("Tarifa", tabela_font, Brushes.Black, new Rectangle(pocetak + 290, height, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString("Osnovica", tabela_font, Brushes.Black, new Rectangle(pocetak + 450, height, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString("PDV", tabela_font, Brushes.Black, new Rectangle(pocetak + 600, height, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString("Iznos", tabela_font, Brushes.Black, new Rectangle(pocetak + 710, height, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
            height += tabela_font.Height;
            prva = new Point(pocetak + 290, height);
            zadnja = new Point(790, height);
            e.Graphics.DrawLine(pe, prva, zadnja);
            height += 2;
            e.Graphics.DrawString("1", tabela_font, Brushes.Black, new Rectangle(pocetak + 290, height, 70, dataGridView1.Rows[0].Height));
            e.Graphics.DrawString("0,00", tabela_font, Brushes.Black, new Rectangle(pocetak + 290, height, 70, dataGridView1.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString((ukupan_rabat).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 390, height, 130, dataGridView1.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString(ukupan_pdv.ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 540, height, 130, dataGridView1.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString(ukupan_rabat.ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView1.Rows[0].Height), stringFormatdesno);
            height += tabela_font.Height;

            prva = new Point(pocetak + 290, height);
            zadnja = new Point(790, height);
            e.Graphics.DrawLine(pe, prva, zadnja);
            height += 2;

            e.Graphics.DrawString((ukupan_rabat).ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 390, height, 130, dataGridView1.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString(ukupan_pdv.ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 540, height, 130, dataGridView1.Rows[0].Height), stringFormatdesno);
            e.Graphics.DrawString(ukupan_rabat.ToString("N", new CultureInfo("is-IS")), tabela_font, Brushes.Black, new Rectangle(pocetak + 650, height, 130, dataGridView1.Rows[0].Height), stringFormatdesno);
            height += tabela_font.Height + 5;

            e.Graphics.DrawString("Za plaćanje posle gore navedenog datuma zaračunava se zakonska kamata", tekst, Brushes.Black, new Rectangle(pocetak, height - 5, 500, dataGridView1.Rows[0].Height));
            height += 50;
            e.Graphics.DrawString("Fakturisao", dataGridView1.Font, Brushes.Black, new Rectangle(20, height, 230, 50), stringFormat);
            e.Graphics.DrawString("Ovlašćeno lice", dataGridView1.Font, Brushes.Black, new Rectangle(270, height, 230, 50), stringFormat);
            e.Graphics.DrawString("Račun preuzeo", dataGridView1.Font, Brushes.Black, new Rectangle(550, height, 250, 50), stringFormat);

            height += 40;
            e.Graphics.DrawString("________________________", dataGridView1.Font, Brushes.Black, new Rectangle(20, height, 200, 50), stringFormat);
            e.Graphics.DrawString("M.P.", dataGridView1.Font, Brushes.Black, new Rectangle(220, height, 50, 50), stringFormat);
            e.Graphics.DrawString("________________________", dataGridView1.Font, Brushes.Black, new Rectangle(270, height, 230, 50), stringFormat);
            e.Graphics.DrawString("M.P.", dataGridView1.Font, Brushes.Black, new Rectangle(500, height, 50, 50), stringFormat);
            e.Graphics.DrawString("________________________", dataGridView1.Font, Brushes.Black, new Rectangle(550, height, 250, 50), stringFormat);
            height += dataGridView1.Font.Height;
            e.Graphics.DrawString("Elvira Suljević", dataGridView1.Font, Brushes.Black, new Rectangle(270, height, 230, 50), stringFormat);
            height += dataGridView1.Font.Height;
            e.Graphics.DrawString("Ovaj dokument je urađen na računaru i punovažan je bez pečata i potpisa", tabela_font, Brushes.Black, new Rectangle(150, height, 380, 50), stringFormat);
            e.Graphics.DrawString("Broj identifikacionog dokumenta", dataGridView1.Font, Brushes.Black, new Rectangle(550, height, 250, 50), stringFormat);
            height += 40;
            e.Graphics.DrawString("________________________", dataGridView1.Font, Brushes.Black, new Rectangle(550, height, 250, 50), stringFormat);

            i = 0;
            strana = 0;
            ukupno = 0;
            ukupno_bez_rabata = 0;
            ukupan_rabat = 0;
            ukupan_pdv = 0;
        }

        private void txtSfr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //Kom.Focus();
                //artikal a = b.DodajArtikal(Convert.ToInt32(txtSfr.Text));
                //txtNaziv.Text = a.opis;
            }
        }

        private void Kom_KeyDown(object sender, KeyEventArgs e)
        {
            this.btnAddItem_Click(sender, e);

        }

        private void txtBar_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //Kom.Focus();
                //artikal a = b.DodajArtikal(txtBar.Text);
                //txtNaziv.Text = a.opis;
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

        private void txtSfr_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnAddItem_Click(object sender, EventArgs e)
        {
            i++;
            double rabat = Convert.ToDouble(txtRabat.Text);
            double cena = double.Parse(txtSfr.Text);

            kolicina = Convert.ToDouble(Kom.Value);
            double iznos_rabat =  cena * kolicina - ((cena * rabat) / 100) * kolicina;
            double pdv = 0;
            double sa_pdvom = cena;
            double iznos = cena * kolicina;
            dt.Rows.Add(i, txtNaziv.Text, "kom", kolicina, cena, iznos, rabat, iznos_rabat, "20", pdv, iznos_rabat / kolicina, iznos_rabat);
            dataGridView1.DataSource = dt;
            double u = 0;
            int k = 0;
            while (k < dataGridView1.Rows.Count - 1)
            {
                u += Convert.ToDouble(dataGridView1.Rows[k].Cells[11].FormattedValue);

                k++;
            }
            textBox5.Text = u.ToString();
            Kom.Value = 1;
            txtRabat.Text = "0";
            
        }

        private void btnRemoveItem_Click(object sender, EventArgs e)
        {


            //try
            //{

            //    dt.Rows.RemoveAt(dataGridView1.CurrentCell.RowIndex);
            //    dataGridView1.DataSource = dt;

            //    int rowIndex = dataGridView1.CurrentCell.RowIndex;
            //    DataGridViewRow row = dataGridView1.Rows[rowIndex];

            //}
            //catch (Exception)
            //{

            //    return;
            //}
            //double u = 0;
            //int k = 0;
            //while (k < dataGridView1.Rows.Count - 1)
            //{
            //    u += Convert.ToDouble(dataGridView1.Rows[k].Cells[8].FormattedValue);

            //    k++;
            //}
            //textBox5.Text = u.ToString();
        }

        public ucOtpremnica()
        {
            InitializeComponent();

            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd.MM.yyyy";
            txtNaziv.Text = "Tehnički pregled vozila";
            txtSfr.Text = "3000";
            txtRabat.Text = "0";
            dt.Columns.Add("Šifra", typeof(int));
            dt.Columns.Add("Opis artikla", typeof(string));
            dt.Columns.Add("JM", typeof(string));
            dt.Columns.Add("Količina", typeof(double));
            dt.Columns.Add("Cena", typeof(double));
            dt.Columns.Add("Iznos", typeof(double));
            dt.Columns.Add("Popust(%)", typeof(double));
            dt.Columns.Add("Popust", typeof(double));
            dt.Columns.Add("Pdv(%)", typeof(string));
            dt.Columns.Add("PDV", typeof(double));
            dt.Columns.Add("Cena sa PDV", typeof(double));
            dt.Columns.Add("Iznos sa PDV", typeof(double));
            
            
        }

        private void ucOtpremnica_Load(object sender, EventArgs e)
        {
            //dateTimePicker1.Format = DateTimePickerFormat.Custom;
            //dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            //dateTimePicker2.Format = DateTimePickerFormat.Custom;
            //dateTimePicker2.CustomFormat = "dd-MM-yyyy";
            //txtRabat.Text = "0";
            //dt.Columns.Add("Opis artikla", typeof(string));
            //dt.Columns.Add("JM", typeof(string));
            //dt.Columns.Add("Količina", typeof(int));
            //dt.Columns.Add("Cena", typeof(double));
            //dt.Columns.Add("Popust(%)", typeof(double));
            //dt.Columns.Add("Iznos", typeof(double));
            //dt.Columns.Add("Pdv(%)", typeof(string));
            //dt.Columns.Add("PDV", typeof(double));
            //dt.Columns.Add("Iznos sa PDV", typeof(double));


        }
    }
}
