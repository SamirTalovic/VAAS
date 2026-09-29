using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Globalization;
using System.Data.OleDb;

namespace TehnickiPregled
{
    public partial class ucPrenos : UserControl
    {
        public ucPrenos()
        {
            InitializeComponent();
        }

        private void ucPrenos_Load(object sender, EventArgs e)
        {
            //richTextBox1.Text = "STR A. NALA\nSjenica, A. Abdagića";
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();

            PrintDocument printDocument = new PrintDocument();

            PaperSize ps = new PaperSize();
            ps.RawKind = (int)PaperKind.A4;
            printDocument.DefaultPageSettings.PaperSize = ps;
            //printDocument.DefaultPageSettings.Landscape = true;
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
            int height = 32;

            Pen p = new Pen(Brushes.Black, 1.0f);
            Pen pe = new Pen(Brushes.Black, 0.5f);
            Font svojipodaci_font = new Font("Arial", 12);
            Font naslov_font = new Font("Arial", 15, FontStyle.Bold);
            Font tekst = new Font("Arial", 10);
            Font ob = new Font("Arial", 5, FontStyle.Bold);
            Font obicni = new Font("Arial", 5);
            Font za_ukupno = new Font("Arial", 10, FontStyle.Bold);
            Font za_licnu_kartu = new Font("Arial", 8);
            Font otpremnica_font = new Font("Arial", 14, FontStyle.Bold);
            Font naslov_font_nebold = new Font("Arial", 15);



            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;

            StringFormat stringFormatdesno = new StringFormat();
            stringFormatdesno.Alignment = StringAlignment.Far;
            stringFormatdesno.LineAlignment = StringAlignment.Near;
            e.Graphics.DrawString("НАЛОГ ЗА ПРЕНОС", naslov_font, Brushes.Black, new Rectangle(407 - 19, height, 400, 50), stringFormatdesno);
            height = 62;
            //stampanje svojih podataka
            e.Graphics.DrawString("платилац", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            e.Graphics.DrawString(richTextBox1.Text, naslov_font, Brushes.Black, new Rectangle(26, height, 354, 59), stringFormat);
            height = 142;
            e.Graphics.DrawString("сврха плаћања", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            e.Graphics.DrawString(richTextBox2.Text, naslov_font_nebold, Brushes.Black, new Rectangle(26, height, 354, 59), stringFormat);
            height = 223;
            e.Graphics.DrawString("прималац", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            e.Graphics.DrawString(richTextBox3.Text, naslov_font, Brushes.Black, new Rectangle(26, height, 354, 59), stringFormat);

            Point po = new Point(397, 62);
            Point po2 = new Point(397, 223 + 59);
            e.Graphics.DrawLine(pe, po, po2);

            //stampanje svojih podataka
            e.Graphics.DrawString("шифра плаћања", za_licnu_kartu, Brushes.Black, new Rectangle(436, 51, 52, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(435, 77, 59, 22));
            e.Graphics.DrawString(textBox1.Text, tekst, Brushes.Black, new Rectangle(436, 80, 51, 50));

            e.Graphics.DrawString("\nвалута", za_licnu_kartu, Brushes.Black, new Rectangle(503, 51, 52, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(502, 77, 59, 22));
            e.Graphics.DrawString(textBox2.Text, tekst, Brushes.Black, new Rectangle(505, 80, 51, 50));

            e.Graphics.DrawString("\nизнос", za_licnu_kartu, Brushes.Black, new Rectangle(583, 51, 52, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(582, 77, 192 + 26 - 19, 22));
            try
            {
                double iznos = double.Parse(textBox3.Text);
                e.Graphics.DrawString(iznos.ToString("N", new CultureInfo("is-IS")), za_ukupno, Brushes.Black, new Rectangle(585, 80, 196, 50));
            }
            catch
            {
                e.Graphics.DrawString(textBox3.Text, za_ukupno, Brushes.Black, new Rectangle(585, 80, 196, 50));
            }

            e.Graphics.DrawString("\nрачун платиоца", za_licnu_kartu, Brushes.Black, new Rectangle(436, 100, 520, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(435, 126, 339 + 26 - 19, 22));
            e.Graphics.DrawString(textBox4.Text, naslov_font, Brushes.Black, new Rectangle(435, 128, 339 + 26 - 19, 20), stringFormat);
            
            e.Graphics.DrawString("\nмодел и позив на број (задужење)", za_licnu_kartu, Brushes.Black, new Rectangle(436, 174 - 26, 500, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(435, 174, 59, 22));
            e.Graphics.DrawString(textBox5.Text, tekst, Brushes.Black, new Rectangle(438, 176, 39, 50));

            e.Graphics.DrawRectangle(pe, new Rectangle(502, 174, 272 + 26 - 19, 22));
            e.Graphics.DrawString(textBox6.Text, tekst, Brushes.Black, new Rectangle(505, 176, 275, 50));

            e.Graphics.DrawString("\nрачун примаоца", za_licnu_kartu, Brushes.Black, new Rectangle(436, 219 - 26, 339, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(435, 219, 339 + 26 - 19, 22));
            e.Graphics.DrawString(textBox7.Text, naslov_font, Brushes.Black, new Rectangle(435, 221, 339+26-19, 20),stringFormat);

            e.Graphics.DrawString("\nмодел и позив на број (одобрење)", za_licnu_kartu, Brushes.Black, new Rectangle(436, 272 - 26, 500, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(435, 272, 59, 22));
            e.Graphics.DrawString(textBox8.Text, tekst, Brushes.Black, new Rectangle(438, 274, 39, 50));

            e.Graphics.DrawRectangle(pe, new Rectangle(502, 272, 272 + 26 - 19, 22));
            e.Graphics.DrawString(textBox9.Text, tekst, Brushes.Black, new Rectangle(505, 274, 275, 50));


            height = 306;
            Point pr = new Point(26, 306);
            Point pr2 = new Point(234, 306);
            e.Graphics.DrawLine(pe, pr, pr2);
            e.Graphics.DrawString("печат и потпис платиоца/примаоца", za_licnu_kartu, Brushes.Black, new Rectangle(26, height + 1, 500, 50));

            height = 340;
            pr = new Point(185, 340);
            pr2 = new Point(380, 340);
            e.Graphics.DrawLine(pe, pr, pr2);
            e.Graphics.DrawString("место и датум пријема", za_licnu_kartu, Brushes.Black, new Rectangle(185, height + 1, 500, 50));
            e.Graphics.DrawString("Sjenica   " + DateTime.Now.Date.ToString("dd.MM.yyyy"), tekst, Brushes.Black, new Rectangle(185, height - 15, 500, 50));


            pr = new Point(435, 340);
            pr2 = new Point(570, 340);
            e.Graphics.DrawLine(pe, pr, pr2);
            e.Graphics.DrawString("датум извршења", za_licnu_kartu, Brushes.Black, new Rectangle(435, height + 1, 500, 50));

            height = 340;
            e.Graphics.DrawString("НАЛОГ ЗА ПРЕНОС", ob, Brushes.Black, new Rectangle(30, height, 500, 50));
            e.Graphics.DrawString(" - obr. 3", obicni, Brushes.Black, new Rectangle(96, height, 500, 50));
            height += obicni.Height;
            e.Graphics.DrawString("Образац бр. 3", za_licnu_kartu, Brushes.Black, new Rectangle(362, 360, 500, 50));
            e.Graphics.DrawRectangle(pe, new Rectangle(672, 325, 19, 18));
            e.Graphics.DrawString("ХИТНО", obicni, Brushes.Black, new Rectangle(669, 325 + 18 + 1, 500, 50));

            po = new Point(0, 376);
            po2 = new Point(793, 376);
            e.Graphics.DrawLine(pe, po, po2);
            height = 32 + 374;
            //e.Graphics.DrawString("НАЛОГ ЗА ПРЕНОС", naslov_font, Brushes.Black, new Rectangle(407 - 19, height, 400, 50), stringFormatdesno);
            //height = 62 + 374;
            ////stampanje svojih podataka
            //e.Graphics.DrawString("платилац", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            //e.Graphics.DrawString(richTextBox1.Text, naslov_font, Brushes.Black, new Rectangle(26, height, 354, 59), stringFormat);
            //height = 142 + 374;
            //e.Graphics.DrawString("сврха плаћања", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            //e.Graphics.DrawString(richTextBox2.Text, naslov_font_nebold, Brushes.Black, new Rectangle(26, height, 354, 59), stringFormat);
            //height = 223 + 374;
            //e.Graphics.DrawString("прималац", za_licnu_kartu, Brushes.Black, new Rectangle(40, height - 12, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(26, height, 354, 59));
            //e.Graphics.DrawString(richTextBox3.Text, naslov_font, Brushes.Black, new Rectangle(26, height , 354, 59), stringFormat);


            //po = new Point(397, 62 + 374);
            //po2 = new Point(397, 223 + 59 + 374);
            //e.Graphics.DrawLine(pe, po, po2);

            ////stampanje svojih podataka
            //e.Graphics.DrawString("шифра плаћања", za_licnu_kartu, Brushes.Black, new Rectangle(436, 51 + 374, 52, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 77 + 374, 59, 22));
            //e.Graphics.DrawString(textBox1.Text, tekst, Brushes.Black, new Rectangle(436, 80 + 374, 51, 50));

            //e.Graphics.DrawString("\nвалута", za_licnu_kartu, Brushes.Black, new Rectangle(503, 51 + 374, 52, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(502, 77 + 374, 59, 22));
            //e.Graphics.DrawString(textBox2.Text, tekst, Brushes.Black, new Rectangle(505, 80 + 374, 51, 50));


            //e.Graphics.DrawString("\nизнос", za_licnu_kartu, Brushes.Black, new Rectangle(583, 51 + 374, 52, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(582, 77 + 374, 192 + 26 - 19, 22));
            //try
            //{
            //    double iznos = double.Parse(textBox3.Text);
            //    e.Graphics.DrawString(iznos.ToString("N", new CultureInfo("is-IS")), za_ukupno, Brushes.Black, new Rectangle(585, 80 + 374, 196, 50));
            //}
            //catch
            //{
            //    e.Graphics.DrawString(textBox3.Text, za_ukupno, Brushes.Black, new Rectangle(585, 80 + 374, 196, 50));
            //}


            //e.Graphics.DrawString("\nрачун платиоца", za_licnu_kartu, Brushes.Black, new Rectangle(436, 100 + 374, 520, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 126 + 374, 339 + 26 - 19, 22));
            //e.Graphics.DrawString(textBox4.Text, naslov_font, Brushes.Black, new Rectangle(435, 128 + 374, 339 + 26 - 19, 20), stringFormat);

            //e.Graphics.DrawString("\nмодел и позив на број (задужење)", za_licnu_kartu, Brushes.Black, new Rectangle(436, 174 - 26 + 374, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 174 + 374, 59, 22));
            //e.Graphics.DrawString(textBox5.Text, tekst, Brushes.Black, new Rectangle(438, 176 + 374, 39, 50));

            //e.Graphics.DrawRectangle(pe, new Rectangle(502, 174 + 374, 272 + 26 - 19, 22));
            //e.Graphics.DrawString(textBox6.Text, tekst, Brushes.Black, new Rectangle(505, 176 + 374, 275, 50));

            //e.Graphics.DrawString("\nрачун примаоца", za_licnu_kartu, Brushes.Black, new Rectangle(436, 219 + 374 - 26, 339, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 219 + 374, 339 + 26 - 19, 22));
            //e.Graphics.DrawString(textBox7.Text, naslov_font, Brushes.Black, new Rectangle(435, 221 + 374, 339+26-19, 20), stringFormat);

            //e.Graphics.DrawString("\nмодел и позив на број (одобрење)", za_licnu_kartu, Brushes.Black, new Rectangle(436, 272 - 26 + 374, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(435, 272 + 374, 59, 22));
            //e.Graphics.DrawString(textBox8.Text, tekst, Brushes.Black, new Rectangle(438, 274 + 374, 39, 50));

            //e.Graphics.DrawRectangle(pe, new Rectangle(502, 272 + 374, 272 + 26 - 19, 22));
            //e.Graphics.DrawString(textBox9.Text, tekst, Brushes.Black, new Rectangle(505, 274 + 374, 275, 50));


            //height = 306 + 374;
            //pr = new Point(26, 306 + 374);
            //pr2 = new Point(234, 306 + 374);
            //e.Graphics.DrawLine(pe, pr, pr2);
            //e.Graphics.DrawString("печат и потпис платиоца/примаоца", za_licnu_kartu, Brushes.Black, new Rectangle(26, height + 1, 500, 50));

            //height = 340 + 374;
            //pr = new Point(185, 340 + 374);
            //pr2 = new Point(380, 340 + 374);
            //e.Graphics.DrawLine(pe, pr, pr2);
            //e.Graphics.DrawString("место и датум пријема", za_licnu_kartu, Brushes.Black, new Rectangle(185, height + 1, 500, 50));
            //e.Graphics.DrawString("Sjenica   " + DateTime.Now.Date.ToString("dd.MM.yyyy"), tekst, Brushes.Black, new Rectangle(185, height - 15, 500, 50));

            //pr = new Point(435, 340 + 374);
            //pr2 = new Point(570, 340 + 374);
            //e.Graphics.DrawLine(pe, pr, pr2);
            //e.Graphics.DrawString("датум извршења", za_licnu_kartu, Brushes.Black, new Rectangle(435, height + 1, 500, 50));

            //height = 340 + 374;
            //e.Graphics.DrawString("NALOG ZA ПРЕНОС", ob, Brushes.Black, new Rectangle(30, height, 500, 50));
            //e.Graphics.DrawString(" - obr. 3", obicni, Brushes.Black, new Rectangle(96, height, 500, 50));
            //height += obicni.Height;
            //e.Graphics.DrawString("Образац бр. 3", za_licnu_kartu, Brushes.Black, new Rectangle(362, 360 + 374, 500, 50));
            //e.Graphics.DrawRectangle(pe, new Rectangle(672, 325 + 374, 19, 18));
            //e.Graphics.DrawString("ХИТНО", obicni, Brushes.Black, new Rectangle(669, 325 + 374 + 18 + 1, 500, 50));

            //po = new Point(0, 376 + 374);
            //po2 = new Point(793, 376 + 374);
            //e.Graphics.DrawLine(pe, po, po2);
        }


    }
}
