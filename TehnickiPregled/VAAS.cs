using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TehnickiPregled
{
    public partial class VAAS : Form
    {
        double second;
        public VAAS()
        {
            InitializeComponent();
            Left = Top = 0;
            Width = Screen.PrimaryScreen.WorkingArea.Width;
            Height = Screen.PrimaryScreen.WorkingArea.Height;
            timer1.Start();
            label1.Left = (this.panel2.Width - label1.Width) / 2;

            ucRegistracija p = new ucRegistracija();
            p.Dock = DockStyle.Fill;
            panel2.Controls.Clear();
            panel2.Controls.Add(p);
        }

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd,
                         int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        const int WS_MINIMIZEBOX = 0x20000;
        const int CS_DBLCLKS = 0x8;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.Style |= WS_MINIMIZEBOX;
                cp.ClassStyle |= CS_DBLCLKS;
                return cp;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lblClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lblMinimize_Click(object sender, EventArgs e)
        {

            WindowState = FormWindowState.Minimized;
        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {
            bunifuFlatButton1.selected = true;
            bunifuFlatButton2.selected = false;
            bunifuFlatButton3.selected = false;
            bunifuFlatButton4.selected = false;
            bunifuFlatButton5.selected = false;
            bunifuFlatButton6.selected = false;
            bunifuFlatButton7.selected = false;

            ucRegistracija p = new ucRegistracija();
            p.Dock = DockStyle.Fill;
            panel2.Controls.Clear();
            panel2.Controls.Add(p);
        }

        private void bunifuFlatButton2_Click(object sender, EventArgs e)
        {
            bunifuFlatButton1.selected = false;
            bunifuFlatButton2.selected = true;
            bunifuFlatButton3.selected = false;
            bunifuFlatButton4.selected = false;
            bunifuFlatButton5.selected = false;
            bunifuFlatButton6.selected = false;
            bunifuFlatButton7.selected = false;

            //ucMagacin p = new ucMagacin();
            //p.Dock = DockStyle.Fill;
            //panel2.Controls.Clear();
            //panel2.Controls.Add(p);
        }

        private void bunifuFlatButton3_Click(object sender, EventArgs e)
        {
            bunifuFlatButton1.selected = false;
            bunifuFlatButton2.selected = false;
            bunifuFlatButton3.selected = true;
            bunifuFlatButton4.selected = false;
            bunifuFlatButton5.selected = false;
            bunifuFlatButton6.selected = false;
            bunifuFlatButton7.selected = false;

            ucOtpremnica p = new ucOtpremnica();
            p.Dock = DockStyle.Fill;
            panel2.Controls.Clear();
            panel2.Controls.Add(p);
        }

        private void bunifuFlatButton4_Click(object sender, EventArgs e)
        {
            bunifuFlatButton1.selected = false;
            bunifuFlatButton2.selected = false;
            bunifuFlatButton3.selected = false;
            bunifuFlatButton4.selected = true;
            bunifuFlatButton5.selected = false;
            bunifuFlatButton6.selected = false;
            bunifuFlatButton7.selected = false;

            //ucDnevni p = new ucDnevni();
            //p.Dock = DockStyle.Fill;
            //panel2.Controls.Clear();
            //panel2.Controls.Add(p);
        }

        private void bunifuFlatButton5_Click(object sender, EventArgs e)
        {
            bunifuFlatButton1.selected = false;
            bunifuFlatButton2.selected = false;
            bunifuFlatButton3.selected = false;
            bunifuFlatButton4.selected = false;
            bunifuFlatButton5.selected = true;
            bunifuFlatButton6.selected = false;
            bunifuFlatButton7.selected = false;

            ucUgovori p = new ucUgovori();
            p.Dock = DockStyle.Fill;
            panel2.Controls.Clear();
            panel2.Controls.Add(p);
        }

        private void bunifuFlatButton6_Click(object sender, EventArgs e)
        {
            bunifuFlatButton1.selected = false;
            bunifuFlatButton2.selected = false;
            bunifuFlatButton3.selected = false;
            bunifuFlatButton4.selected = false;
            bunifuFlatButton5.selected = false;
            bunifuFlatButton6.selected = true;
            bunifuFlatButton7.selected = false;

            ucUplate p = new ucUplate();
            p.Dock = DockStyle.Fill;
            panel2.Controls.Clear();
            panel2.Controls.Add(p);
        }

        private void bunifuFlatButton7_Click(object sender, EventArgs e)
        {
            bunifuFlatButton1.selected = false;
            bunifuFlatButton2.selected = false;
            bunifuFlatButton3.selected = false;
            bunifuFlatButton4.selected = false;
            bunifuFlatButton5.selected = false;
            bunifuFlatButton6.selected = false;
            bunifuFlatButton7.selected = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            second += 0.1;

            if (second == 0.1)
            {
                timer1.Stop();
               // bunifuTransition1.ShowSync(pictureBox1);
                //
            }

            //if (second > 3.0)
            //{
            //    timer1.Stop();
            //    bunifuTransition1.HideSync(pictureBox1);
            //    //
            //}
                //if (second == 3.0 )
                //{
                //    MessageBox.Show("Presek");
                //    timer1.Stop();
                //    bunifuTransition1.HideSync(pictureBox1);

                //}
            }

        private void btnMinMax_Click(object sender, EventArgs e)
        {
            panel2.Dock = DockStyle.Fill;
            if (SlideMenu.Width == 294)
            {
                SlideMenu.Visible = false;
                SlideMenu.Width = 65;
                SlideMenu.Dock = DockStyle.Left;
                bunifuTransitionPanel.ShowSync(SlideMenu);
                label1.Left = (this.panel2.Width - label1.Width) / 2;
                lblClose.BringToFront();
            }
            else
            {
                SlideMenu.Visible = false;
                bunifuTransitionPanel.HideSync(SlideMenu);
                SlideMenu.Width = 294;
                bunifuTransitionPanel.ShowSync(SlideMenu);
                label1.Left = (this.panel2.Width - label1.Width) / 2;
                lblClose.Visible = true;
                lblClose.BringToFront();
            }
        }


        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
    }
}
