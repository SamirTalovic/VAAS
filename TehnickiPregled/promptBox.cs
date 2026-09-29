using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TehnickiPregled
{
    public static class promptBox
    {
        static Broker b = new Broker();
        public static string ShowDialog(string caption, string text)
        {
            Form prompt = new Form();
            prompt.BackColor = System.Drawing.Color.White;
            prompt.Width = 280;
            prompt.Height = 200;
            prompt.Text = caption;
            Label textLabel = new Label() { Left = 16, Top = 20, Width = 240, Text = text };
            TextBox textBox = new TextBox() { Left = 16, Top = 40, Width = 240, TabIndex = 0, TabStop = true, Font = new System.Drawing.Font("Times New Roman", 15) };
            Button confirmation = new Button() { Text = "Potvrdi", Left = 50, Width = 150, Top = 88, TabIndex = 1, TabStop = true, BackColor = System.Drawing.Color.Blue, ForeColor = System.Drawing.Color.White, Font = new System.Drawing.Font("Times New Roman", 15), Height = 30 };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;
            prompt.StartPosition = FormStartPosition.CenterScreen;
            prompt.ShowDialog();
            return string.Format("{0}", textBox.Text);
        }
    }
}
