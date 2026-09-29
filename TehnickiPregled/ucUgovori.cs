using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using System.Drawing.Printing;
using SaobracajnaNET;
using System.IO;
using System.Diagnostics;
using System.Data.OleDb;
using System.Drawing.Drawing2D;
using Word = Microsoft.Office.Interop.Word;
using BashChelik;
using System.Drawing.Imaging;

namespace TehnickiPregled
{
    public partial class ucUgovori : UserControl
    {
        Broker b = new Broker();
        DataTable dt = new DataTable();
        int i = 0;
        public ucUgovori()
        {
            InitializeComponent();
            
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

      
        private void label90_Click(object sender, EventArgs e)
        {

        }

        private void bunifuThinButton29_Click(object sender, EventArgs e)
        {
            var readers = SaobracajnaReader.GetAllReaders();
            if (readers.Count == 0)
            {
                MessageBox.Show("Nema povezanih citaca.");

                return;
            }

            ReaderDescriptor descriptor = readers[0];


            try
            {
                var data = SaobracajnaReader.ReadAll(descriptor);
                PrintData(data);
                string[] podaci = cbFLMesto.Text.Split(',');

                txtFLOps.Text = podaci[1];
                //cbFLMesto.Text = podaci[0];
                txtFLUlica.Text = podaci[2];
                txtFLBroj.Text = podaci[3];
                //richTextBox1.Text = txtFLIme.Text + " " + txtFLPrez.Text + Environment.NewLine + podaci[2] + " " + podaci[3] + " " + podaci[1];

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void PrintData(AllData data)
        {
            PrintStruct(data.VehicleData);
            PrintStruct(data.DocumentData);
            PrintStruct(data.PersonalData);
            MessageBox.Show(data.RegistrationData.Count + " uspešno čitanje podataka sa saobraćajne dozvole");
            
        }

        private void PrintStruct(object instance)
        {
            var type = instance.GetType();
            int i = 0;
            var fields = type.GetFields();
            foreach (var field in fields)
            {
                var value = field.GetValue(instance);
                switch (field.Name)
                {
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
                        cbVBoja.Text = value.ToString();
                        break;
                    case "OwnerName":
                        txtFLIme.Text = value.ToString();
                        break;
                    case "OwnersSurnameOrBusinessName":
                        //if (checkBox4.Checked)
                        //{
                            txtFLPrez.Text = value.ToString();
                        //}
                        //else
                          //  txtPLNaziv.Text = value.ToString();
                        break;
                    case "OwnerAddress":
                        cbFLMesto.Text = value.ToString();
                        break;
                   case "OwnersPersonalNo":
                        //case "ownersPersonalNo":
                        txtFLMatBr.Text = value.ToString();
                        break;
                 default:
                       // MessageBox.Show(type.Name + " " + field.Name + ", " + value);
                        break;

                }
                
                //Console.WriteLine("{0}: {1}", type.Name + "." + field.Name, value);
                i++;
            }
         
        }

   
    
        private void ucRegistracija_Load(object sender, EventArgs e)
        {
            List<string> boje = new List<string>();
            foreach (var b in b.lista_boja())
            {
                boje.Add(b.BojaSifra + " " + b.BojaNaziv);
            }
            cbVBoja.DataSource = boje;
            cbVBoja.Text = "";

            List<string> marke = new List<string>();
            foreach (var b in b.lista_marka())
            {
                marke.Add(b.KaroserijaNaziv);
            }
            cbVMarka.DataSource = marke;
            cbVMarka.Text = "";


            List<string> specijalna = new List<string>();
            foreach (var s in b.lista_spec())
            {
                specijalna.Add(s.Opis);

            }
            cbVSpec.DataSource = specijalna;
            cbVSpec.Text = "";

            List<string> str = new List<string>();
            foreach (var i in b.vrvozila())
            {
                str.Add(i.NazivVrste);
            }
            cbVrstaVoz.DataSource = str;
            cbVrstaVoz.Text = "";

            //cbTPKontr1.DataSource = b.nadjiKotrolore();
            //cbTPKontr2.DataSource = b.nadjiKotrolore();
            
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
        
        int iTotalFields = 0;
        private void bunifuThinButton214_Click_1(object sender, EventArgs e)
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
                        oWord.Selection.TypeText(txtFLPrez.Text + " "+txtFLIme.Text);
                    }
                    if (fieldName == "ULICA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLUlica.Text);
                    }
                    if (fieldName == "BROJ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLBroj.Text);
                    }
                    if (fieldName == "GRAD")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLOps.Text);
                    }
                    if (fieldName == "JMBG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLMatBr.Text);
                    }
                    if (fieldName == "LICNA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtTPBrDokumentaStr.Text);
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

        private void bunifuThinButton216_Click(object sender, EventArgs e)
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
                        oWord.Selection.TypeText(txtFLPrez.Text + " " + txtFLIme.Text);
                    }
                    if (fieldName == "ULICA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLUlica.Text);
                    }
                    if (fieldName == "BROJ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLBroj.Text);
                    }
                    if (fieldName == "GRAD")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLOps.Text);
                    }
                    if (fieldName == "JMBG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLMatBr.Text);
                    }
                    if (fieldName == "LICNA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtTPBrDokumentaStr.Text);
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

        private void bunifuThinButton213_Click_1(object sender, EventArgs e)
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
                        oWord.Selection.TypeText(txtFLPrez.Text + " " + txtFLIme.Text);
                    }
                    if (fieldName == "ULICA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLUlica.Text);
                    }
                    if (fieldName == "BROJ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLBroj.Text);
                    }
                    if (fieldName == "GRAD")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLOps.Text);
                    }
                    if (fieldName == "JMBG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLMatBr.Text);
                    }
                    if (fieldName == "LICNA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtTPBrDokumentaStr.Text);
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

        private void bunifuThinButton215_Click(object sender, EventArgs e)
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
                        oWord.Selection.TypeText(txtFLPrez.Text + " " + txtFLIme.Text);
                    }
                    if (fieldName == "ULICA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLUlica.Text);
                    }
                    if (fieldName == "BROJ")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLBroj.Text);
                    }
                    if (fieldName == "GRAD")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLOps.Text);
                    }
                    if (fieldName == "JMBG")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtFLMatBr.Text);
                    }
                    if (fieldName == "LICNA")
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(txtTPBrDokumentaStr.Text);
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

        private void bunifuThinButton217_Click(object sender, EventArgs e)
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
                txtGRAD.Text = data.VariablePersonalData.Place;
                txtULiBR.Text = data.VariablePersonalData.Street + " " + data.VariablePersonalData.HouseNumber;
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

                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void bunifuThinButton218_Click(object sender, EventArgs e)
        {
            try
            {
                var data = ElectronicIdReader.ReadAll();
                    //txtLicnaKartaK.Text = data.DocumentData.DocumentRegistrationNumber;

                    //    Console.WriteLine("Fiksni podaci ------------------------------------------");
                    
                    txtJMBGK.Text = data.FixedPersonalData.PersonalNumber;
                    txtPIK.Text = data.FixedPersonalData.Surname + " " + data.FixedPersonalData.GivenName;
                    //txtULiBR.Text = d1.placeOfBirth;
                    //txtMestoRodjenja.Text += "," + d1.communityOfBirth;
                    //txtGRAD.Text += ", " + d1.stateOfBirth;


                    //    Console.WriteLine("Variablini podaci ------------------------------------------");
                    
                    //txtPrebivaliste.Text= d2.state;
                    //txtPrebivaliste.Text = d2.community;
                    txtGradK.Text = data.VariablePersonalData.Place;
                    txtULiBK.Text = data.VariablePersonalData.Street + " " + data.VariablePersonalData.HouseNumber;
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
               
                MessageBox.Show(ex.Message.ToString());
            }
        }

    private void bunifuThinButton27_Click(object sender, EventArgs e)
        {

        }

        private void btnAddItem_Click(object sender, EventArgs e)
        {

        }

        private void txtSfr_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void Kom_KeyDown(object sender, KeyEventArgs e)
        {

        }
    }
}