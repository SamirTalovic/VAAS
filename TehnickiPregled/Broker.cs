using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace TehnickiPregled
{
    public class Broker
    {
        OleDbConnection connection;
        OleDbCommand command;

        private void ConnectTo()
        {
            connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=baza.accdb;Persist Security Info=False");
            command = connection.CreateCommand();
        }

        public Broker()
        {
            ConnectTo();
        }

        internal List<karoserije> lista_ka(string naziv)
        {
            List<karoserije> lista = new List<karoserije>();
            try
            {
                command.CommandText = "SELECT * FROM karoserije INNER JOIN vrstavozila ON vrstavozila.ID = karoserije.VrstaSifra WHERE vrstavozila.NazivVrste = '" + naziv + "'";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    karoserije k = new karoserije();
                    k.KaroserijaNaziv = reader["KaroserijaNaziv"].ToString();
                    lista.Add(k);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista;
        }
        internal List<karoserije> lista_dec(string naziv)
        {
            List<karoserije> lista = new List<karoserije>();
            try
            {
                command.CommandText = "SELECT * FROM decimalna WHERE OblikKaroserije = '" + naziv + "'";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    karoserije k = new karoserije();
                    k.KaroserijaNaziv = reader["Naziv"].ToString();
                    lista.Add(k);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista;
        }

        internal int nadjiOtpremnicu()
        {
            int s = 0;
            try
            {
                command.CommandText = "SELECT * FROM brojacOtpremnica;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {

                    s = int.Parse(reader["broj"].ToString());
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
                s = 0;
            }
            return s;
        }

        internal string nadjiMarginuLevo()
        {
            string s = "";
            try
            {
                command.CommandText = "SELECT * FROM margine;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {

                    s = reader["levo"].ToString();
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return s;
        }
        internal string nadjiMarginuDesno()
        {
            string s = "";
            try
            {
                command.CommandText = "SELECT * FROM margine;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {

                    s = reader["desno"].ToString();
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return s;
        }
        internal List<string> lista_delova(string v)
        {
            List<string> lista = new List<string>();
            try
            {
                command.CommandText = "SELECT * FROM delovi WHERE Grupa = '" + v + "';";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string s = reader["NazivDela"].ToString();

                    lista.Add(s);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista;
        }

        internal List<string> lista_grupa()
        {
            List<string> lista = new List<string>();
            try
            {
                command.CommandText = "SELECT DISTINCT Grupa FROM delovi;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string s = reader["Grupa"].ToString();

                    lista.Add(s);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista;
        }

        internal List<neispravnost> lista_kvarova(string deo)
        {
            List<neispravnost> lista = new List<neispravnost>();
            try
            {
                command.CommandText = "SELECT * FROM neispravnosti WHERE Deo = '" + deo + "';";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    neispravnost n = new neispravnost();
                    n.ID = reader["ID"].ToString();
                    n.Neispravnost = reader["Neispravnost"].ToString();
                    n.Deo = reader["Deo"].ToString();
                    lista.Add(n);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista;
        }

        internal List<specijalnaVozila> lista_spec()
        {
            List<specijalnaVozila> lista = new List<specijalnaVozila>();
            try
            {
                command.CommandText = "SELECT * FROM specijalnaVozila";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    specijalnaVozila s = new specijalnaVozila();
                    s.Opis = reader["Opis"].ToString();
                    lista.Add(s);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista;

        }
        internal List<vozilo> ucitajTabliceIsticu(string datumOd, string datumDo)
        {
            List<vozilo> voz = new List<vozilo>();
            try
            {
                //command.CommandText = "select * from registracija where DateAdd('yyyy', 1, CDate(Format(Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4),'dd / mm / yyyy'))) >= CDate(Format('" + datumOd + "','dd / mm / yyyy')) and DateAdd('yyyy', 1, CDate(Format(Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4),'dd / mm / yyyy'))) <= CDate(Format('" + datumDo + "','dd / mm / yyyy')) order by CDate(Format(Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4),'dd / mm / yyyy'));";

                command.CommandText = "select * from registracija where DateAdd('yyyy', 0, CDate(Format(Left([DatumIsteka],2) & '/' & MID([DatumIsteka],4,2) & '/' & MID([DatumIsteka],7,4),'dd / mm / yyyy'))) >= CDate(Format('" + datumOd + "','dd / mm / yyyy')) and DateAdd('yyyy', 0, CDate(Format(Left([DatumIsteka],2) & '/' & MID([DatumIsteka],4,2) & '/' & MID([DatumIsteka],7,4),'dd / mm / yyyy'))) <= CDate(Format('" + datumDo + "','dd / mm / yyyy')) order by CDate(Format(Left([DatumIsteka],2) & '/' & MID([DatumIsteka],4,2) & '/' & MID([DatumIsteka],7,4),'dd / mm / yyyy'));";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    vozilo v = new vozilo();
                    v.IDBroj = reader["IDBroj"].ToString();
                    v.RegistarskiBroj = reader["RegistarskiBroj"].ToString();
                    v.VrstaVozila = reader["VrstaVozila"].ToString();
                    v.PodvrstaVozila = reader["PodvrstaVozila"].ToString();
                    v.klasa = reader["klasa"].ToString();
                    v.oblikkaroserije = reader["oblikkaroserije"].ToString();
                    v.OblikKaroserijeSkraceno = reader["oblikkaroserijeSkraceno"].ToString();
                    v.specijalno = reader["specijalno"].ToString();
                    v.Marka = reader["Marka"].ToString();
                    v.Model = reader["Model"].ToString();
                    v.Tip = reader["Tip"].ToString();
                    v.GodinaProizvodnje = reader["GodinaProizvodnje"].ToString();
                    v.BrojSasije = reader["BrojSasije"].ToString();
                    v.BrojMotora = reader["BrojMotora"].ToString();
                    v.Boja = reader["Boja"].ToString();
                    v.DatumPrveRegostracije = reader["DatumPrveRegostracije"].ToString();
                    v.Odnos = reader["Odnos"].ToString();
                    v.Snaga = reader["Snaga"].ToString();
                    v.Zapremina = reader["Zapremina"].ToString();
                    v.NajvecaMasa = reader["NajvecaMasa"].ToString();
                    v.Tezina = reader["Tezina"].ToString();
                    v.Nosivost = reader["Nosivost"].ToString();
                    v.Sedenje = reader["Sedenje"].ToString();
                    v.Stajanje = reader["Stajanje"].ToString();
                    v.BrojOsovina = reader["BrojOsovina"].ToString();
                    v.Gorivo = reader["Gorivo"].ToString();
                    v.Homologacija = reader["Homologacija"].ToString();
                    v.ImeVL = reader["ImeVL"].ToString();
                    v.PrezimeVL = reader["PrezimeVL"].ToString();
                    v.MaticniBrojVL = reader["MaticniBrojVL"].ToString();
                    v.BRLKVL = reader["BRLKVL"].ToString();
                    v.BrojPosteVL = reader["BrojPosteVL"].ToString();
                    v.MestoVL = reader["MestoVL"].ToString();
                    v.UlicaVL = reader["UlicaVL"].ToString();
                    v.BrojVL = reader["BrojVL"].ToString();
                    v.TelefonaVL = reader["TelefonaVL"].ToString();
                    v.ZiroRacunVL = reader["ZiroRacunVL"].ToString();
                    v.ImeDL = reader["ImeDL"].ToString();
                    v.PrezimeDL = reader["PrezimeDL"].ToString();
                    v.MaticniBrojDL = reader["MaticniBrojDL"].ToString();
                    v.BRLKDL = reader["BRLKDL"].ToString();
                    v.BrojPosteDL = reader["BrojPosteDL"].ToString();
                    v.MestoDL = reader["MestoDL"].ToString();
                    v.UlicaDL = reader["UlicaDL"].ToString();
                    v.BrojDL = reader["BrojDL"].ToString();
                    v.TelefonaDL = reader["TelefonaDL"].ToString();
                    v.Napomena = reader["Napomena"].ToString();
                    v.Datum = reader["Datum"].ToString();
                    v.VremeOd = reader["VremeOd"].ToString();
                    v.VremeDo = reader["VremeDo"].ToString();
                    v.Kontrolor1 = reader["Kontrolor1"].ToString();
                    v.Kontrolor2 = reader["Kontrolor2"].ToString();
                    v.Predjenikm = reader["Predjenikm"].ToString();
                    v.PoslednjeUverenjeBroj = reader["PoslednjeUverenjeBroj"].ToString();
                    v.PoslednjeUverenjeDatum = reader["PoslednjeUverenjeDatum"].ToString();
                    v.DecimalnaOznaka = reader["DecimalnaOznaka"].ToString();
                    v.OblikKaroserijeSkraceno = reader["OblikKaroserijeSkraceno"].ToString();
                    v.BojaL = reader["BojaL"].ToString(); v.BojaD = reader["BojaD"].ToString();
                    v.BojaD = reader["BojaD"].ToString();
                    v.Ispravnost = reader["Ispravnost"].ToString();
                    v.BrSati = reader["BrSati"].ToString();
                    v.Kuka = reader["Kuka"].ToString();
                    v.OpterecenjeN1 = reader["OpterecenjeN1"].ToString();
                    v.OpterecenjeN2 = reader["OpterecenjeN2"].ToString();
                    v.OpterecenjeN3 = reader["OpterecenjeN3"].ToString();
                    v.OpterecenjeN4 = reader["OpterecenjeN4"].ToString();
                    v.OpterecenjeN5 = reader["OpterecenjeN5"].ToString();
                    v.PneumatikL1 = reader["PneumatikL1"].ToString();
                    v.PneumatikL2 = reader["PneumatikL2"].ToString();
                    v.PneumatikL3 = reader["PneumatikL3"].ToString();
                    v.PneumatikL4 = reader["PneumatikL4"].ToString();
                    v.PneumatikL5 = reader["PneumatikL5"].ToString();
                    v.PneumatikD1 = reader["PneumatikD1"].ToString();
                    v.PneumatikD2 = reader["PneumatikD2"].ToString();
                    v.PneumatikD3 = reader["PneumatikD3"].ToString();
                    v.PneumatikD4 = reader["PneumatikD4"].ToString();
                    v.PneumatikD5 = reader["PneumatikD5"].ToString();
                    v.zaustavljanje = reader["Zaustavljanje"].ToString();
                    v.terensko = reader["terensko"].ToString();
                    v.komentar = reader["komentar"].ToString();
                    v.rejting = reader["rejting"].ToString();
                    v.decplusoznaka = reader["Decplusoznaka"].ToString();
                    v.datumIsteka = reader["datumIsteka"].ToString();
                    v.Poruka = reader["Poruka"].ToString();
                    voz.Add(v);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return voz;
        }
        internal List<vozilo> ucitajTabliceDolazili(string datumOd, string datumDo)
        {
            List<vozilo> voz = new List<vozilo>();
            try
            {
                //command.CommandText = "select * from registracija where DateAdd('yyyy', 1, CDate(Format(Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4),'dd / mm / yyyy'))) >= CDate(Format('" + datumOd + "','dd / mm / yyyy')) and DateAdd('yyyy', 1, CDate(Format(Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4),'dd / mm / yyyy'))) <= CDate(Format('" + datumDo + "','dd / mm / yyyy')) order by CDate(Format(Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4),'dd / mm / yyyy'));";
                command.CommandText = "select * from registracija where DateAdd('yyyy', 0, CDate(Format(Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4),'dd / mm / yyyy'))) >= CDate(Format('" + datumOd + "','dd / mm / yyyy')) and DateAdd('yyyy', 0, CDate(Format(Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4),'dd / mm / yyyy'))) <= CDate(Format('" + datumDo + "','dd / mm / yyyy')) order by CDate(Format(Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4),'dd / mm / yyyy'));";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    vozilo v = new vozilo();
                    v.IDBroj = reader["IDBroj"].ToString();
                    v.RegistarskiBroj = reader["RegistarskiBroj"].ToString();
                    v.VrstaVozila = reader["VrstaVozila"].ToString();
                    v.PodvrstaVozila = reader["PodvrstaVozila"].ToString();
                    v.klasa = reader["klasa"].ToString();
                    v.oblikkaroserije = reader["oblikkaroserije"].ToString();
                    v.OblikKaroserijeSkraceno = reader["oblikkaroserijeSkraceno"].ToString();
                    v.specijalno = reader["specijalno"].ToString();
                    v.Marka = reader["Marka"].ToString();
                    v.Model = reader["Model"].ToString();
                    v.Tip = reader["Tip"].ToString();
                    v.GodinaProizvodnje = reader["GodinaProizvodnje"].ToString();
                    v.BrojSasije = reader["BrojSasije"].ToString();
                    v.BrojMotora = reader["BrojMotora"].ToString();
                    v.Boja = reader["Boja"].ToString();
                    v.DatumPrveRegostracije = reader["DatumPrveRegostracije"].ToString();
                    v.Odnos = reader["Odnos"].ToString();
                    v.Snaga = reader["Snaga"].ToString();
                    v.Zapremina = reader["Zapremina"].ToString();
                    v.NajvecaMasa = reader["NajvecaMasa"].ToString();
                    v.Tezina = reader["Tezina"].ToString();
                    v.Nosivost = reader["Nosivost"].ToString();
                    v.Sedenje = reader["Sedenje"].ToString();
                    v.Stajanje = reader["Stajanje"].ToString();
                    v.BrojOsovina = reader["BrojOsovina"].ToString();
                    v.Gorivo = reader["Gorivo"].ToString();
                    v.Homologacija = reader["Homologacija"].ToString();
                    v.ImeVL = reader["ImeVL"].ToString();
                    v.PrezimeVL = reader["PrezimeVL"].ToString();
                    v.MaticniBrojVL = reader["MaticniBrojVL"].ToString();
                    v.BRLKVL = reader["BRLKVL"].ToString();
                    v.BrojPosteVL = reader["BrojPosteVL"].ToString();
                    v.MestoVL = reader["MestoVL"].ToString();
                    v.UlicaVL = reader["UlicaVL"].ToString();
                    v.BrojVL = reader["BrojVL"].ToString();
                    v.TelefonaVL = reader["TelefonaVL"].ToString();
                    v.ZiroRacunVL = reader["ZiroRacunVL"].ToString();
                    v.ImeDL = reader["ImeDL"].ToString();
                    v.PrezimeDL = reader["PrezimeDL"].ToString();
                    v.MaticniBrojDL = reader["MaticniBrojDL"].ToString();
                    v.BRLKDL = reader["BRLKDL"].ToString();
                    v.BrojPosteDL = reader["BrojPosteDL"].ToString();
                    v.MestoDL = reader["MestoDL"].ToString();
                    v.UlicaDL = reader["UlicaDL"].ToString();
                    v.BrojDL = reader["BrojDL"].ToString();
                    v.TelefonaDL = reader["TelefonaDL"].ToString();
                    v.Napomena = reader["Napomena"].ToString();
                    v.Datum = reader["Datum"].ToString();
                    v.VremeOd = reader["VremeOd"].ToString();
                    v.VremeDo = reader["VremeDo"].ToString();
                    v.Kontrolor1 = reader["Kontrolor1"].ToString();
                    v.Kontrolor2 = reader["Kontrolor2"].ToString();
                    v.Predjenikm = reader["Predjenikm"].ToString();
                    v.PoslednjeUverenjeBroj = reader["PoslednjeUverenjeBroj"].ToString();
                    v.PoslednjeUverenjeDatum = reader["PoslednjeUverenjeDatum"].ToString();
                    v.DecimalnaOznaka = reader["DecimalnaOznaka"].ToString();
                    v.OblikKaroserijeSkraceno = reader["OblikKaroserijeSkraceno"].ToString();
                    v.BojaL = reader["BojaL"].ToString(); v.BojaD = reader["BojaD"].ToString();
                    v.BojaD = reader["BojaD"].ToString();
                    v.Ispravnost = reader["Ispravnost"].ToString();
                    v.BrSati = reader["BrSati"].ToString();
                    v.Kuka = reader["Kuka"].ToString();
                    v.OpterecenjeN1 = reader["OpterecenjeN1"].ToString();
                    v.OpterecenjeN2 = reader["OpterecenjeN2"].ToString();
                    v.OpterecenjeN3 = reader["OpterecenjeN3"].ToString();
                    v.OpterecenjeN4 = reader["OpterecenjeN4"].ToString();
                    v.OpterecenjeN5 = reader["OpterecenjeN5"].ToString();
                    v.PneumatikL1 = reader["PneumatikL1"].ToString();
                    v.PneumatikL2 = reader["PneumatikL2"].ToString();
                    v.PneumatikL3 = reader["PneumatikL3"].ToString();
                    v.PneumatikL4 = reader["PneumatikL4"].ToString();
                    v.PneumatikL5 = reader["PneumatikL5"].ToString();
                    v.PneumatikD1 = reader["PneumatikD1"].ToString();
                    v.PneumatikD2 = reader["PneumatikD2"].ToString();
                    v.PneumatikD3 = reader["PneumatikD3"].ToString();
                    v.PneumatikD4 = reader["PneumatikD4"].ToString();
                    v.PneumatikD5 = reader["PneumatikD5"].ToString();
                    v.zaustavljanje = reader["Zaustavljanje"].ToString();
                    v.terensko = reader["terensko"].ToString();
                    v.komentar = reader["komentar"].ToString();
                    v.rejting = reader["rejting"].ToString();
                    v.decplusoznaka = reader["Decplusoznaka"].ToString();
                    v.datumIsteka = reader["datumIsteka"].ToString();
                    v.Poruka = reader["Poruka"].ToString();
                    voz.Add(v);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return voz;
        }
        internal List<vozilo> ucitajTablice()
        {
            List<vozilo> voz = new List<vozilo>();
            try
            {
                command.CommandText = "SELECT * FROM registracija order by RegistarskiBroj,CDate( Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4)) DESC;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    vozilo v = new vozilo();
                    v.IDBroj = reader["IDBroj"].ToString();
                    v.RegistarskiBroj = reader["RegistarskiBroj"].ToString();
                    v.VrstaVozila = reader["VrstaVozila"].ToString();
                    v.PodvrstaVozila = reader["PodvrstaVozila"].ToString();
                    v.klasa = reader["klasa"].ToString();
                    v.oblikkaroserije = reader["oblikkaroserije"].ToString();
                    v.OblikKaroserijeSkraceno = reader["oblikkaroserijeSkraceno"].ToString();
                    v.specijalno = reader["specijalno"].ToString();
                    v.Marka = reader["Marka"].ToString();
                    v.Model = reader["Model"].ToString();
                    v.Tip = reader["Tip"].ToString();
                    v.GodinaProizvodnje = reader["GodinaProizvodnje"].ToString();
                    v.BrojSasije = reader["BrojSasije"].ToString();
                    v.BrojMotora = reader["BrojMotora"].ToString();
                    v.Boja = reader["Boja"].ToString();
                    v.DatumPrveRegostracije = reader["DatumPrveRegostracije"].ToString();
                    v.Odnos = reader["Odnos"].ToString();
                    v.Snaga = reader["Snaga"].ToString();
                    v.Zapremina = reader["Zapremina"].ToString();
                    v.NajvecaMasa = reader["NajvecaMasa"].ToString();
                    v.Tezina = reader["Tezina"].ToString();
                    v.Nosivost = reader["Nosivost"].ToString();
                    v.Sedenje = reader["Sedenje"].ToString();
                    v.Stajanje = reader["Stajanje"].ToString();
                    v.BrojOsovina = reader["BrojOsovina"].ToString();
                    v.Gorivo = reader["Gorivo"].ToString();
                    v.Homologacija = reader["Homologacija"].ToString();
                    v.ImeVL = reader["ImeVL"].ToString();
                    v.PrezimeVL = reader["PrezimeVL"].ToString();
                    v.MaticniBrojVL = reader["MaticniBrojVL"].ToString();
                    v.BRLKVL = reader["BRLKVL"].ToString();
                    v.BrojPosteVL = reader["BrojPosteVL"].ToString();
                    v.MestoVL = reader["MestoVL"].ToString();
                    v.UlicaVL = reader["UlicaVL"].ToString();
                    v.BrojVL = reader["BrojVL"].ToString();
                    v.TelefonaVL = reader["TelefonaVL"].ToString();
                    v.ZiroRacunVL = reader["ZiroRacunVL"].ToString();
                    v.ImeDL = reader["ImeDL"].ToString();
                    v.PrezimeDL = reader["PrezimeDL"].ToString();
                    v.MaticniBrojDL = reader["MaticniBrojDL"].ToString();
                    v.BRLKDL = reader["BRLKDL"].ToString();
                    v.BrojPosteDL = reader["BrojPosteDL"].ToString();
                    v.MestoDL = reader["MestoDL"].ToString();
                    v.UlicaDL = reader["UlicaDL"].ToString();
                    v.BrojDL = reader["BrojDL"].ToString();
                    v.TelefonaDL = reader["TelefonaDL"].ToString();
                    v.Napomena = reader["Napomena"].ToString();
                    v.Datum = reader["Datum"].ToString();
                    v.VremeOd = reader["VremeOd"].ToString();
                    v.VremeDo = reader["VremeDo"].ToString();
                    v.Kontrolor1 = reader["Kontrolor1"].ToString();
                    v.Kontrolor2 = reader["Kontrolor2"].ToString();
                    v.Predjenikm = reader["Predjenikm"].ToString();
                    v.PoslednjeUverenjeBroj = reader["PoslednjeUverenjeBroj"].ToString();
                    v.PoslednjeUverenjeDatum = reader["PoslednjeUverenjeDatum"].ToString();
                    v.DecimalnaOznaka = reader["DecimalnaOznaka"].ToString();
                    v.OblikKaroserijeSkraceno = reader["OblikKaroserijeSkraceno"].ToString();
                    v.BojaL = reader["BojaL"].ToString(); v.BojaD = reader["BojaD"].ToString();
                    v.BojaD = reader["BojaD"].ToString();
                    v.Ispravnost = reader["Ispravnost"].ToString();
                    v.BrSati = reader["BrSati"].ToString();
                    v.Kuka = reader["Kuka"].ToString();
                    v.OpterecenjeN1 = reader["OpterecenjeN1"].ToString();
                    v.OpterecenjeN2 = reader["OpterecenjeN2"].ToString();
                    v.OpterecenjeN3 = reader["OpterecenjeN3"].ToString();
                    v.OpterecenjeN4 = reader["OpterecenjeN4"].ToString();
                    v.OpterecenjeN5 = reader["OpterecenjeN5"].ToString();
                    v.PneumatikL1 = reader["PneumatikL1"].ToString();
                    v.PneumatikL2 = reader["PneumatikL2"].ToString();
                    v.PneumatikL3 = reader["PneumatikL3"].ToString();
                    v.PneumatikL4 = reader["PneumatikL4"].ToString();
                    v.PneumatikL5 = reader["PneumatikL5"].ToString();
                    v.PneumatikD1 = reader["PneumatikD1"].ToString();
                    v.PneumatikD2 = reader["PneumatikD2"].ToString();
                    v.PneumatikD3 = reader["PneumatikD3"].ToString();
                    v.PneumatikD4 = reader["PneumatikD4"].ToString();
                    v.PneumatikD5 = reader["PneumatikD5"].ToString();
                    v.zaustavljanje = reader["Zaustavljanje"].ToString();
                    v.terensko = reader["terensko"].ToString();
                    v.komentar = reader["komentar"].ToString();
                    v.rejting = reader["rejting"].ToString();
                    v.decplusoznaka = reader["Decplusoznaka"].ToString();
                    v.datumIsteka = reader["datumIsteka"].ToString();
                    v.duzina = reader["Duzina"].ToString();
                    v.sirina = reader["Sirina"].ToString();
                    v.visina = reader["Visina"].ToString();
                    v.BrojDokumenta = reader["BrojDokumenta"].ToString();
                    v.SerijskiBroj = reader["SerijskiBroj"].ToString();
                    v.VlaznostKT = reader["VlaznostKT"].ToString();
                    v.Osiguranje = reader["Osiguranje"].ToString();
                    v.Prioritet = reader["Prioritet"].ToString();
                    v.DLZiroRacun = reader["DLZiroRacun"].ToString();
                    v.DatumIzdavanjaSaobracajne = reader["DatumIzdavanjaSaobracajne"].ToString();



                    voz.Add(v);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return voz;
        }
        internal List<Faktura> ucitajFakture()
        {
            List<Faktura> voz = new List<Faktura>();
            try
            {
                command.CommandText = "SELECT * FROM Faktura ORDER BY ID;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    Faktura v = new Faktura()
                    {
                        Cena = reader["Cena"].ToString(),
                        Grad = reader["Grad"].ToString(),
                        Komada = reader["Komada"].ToString(),
                        MaticniBroj = reader["MaticniBroj"].ToString(),
                        NazivArtikla = reader["NazivArtikla"].ToString(),
                        NazivFirme = reader["NazivFirme"].ToString(),
                        PIB = reader["PIB"].ToString(),
                        Rabat = reader["Rabat"].ToString(),
                        Telefon = reader["Telefon"].ToString(),
                        TR = reader["TR"].ToString(),
                        Ulica = reader["Ulica"].ToString(),
                    };
                    voz.Add(v);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return voz;
        }

        internal List<voziloSasije> ucitajSasije()
        {
            List<voziloSasije> voz = new List<voziloSasije>();
            try
            {
                command.CommandText = "SELECT * FROM registracija order by BrojSasije,CDate( Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4)) DESC;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    voziloSasije v = new voziloSasije();
                    v.IDBroj = reader["IDBroj"].ToString();
                    v.RegistarskiBroj = reader["RegistarskiBroj"].ToString();
                    v.VrstaVozila = reader["VrstaVozila"].ToString();
                    v.PodvrstaVozila = reader["PodvrstaVozila"].ToString();
                    v.klasa = reader["klasa"].ToString();
                    v.oblikkaroserije = reader["oblikkaroserije"].ToString();
                    v.OblikKaroserijeSkraceno = reader["oblikkaroserijeSkraceno"].ToString();
                    v.specijalno = reader["specijalno"].ToString();
                    v.Marka = reader["Marka"].ToString();
                    v.Model = reader["Model"].ToString();
                    v.Tip = reader["Tip"].ToString();
                    v.GodinaProizvodnje = reader["GodinaProizvodnje"].ToString();
                    v.BrojSasije = reader["BrojSasije"].ToString();
                    v.BrojMotora = reader["BrojMotora"].ToString();
                    v.Boja = reader["Boja"].ToString();
                    v.DatumPrveRegostracije = reader["DatumPrveRegostracije"].ToString();
                    v.Odnos = reader["Odnos"].ToString();
                    v.Snaga = reader["Snaga"].ToString();
                    v.Zapremina = reader["Zapremina"].ToString();
                    v.NajvecaMasa = reader["NajvecaMasa"].ToString();
                    v.Tezina = reader["Tezina"].ToString();
                    v.Nosivost = reader["Nosivost"].ToString();
                    v.Sedenje = reader["Sedenje"].ToString();
                    v.Stajanje = reader["Stajanje"].ToString();
                    v.BrojOsovina = reader["BrojOsovina"].ToString();
                    v.Gorivo = reader["Gorivo"].ToString();
                    v.Homologacija = reader["Homologacija"].ToString();
                    v.ImeVL = reader["ImeVL"].ToString();
                    v.PrezimeVL = reader["PrezimeVL"].ToString();
                    v.MaticniBrojVL = reader["MaticniBrojVL"].ToString();
                    v.BRLKVL = reader["BRLKVL"].ToString();
                    v.BrojPosteVL = reader["BrojPosteVL"].ToString();
                    v.MestoVL = reader["MestoVL"].ToString();
                    v.UlicaVL = reader["UlicaVL"].ToString();
                    v.BrojVL = reader["BrojVL"].ToString();
                    v.TelefonaVL = reader["TelefonaVL"].ToString();
                    v.ZiroRacunVL = reader["ZiroRacunVL"].ToString();
                    v.ImeDL = reader["ImeDL"].ToString();
                    v.PrezimeDL = reader["PrezimeDL"].ToString();
                    v.MaticniBrojDL = reader["MaticniBrojDL"].ToString();
                    v.BRLKDL = reader["BRLKDL"].ToString();
                    v.BrojPosteDL = reader["BrojPosteDL"].ToString();
                    v.MestoDL = reader["MestoDL"].ToString();
                    v.UlicaDL = reader["UlicaDL"].ToString();
                    v.BrojDL = reader["BrojDL"].ToString();
                    v.TelefonaDL = reader["TelefonaDL"].ToString();
                    v.Napomena = reader["Napomena"].ToString();
                    v.Datum = reader["Datum"].ToString();
                    v.VremeOd = reader["VremeOd"].ToString();
                    v.VremeDo = reader["VremeDo"].ToString();
                    v.Kontrolor1 = reader["Kontrolor1"].ToString();
                    v.Kontrolor2 = reader["Kontrolor2"].ToString();
                    v.Predjenikm = reader["Predjenikm"].ToString();
                    v.PoslednjeUverenjeBroj = reader["PoslednjeUverenjeBroj"].ToString();
                    v.PoslednjeUverenjeDatum = reader["PoslednjeUverenjeDatum"].ToString();
                    v.DecimalnaOznaka = reader["DecimalnaOznaka"].ToString();
                    v.OblikKaroserijeSkraceno = reader["OblikKaroserijeSkraceno"].ToString();
                    v.BojaL = reader["BojaL"].ToString(); v.BojaD = reader["BojaD"].ToString();
                    v.BojaD = reader["BojaD"].ToString();
                    v.Ispravnost = reader["Ispravnost"].ToString();
                    v.BrSati = reader["BrSati"].ToString();
                    v.Kuka = reader["Kuka"].ToString();
                    v.OpterecenjeN1 = reader["OpterecenjeN1"].ToString();
                    v.OpterecenjeN2 = reader["OpterecenjeN2"].ToString();
                    v.OpterecenjeN3 = reader["OpterecenjeN3"].ToString();
                    v.OpterecenjeN4 = reader["OpterecenjeN4"].ToString();
                    v.OpterecenjeN5 = reader["OpterecenjeN5"].ToString();
                    v.PneumatikL1 = reader["PneumatikL1"].ToString();
                    v.PneumatikL2 = reader["PneumatikL2"].ToString();
                    v.PneumatikL3 = reader["PneumatikL3"].ToString();
                    v.PneumatikL4 = reader["PneumatikL4"].ToString();
                    v.PneumatikL5 = reader["PneumatikL5"].ToString();
                    v.PneumatikD1 = reader["PneumatikD1"].ToString();
                    v.PneumatikD2 = reader["PneumatikD2"].ToString();
                    v.PneumatikD3 = reader["PneumatikD3"].ToString();
                    v.PneumatikD4 = reader["PneumatikD4"].ToString();
                    v.PneumatikD5 = reader["PneumatikD5"].ToString();
                    v.zaustavljanje = reader["Zaustavljanje"].ToString();
                    v.terensko = reader["terensko"].ToString();
                    v.komentar = reader["komentar"].ToString();
                    v.rejting = reader["rejting"].ToString();
                    v.decplusoznaka = reader["Decplusoznaka"].ToString();
                    v.datumIsteka = reader["datumIsteka"].ToString();
                    v.duzina = reader["Duzina"].ToString();
                    v.sirina = reader["Sirina"].ToString();
                    v.visina = reader["Visina"].ToString();
                    v.BrojDokumenta = reader["BrojDokumenta"].ToString();
                    v.SerijskiBroj = reader["SerijskiBroj"].ToString();
                    v.VlaznostKT = reader["VlaznostKT"].ToString();
                    v.Osiguranje = reader["Osiguranje"].ToString();
                    v.Prioritet = reader["Prioritet"].ToString();
                    v.DLZiroRacun = reader["DLZiroRacun"].ToString();
                    v.DatumIzdavanjaSaobracajne = reader["DatumIzdavanjaSaobracajne"].ToString();

                    voz.Add(v);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return voz;
        }
        internal List<voziloId> ucitajIdjeve()
        {
            List<voziloId> voz = new List<voziloId>();
            try
            {
                command.CommandText = "SELECT * FROM registracija order by IDBroj DESC;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    voziloId v = new voziloId();
                    v.IDBroj = reader["IDBroj"].ToString();
                    v.RegistarskiBroj = reader["RegistarskiBroj"].ToString();
                    v.VrstaVozila = reader["VrstaVozila"].ToString();
                    v.PodvrstaVozila = reader["PodvrstaVozila"].ToString();
                    v.klasa = reader["klasa"].ToString();
                    v.oblikkaroserije = reader["oblikkaroserije"].ToString();
                    v.OblikKaroserijeSkraceno = reader["oblikkaroserijeSkraceno"].ToString();
                    v.specijalno = reader["specijalno"].ToString();
                    v.Marka = reader["Marka"].ToString();
                    v.Model = reader["Model"].ToString();
                    v.Tip = reader["Tip"].ToString();
                    v.GodinaProizvodnje = reader["GodinaProizvodnje"].ToString();
                    v.BrojSasije = reader["BrojSasije"].ToString();
                    v.BrojMotora = reader["BrojMotora"].ToString();
                    v.Boja = reader["Boja"].ToString();
                    v.DatumPrveRegostracije = reader["DatumPrveRegostracije"].ToString();
                    v.Odnos = reader["Odnos"].ToString();
                    v.Snaga = reader["Snaga"].ToString();
                    v.Zapremina = reader["Zapremina"].ToString();
                    v.NajvecaMasa = reader["NajvecaMasa"].ToString();
                    v.Tezina = reader["Tezina"].ToString();
                    v.Nosivost = reader["Nosivost"].ToString();
                    v.Sedenje = reader["Sedenje"].ToString();
                    v.Stajanje = reader["Stajanje"].ToString();
                    v.BrojOsovina = reader["BrojOsovina"].ToString();
                    v.Gorivo = reader["Gorivo"].ToString();
                    v.Homologacija = reader["Homologacija"].ToString();
                    v.ImeVL = reader["ImeVL"].ToString();
                    v.PrezimeVL = reader["PrezimeVL"].ToString();
                    v.MaticniBrojVL = reader["MaticniBrojVL"].ToString();
                    v.BRLKVL = reader["BRLKVL"].ToString();
                    v.BrojPosteVL = reader["BrojPosteVL"].ToString();
                    v.MestoVL = reader["MestoVL"].ToString();
                    v.UlicaVL = reader["UlicaVL"].ToString();
                    v.BrojVL = reader["BrojVL"].ToString();
                    v.TelefonaVL = reader["TelefonaVL"].ToString();
                    v.ZiroRacunVL = reader["ZiroRacunVL"].ToString();
                    v.ImeDL = reader["ImeDL"].ToString();
                    v.PrezimeDL = reader["PrezimeDL"].ToString();
                    v.MaticniBrojDL = reader["MaticniBrojDL"].ToString();
                    v.BRLKDL = reader["BRLKDL"].ToString();
                    v.BrojPosteDL = reader["BrojPosteDL"].ToString();
                    v.MestoDL = reader["MestoDL"].ToString();
                    v.UlicaDL = reader["UlicaDL"].ToString();
                    v.BrojDL = reader["BrojDL"].ToString();
                    v.TelefonaDL = reader["TelefonaDL"].ToString();
                    v.Napomena = reader["Napomena"].ToString();
                    v.Datum = reader["Datum"].ToString();
                    v.VremeOd = reader["VremeOd"].ToString();
                    v.VremeDo = reader["VremeDo"].ToString();
                    v.Kontrolor1 = reader["Kontrolor1"].ToString();
                    v.Kontrolor2 = reader["Kontrolor2"].ToString();
                    v.Predjenikm = reader["Predjenikm"].ToString();
                    v.PoslednjeUverenjeBroj = reader["PoslednjeUverenjeBroj"].ToString();
                    v.PoslednjeUverenjeDatum = reader["PoslednjeUverenjeDatum"].ToString();
                    v.DecimalnaOznaka = reader["DecimalnaOznaka"].ToString();
                    v.OblikKaroserijeSkraceno = reader["OblikKaroserijeSkraceno"].ToString();
                    v.BojaL = reader["BojaL"].ToString(); v.BojaD = reader["BojaD"].ToString();
                    v.BojaD = reader["BojaD"].ToString();
                    v.Ispravnost = reader["Ispravnost"].ToString();
                    v.BrSati = reader["BrSati"].ToString();
                    v.Kuka = reader["Kuka"].ToString();
                    v.OpterecenjeN1 = reader["OpterecenjeN1"].ToString();
                    v.OpterecenjeN2 = reader["OpterecenjeN2"].ToString();
                    v.OpterecenjeN3 = reader["OpterecenjeN3"].ToString();
                    v.OpterecenjeN4 = reader["OpterecenjeN4"].ToString();
                    v.OpterecenjeN5 = reader["OpterecenjeN5"].ToString();
                    v.PneumatikL1 = reader["PneumatikL1"].ToString();
                    v.PneumatikL2 = reader["PneumatikL2"].ToString();
                    v.PneumatikL3 = reader["PneumatikL3"].ToString();
                    v.PneumatikL4 = reader["PneumatikL4"].ToString();
                    v.PneumatikL5 = reader["PneumatikL5"].ToString();
                    v.PneumatikD1 = reader["PneumatikD1"].ToString();
                    v.PneumatikD2 = reader["PneumatikD2"].ToString();
                    v.PneumatikD3 = reader["PneumatikD3"].ToString();
                    v.PneumatikD4 = reader["PneumatikD4"].ToString();
                    v.PneumatikD5 = reader["PneumatikD5"].ToString();
                    v.zaustavljanje = reader["Zaustavljanje"].ToString();
                    v.terensko = reader["terensko"].ToString();
                    v.komentar = reader["komentar"].ToString();
                    v.rejting = reader["rejting"].ToString();
                    v.decplusoznaka = reader["Decplusoznaka"].ToString();
                    v.datumIsteka = reader["datumIsteka"].ToString();
                    v.duzina = reader["Duzina"].ToString();
                    v.sirina = reader["Sirina"].ToString();
                    v.visina = reader["Visina"].ToString();
                    v.BrojDokumenta = reader["BrojDokumenta"].ToString();
                    v.SerijskiBroj = reader["SerijskiBroj"].ToString();
                    v.VlaznostKT = reader["VlaznostKT"].ToString();
                    v.Osiguranje = reader["Osiguranje"].ToString();
                    v.Prioritet = reader["Prioritet"].ToString();
                    v.DLZiroRacun = reader["DLZiroRacun"].ToString();
                    v.DatumIzdavanjaSaobracajne = reader["DatumIzdavanjaSaobracajne"].ToString();

                    voz.Add(v);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return voz;
        }
        internal List<vozilo> ucitajSortPoId()
        {
            List<vozilo> voz = new List<vozilo>();
            try
            {
                command.CommandText = "SELECT * FROM registracija order by idbroj;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    vozilo v = new vozilo();
                    v.IDBroj = reader["IDBroj"].ToString();
                    v.RegistarskiBroj = reader["RegistarskiBroj"].ToString();
                    v.VrstaVozila = reader["VrstaVozila"].ToString();
                    v.PodvrstaVozila = reader["PodvrstaVozila"].ToString();
                    v.klasa = reader["klasa"].ToString();
                    v.oblikkaroserije = reader["oblikkaroserije"].ToString();
                    v.OblikKaroserijeSkraceno = reader["oblikkaroserijeSkraceno"].ToString();
                    v.specijalno = reader["specijalno"].ToString();
                    v.Marka = reader["Marka"].ToString();
                    v.Model = reader["Model"].ToString();
                    v.Tip = reader["Tip"].ToString();
                    v.GodinaProizvodnje = reader["GodinaProizvodnje"].ToString();
                    v.BrojSasije = reader["BrojSasije"].ToString();
                    v.BrojMotora = reader["BrojMotora"].ToString();
                    v.Boja = reader["Boja"].ToString();
                    v.DatumPrveRegostracije = reader["DatumPrveRegostracije"].ToString();
                    v.Odnos = reader["Odnos"].ToString();
                    v.Snaga = reader["Snaga"].ToString();
                    v.Zapremina = reader["Zapremina"].ToString();
                    v.NajvecaMasa = reader["NajvecaMasa"].ToString();
                    v.Tezina = reader["Tezina"].ToString();
                    v.Nosivost = reader["Nosivost"].ToString();
                    v.Sedenje = reader["Sedenje"].ToString();
                    v.Stajanje = reader["Stajanje"].ToString();
                    v.BrojOsovina = reader["BrojOsovina"].ToString();
                    v.Gorivo = reader["Gorivo"].ToString();
                    v.Homologacija = reader["Homologacija"].ToString();
                    v.ImeVL = reader["ImeVL"].ToString();
                    v.PrezimeVL = reader["PrezimeVL"].ToString();
                    v.MaticniBrojVL = reader["MaticniBrojVL"].ToString();
                    v.BRLKVL = reader["BRLKVL"].ToString();
                    v.BrojPosteVL = reader["BrojPosteVL"].ToString();
                    v.MestoVL = reader["MestoVL"].ToString();
                    v.UlicaVL = reader["UlicaVL"].ToString();
                    v.BrojVL = reader["BrojVL"].ToString();
                    v.TelefonaVL = reader["TelefonaVL"].ToString();
                    v.ZiroRacunVL = reader["ZiroRacunVL"].ToString();
                    v.ImeDL = reader["ImeDL"].ToString();
                    v.PrezimeDL = reader["PrezimeDL"].ToString();
                    v.MaticniBrojDL = reader["MaticniBrojDL"].ToString();
                    v.BRLKDL = reader["BRLKDL"].ToString();
                    v.BrojPosteDL = reader["BrojPosteDL"].ToString();
                    v.MestoDL = reader["MestoDL"].ToString();
                    v.UlicaDL = reader["UlicaDL"].ToString();
                    v.BrojDL = reader["BrojDL"].ToString();
                    v.TelefonaDL = reader["TelefonaDL"].ToString();
                    v.Napomena = reader["Napomena"].ToString();
                    v.Datum = reader["Datum"].ToString();
                    v.VremeOd = reader["VremeOd"].ToString();
                    v.VremeDo = reader["VremeDo"].ToString();
                    v.Kontrolor1 = reader["Kontrolor1"].ToString();
                    v.Kontrolor2 = reader["Kontrolor2"].ToString();
                    v.Predjenikm = reader["Predjenikm"].ToString();
                    v.PoslednjeUverenjeBroj = reader["PoslednjeUverenjeBroj"].ToString();
                    v.PoslednjeUverenjeDatum = reader["PoslednjeUverenjeDatum"].ToString();
                    v.DecimalnaOznaka = reader["DecimalnaOznaka"].ToString();
                    v.OblikKaroserijeSkraceno = reader["OblikKaroserijeSkraceno"].ToString();
                    v.BojaL = reader["BojaL"].ToString(); v.BojaD = reader["BojaD"].ToString();
                    v.BojaD = reader["BojaD"].ToString();
                    v.Ispravnost = reader["Ispravnost"].ToString();
                    v.BrSati = reader["BrSati"].ToString();
                    v.Kuka = reader["Kuka"].ToString();
                    v.OpterecenjeN1 = reader["OpterecenjeN1"].ToString();
                    v.OpterecenjeN2 = reader["OpterecenjeN2"].ToString();
                    v.OpterecenjeN3 = reader["OpterecenjeN3"].ToString();
                    v.OpterecenjeN4 = reader["OpterecenjeN4"].ToString();
                    v.OpterecenjeN5 = reader["OpterecenjeN5"].ToString();
                    v.PneumatikL1 = reader["PneumatikL1"].ToString();
                    v.PneumatikL2 = reader["PneumatikL2"].ToString();
                    v.PneumatikL3 = reader["PneumatikL3"].ToString();
                    v.PneumatikL4 = reader["PneumatikL4"].ToString();
                    v.PneumatikL5 = reader["PneumatikL5"].ToString();
                    v.PneumatikD1 = reader["PneumatikD1"].ToString();
                    v.PneumatikD2 = reader["PneumatikD2"].ToString();
                    v.PneumatikD3 = reader["PneumatikD3"].ToString();
                    v.PneumatikD4 = reader["PneumatikD4"].ToString();
                    v.PneumatikD5 = reader["PneumatikD5"].ToString();
                    v.zaustavljanje = reader["Zaustavljanje"].ToString();
                    v.terensko = reader["terensko"].ToString();
                    v.komentar = reader["komentar"].ToString();
                    v.rejting = reader["rejting"].ToString();
                    v.decplusoznaka = reader["Decplusoznaka"].ToString();
                    v.datumIsteka = reader["datumIsteka"].ToString();

                    voz.Add(v);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return voz;
        }

        internal List<vozilo2> ucitajMusterije(string s)
        {
            List<vozilo2> voz = new List<vozilo2>();
            try
            {
                command.CommandText = "SELECT * FROM registracija WHERE (ImeVL LIKE '%" + s + "%' OR PrezimeVL LIKE '%" + s + "%' ) OR (ImeDL LIKE '%" + s + "%' OR PrezimeDL LIKE '%" + s + "%' );";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    vozilo2 v = new vozilo2();
                    v.IDBroj = reader["IDBroj"].ToString();
                    v.RegistarskiBroj = reader["RegistarskiBroj"].ToString();
                    v.VrstaVozila = reader["VrstaVozila"].ToString();
                    v.PodvrstaVozila = reader["PodvrstaVozila"].ToString();
                    v.klasa = reader["klasa"].ToString();
                    v.oblikkaroserije = reader["oblikkaroserije"].ToString();
                    v.OblikKaroserijeSkraceno = reader["oblikkaroserijeSkraceno"].ToString();
                    v.specijalno = reader["specijalno"].ToString();
                    v.Marka = reader["Marka"].ToString();
                    v.Model = reader["Model"].ToString();
                    v.Tip = reader["Tip"].ToString();
                    v.GodinaProizvodnje = reader["GodinaProizvodnje"].ToString();
                    v.BrojSasije = reader["BrojSasije"].ToString();
                    v.BrojMotora = reader["BrojMotora"].ToString();
                    v.Boja = reader["Boja"].ToString();
                    v.DatumPrveRegostracije = reader["DatumPrveRegostracije"].ToString();
                    v.Odnos = reader["Odnos"].ToString();
                    v.Snaga = reader["Snaga"].ToString();
                    v.Zapremina = reader["Zapremina"].ToString();
                    v.NajvecaMasa = reader["NajvecaMasa"].ToString();
                    v.Tezina = reader["Tezina"].ToString();
                    v.Nosivost = reader["Nosivost"].ToString();
                    v.Sedenje = reader["Sedenje"].ToString();
                    v.Stajanje = reader["Stajanje"].ToString();
                    v.BrojOsovina = reader["BrojOsovina"].ToString();
                    v.Gorivo = reader["Gorivo"].ToString();
                    v.Homologacija = reader["Homologacija"].ToString();
                    v.ImeVL = reader["ImeVL"].ToString();
                    v.PrezimeVL = reader["PrezimeVL"].ToString();
                    v.MaticniBrojVL = reader["MaticniBrojVL"].ToString();
                    v.BRLKVL = reader["BRLKVL"].ToString();
                    v.BrojPosteVL = reader["BrojPosteVL"].ToString();
                    v.MestoVL = reader["MestoVL"].ToString();
                    v.UlicaVL = reader["UlicaVL"].ToString();
                    v.BrojVL = reader["BrojVL"].ToString();
                    v.TelefonaVL = reader["TelefonaVL"].ToString();
                    v.ZiroRacunVL = reader["ZiroRacunVL"].ToString();
                    v.ImeDL = reader["ImeDL"].ToString();
                    v.PrezimeDL = reader["PrezimeDL"].ToString();
                    v.MaticniBrojDL = reader["MaticniBrojDL"].ToString();
                    v.BRLKDL = reader["BRLKDL"].ToString();
                    v.BrojPosteDL = reader["BrojPosteDL"].ToString();
                    v.MestoDL = reader["MestoDL"].ToString();
                    v.UlicaDL = reader["UlicaDL"].ToString();
                    v.BrojDL = reader["BrojDL"].ToString();
                    v.TelefonaDL = reader["TelefonaDL"].ToString();
                    v.Napomena = reader["Napomena"].ToString();
                    v.Datum = reader["Datum"].ToString();
                    v.VremeOd = reader["VremeOd"].ToString();
                    v.VremeDo = reader["VremeDo"].ToString();
                    v.Kontrolor1 = reader["Kontrolor1"].ToString();
                    v.Kontrolor2 = reader["Kontrolor2"].ToString();
                    v.Predjenikm = reader["Predjenikm"].ToString();
                    v.PoslednjeUverenjeBroj = reader["PoslednjeUverenjeBroj"].ToString();
                    v.PoslednjeUverenjeDatum = reader["PoslednjeUverenjeDatum"].ToString();
                    v.DecimalnaOznaka = reader["DecimalnaOznaka"].ToString();
                    v.OblikKaroserijeSkraceno = reader["OblikKaroserijeSkraceno"].ToString();
                    v.BojaL = reader["BojaL"].ToString(); v.BojaD = reader["BojaD"].ToString();
                    v.BojaD = reader["BojaD"].ToString();
                    v.Ispravnost = reader["Ispravnost"].ToString();
                    v.BrSati = reader["BrSati"].ToString();
                    v.Kuka = reader["Kuka"].ToString();
                    v.OpterecenjeN1 = reader["OpterecenjeN1"].ToString();
                    v.OpterecenjeN2 = reader["OpterecenjeN2"].ToString();
                    v.OpterecenjeN3 = reader["OpterecenjeN3"].ToString();
                    v.OpterecenjeN4 = reader["OpterecenjeN4"].ToString();
                    v.OpterecenjeN5 = reader["OpterecenjeN5"].ToString();
                    v.PneumatikL1 = reader["PneumatikL1"].ToString();
                    v.PneumatikL2 = reader["PneumatikL2"].ToString();
                    v.PneumatikL3 = reader["PneumatikL3"].ToString();
                    v.PneumatikL4 = reader["PneumatikL4"].ToString();
                    v.PneumatikL5 = reader["PneumatikL5"].ToString();
                    v.PneumatikD1 = reader["PneumatikD1"].ToString();
                    v.PneumatikD2 = reader["PneumatikD2"].ToString();
                    v.PneumatikD3 = reader["PneumatikD3"].ToString();
                    v.PneumatikD4 = reader["PneumatikD4"].ToString();
                    v.PneumatikD5 = reader["PneumatikD5"].ToString();
                    v.zaustavljanje = reader["Zaustavljanje"].ToString();
                    v.terensko = reader["terensko"].ToString();
                    v.komentar = reader["komentar"].ToString();
                    v.rejting = reader["rejting"].ToString();
                    v.decplusoznaka = reader["Decplusoznaka"].ToString();
                    v.datumIsteka = reader["datumIsteka"].ToString();
                    v.duzina = reader["Duzina"].ToString();
                    v.sirina = reader["Sirina"].ToString();
                    v.visina = reader["Visina"].ToString();
                    v.BrojDokumenta = reader["BrojDokumenta"].ToString();
                    v.SerijskiBroj = reader["SerijskiBroj"].ToString();
                    v.VlaznostKT = reader["VlaznostKT"].ToString();
                    v.DatumIzdavanjaSaobracajne = reader["DatumIzdavanjaSaobracajne"].ToString();



                    voz.Add(v);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return voz;
        }

        internal List<boje> lista_boja()
        {
            List<boje> lista = new List<boje>();
            try
            {
                command.CommandText = "SELECT * FROM boje";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    boje b = new boje();
                    b.BojaSifra = reader["BojaSifra"].ToString();
                    b.BojaNaziv = reader["BojaNaziv"].ToString();
                    lista.Add(b);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista;

        }
        internal List<karoserije> lista_marka()
        {
            List<karoserije> lista = new List<karoserije>();
            try
            {
                command.CommandText = "SELECT * FROM marke";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    karoserije b = new karoserije();
                    b.KaroserijaNaziv = reader["NazivMarke"].ToString();
                    lista.Add(b);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista;

        }

        internal List<karoserije> lista_kar_sif(string kar)
        {
            List<karoserije> lista = new List<karoserije>();
            try
            {
                command.CommandText = "SELECT * FROM VrsteKaroserija WHERE karoserija = '" + kar + "';";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    karoserije b = new karoserije();
                    b.KaroserijaNaziv = reader["ID"].ToString();
                    lista.Add(b);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista;

        }


        internal void unesiBoju(string sifra, string naziv)
        {
            try
            {
                command.CommandText = "INSERT INTO boje VALUES ( '" + sifra + "',  '" + naziv + "');";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Uspešno ste izvrsili snimanje");
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }

        }

        internal void unesiKorisnika(string IDBroj, string RegistarskiBroj, string VrstaVozila, string PodvrstaVozila, string klasa, string oblikkaroserije,
         string specijalno, string Marka, string Model, string Tip, string GodinaProizvodnje, string BrojSasije,
         string BrojMotora, string Boja, string DatumPrveRegostracije, string Odnos, string Snaga, string Zapremina,
         string NajvecaMasa, string Tezina, string Nosivost, string Sedenje, string Stajanje, string BrojOsovina,
         string Gorivo, string Homologacija, string ImeVL, string PrezimeVL, string MaticniBrojVL, string BRLKVL, string BrojPosteVL,
         string MestoVL, string UlicaVL, string BrojVL, string TelefonaVL, string ZiroRacunVL, string ImeDL, string PrezimeDL,
         string MaticniBrojDL, string BRLKDL, string BrojPosteDL, string MestoDL, string UlicaDL, string BrojDL, string TelefonaDL,
         string Napomena, string Datum, string VremeOd, string VremeDo,
         string Kontrolor1, string Kontrolor2, string Predjenikm,
         string PoslednjeUverenjeBroj, string PoslednjeUverenjeDatum, string DecimalnaOznaka, string OblikKaroserijeSkraceno, string BojaL,
         string BojaD, string Ispravnost, string BrSati, string Kuka, string OpterecenjeN1, string OpterecenjeN2,
         string OpterecenjeN3, string OpterecenjeN4, string OpterecenjeN5, string PneumatikL1, string PneumatikL2, string PneumatikL3,
         string PneumatikL4, string PneumatikL5, string PneumatikD1, string PneumatikD2, string PneumatikD3, string PneumatikD4, string PneumatikD5, string zaustavljanje, string terensko
            , string komentar, string rejting, string decplusoznaka, string datumIsteka, string duzina, string sirina, string visina, string BrojDokumenta, string SerijskiBroj, string Poruka, string VlaznostKT, string Osiguranje, string Prioritet, string DLZiroRacun, string DatumIzdavanjaSaobracajne)
        {
            try
            {
                command.CommandText = "INSERT INTO registracija VALUES ( '" + IDBroj + "',  '" + RegistarskiBroj + "',  '" + VrstaVozila + "',  '" + PodvrstaVozila + "',  '" + klasa
                    + "', '" + oblikkaroserije + "',  '" + specijalno + "', '" + Marka + "',  '" + Model + "', '" + Tip
                    + "', '" + GodinaProizvodnje + "',  '" + BrojSasije + "',  '" + BrojMotora + "',  '" + Boja + "',  '" + DatumPrveRegostracije
                    + "', '" + Odnos + "',  '" + Snaga + "', '" + Zapremina + "',  '" + NajvecaMasa + "',  '" + Tezina + "',  '" + Nosivost + "',  '" + Sedenje
                    + "', '" + Stajanje + "',  '" + BrojOsovina + "', '" + Gorivo + "',  '" + Homologacija + "',  '" + ImeVL
                    + "', '" + PrezimeVL + "',  '" + MaticniBrojVL + "',  '" + BRLKVL + "',  '" + BrojPosteVL + "',  '" + MestoVL + "',  '" + UlicaVL + "',  '" + BrojVL
                    + "', '" + TelefonaVL + "',  '" + ZiroRacunVL + "',  '" + ImeDL + "',  '" + PrezimeDL + "',  '" + MaticniBrojDL + "',  '" + BRLKDL + "',  '" + BrojPosteDL
                    + "', '" + MestoDL + "',  '" + UlicaDL + "',  '" + BrojDL + "',  '" + TelefonaDL + "',  '" + Napomena + "',  '" + Datum + "',  '" + VremeOd + "', '" + VremeDo
                    + "', '" + Kontrolor1 + "', '" + Kontrolor2 + "',  '" + Predjenikm + "', '" + PoslednjeUverenjeBroj + "', '" + PoslednjeUverenjeDatum
                    + "', '" + DecimalnaOznaka + "', '" + OblikKaroserijeSkraceno + "', '" + BojaL + "', '" + BojaD + "', '" + Ispravnost + "', '" + BrSati
                    + "', '" + Kuka + "', '" + OpterecenjeN1 + "', '" + OpterecenjeN2 + "', '" + OpterecenjeN3 + "', '" + OpterecenjeN4 + "', '" + OpterecenjeN5 + "', '" + PneumatikL1
                    + "', '" + PneumatikL2 + "', '" + PneumatikL3 + "', '" + PneumatikL4 + "', '" + PneumatikL5 + "', '" + PneumatikD1 + "', '" + PneumatikD2 + "', '" + PneumatikD3
                    + "', '" + PneumatikD4 + "', '" + PneumatikD5 + "', '" + zaustavljanje + "', '" + terensko + "', '" + komentar + "', '" + rejting + "', '" + decplusoznaka + "', '"
                    + datumIsteka + "', '" + duzina + "', '" + sirina + "', '" + visina + "', '" + BrojDokumenta + "', '" + SerijskiBroj + "', '" + Poruka + "', '" + VlaznostKT + "', '" + Osiguranje + "', '" + Prioritet + "', '" + DLZiroRacun + "','" +DatumIzdavanjaSaobracajne +"');";


                //command.CommandText = "INSERT INTO registracija VALUES ( '" + IDBroj + "',  '" + RegistarskiBroj + "',  '" + VrstaVozila + "', '',  '" + klasa
                //    + "', '" + oblikkaroserije + "',             '" + specijalno + "', '" + zemljaproiz + "', '" + Marka + "',  '" + Model + "', '" + Tip
                //    + "', '" + GodinaProizvodnje + "',  '" + BrojSasije + "',  '" + BrojMotora + "',  '" + Namena + "',             '" + Boja + "',  '" + DatumPrveRegostracije
                //    + "',  '" + Odnos + "',  '" + Snaga + "', '" + Zapremina + "',  '" + NajvecaMasa + "',  '" + Tezina + "',  '" + Nosivost + "',  '" + Sedenje
                //    + "',  '" + Stajanje + "',  '" + BrojOsovina + "', '" + PogonskihOsovina + "',  '" + Gorivo + "',  '" + Homologacija + "',  '" + MaticniBrojFL
                //    + "',  '" + PrezimeFL + "',  '" + ImeFL + "',  '" + DrzavljanstvoFL + "',  '" + OpstinaFL + "',  '" + BrojPosteFL + "', '" + MestoFL + "',  '" + UlicaFL
                //    + "',  '" + BrojFL + "',  '" + TelefonaFL + "',  '" + EmailFL + "',             '" + NazivPL + "',  '" + MaticniBrojPL + "',  '" + PIB + "',  '" + Delatnost
                //    + "',  '" + OpstinaPL + "',  '" + UlicaPL + "',  '" + BrojPL + "',  '" + BrojPostePL + "',             '" + MestoPL + "',  '" + BrojTelefonaPL
                //    + "',  '" + EmailPL + "',  '" + Napomena + "',  '" + Datum + "',  '" + VremeOd + "', '" + VremeDo + "', '" + ImePrezimeStranke + "', '" + BrojDokumenta
                //    + "',  '" + Kontrolor1 + "', '" + Kontrolor2 + "', '" + brTelStranke + "',  '" + Predjenikm + "', '" + PoslednjeUverenjeBroj + "', '" + PoslednjeUverenjeDatum
                //    + "', '" + DecimalnaOznaka + "', '" + OblikKaroserijeSkraceno + "', '" + BojaL + "', '" + BojaD + "', '" + Ispravnost + "', '" + Podvrsta + "');";


                //command.CommandText = "INSERT INTO registracija VALUES ('" + IDBroj + "','" + RegistarskiBroj + "','" + VrstaVozila + "','" + PodvrstaVozila + "','" + klasa + "','" + RegistarskiBroj + "','" + VrstaVozila + 
                //    "','" + oblikkaroserije + "','" + specijalno + "','" + zemljaproiz + "','" + Marka + "','" + Model + "','" + Tip + "','" +
                //    GodinaProizvodnje + "','" + BrojSasije + "','" + BrojMotora + "','" + Namena + "','" + Boja + "','" + DatumPrveRegostracije + "','" + 
                //    "','" + Odnos + "','" + Snaga + "','" + Zapremina + "','" + NajvecaMasa + "','" + Tezina + "','" + Nosivost + "','" + Sedenje + "','" + Stajanje + "','" + BrojOsovina + "','" +
                //    PogonskihOsovina + "','" + Gorivo + "','" + Homologacija + "','" + MaticniBrojFL + "','" + PrezimeFL + "','" + ImeFL + "','" +
                //    DrzavljanstvoFL + "','" + OpstinaFL + "','" + BrojPosteFL + "','" + MestoFL + "','" + UlicaFL + "','" + BrojFL + "','" + TelefonaFL +
                //    "','" + EmailFL + "','" + NazivPL + "','" + MaticniBrojPL + "','" + PIB + "','" + Delatnost + "','" + OpstinaPL + "','" + UlicaPL +
                //    "','" + BrojPL + "','" + BrojPostePL + "','" + MestoPL + "','" + BrojTelefonaPL + "','" + EmailPL + "','" + Napomena + "','" + Datum +
                //    "','"  + VremeOd + "','" + VremeDo + "','" + ImePrezimeStranke + "','" + BrojDokumenta + "','" + Kontrolor1 + "','" + Kontrolor2 + "')";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Uspešno ste izvrsili snimanje");
            }
            catch (OleDbException ex)
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
        }

        //internal bool instaliranje()
        //{
        //    try
        //    {
        //        command.CommandText = "SELECT * FROM install WHERE Date > '" + DateTime.Now + "'";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        int count = 0;
        //        OleDbDataReader reader = command.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            count++;
        //        }
        //        if (count == 1)
        //        {
        //            return true;
        //        }
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return false;
        //}

        //internal OleDbCommand popuniKorisnike()
        //{
        //    OleDbCommand command = new OleDbCommand();
        //    try
        //    {
        //        connection.Open();
        //        command.Connection = connection;
        //        string query = "SELECT * FROM korisnik";
        //        command.CommandText = query;
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return command;
        //}

        //internal void izmeniKorisnika(string ime, string prezime, string sifra)
        //{
        //    try
        //    {
        //        command.CommandText = "UPDATE korisnik SET Sifra= '" + sifra + "', Prezime= '" + prezime + "'  WHERE Ime= '" + ime + "'";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //        MessageBox.Show("Uspešno ste izvršili izmenu korisnika");
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        //internal void brisiKorisnika(string text)
        //{
        //    try
        //    {
        //        command.CommandText = "DELETE FROM korisnik WHERE Ime = '" + text + "'";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //        MessageBox.Show("Uspešno ste obrisali artikal");
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        internal List<vrstavozila> vrvozila()
        {
            List<vrstavozila> lista_vr = new List<vrstavozila>();
            try
            {
                command.CommandText = "SELECT * FROM vrstavozila ORDER BY ID";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    vrstavozila vr = new vrstavozila();
                    vr.ID = reader["ID"].ToString();
                    vr.NazivVrste = reader["NazivVrste"].ToString();
                    lista_vr.Add(vr);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista_vr;
        }


        internal vozilo reg(string brSasije)
        {
            vozilo v = new vozilo();
            try
            {
                command.CommandText = "SELECT TOP 1 * FROM registracija WHERE BrojSasije = '" + brSasije + "' order by CDate( Left([Datum],2) & '/' & MID([Datum],4,2) & '/' & MID([Datum],7,4)) DESC;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    v.IDBroj = reader["IDBroj"].ToString();
                    v.RegistarskiBroj = reader["RegistarskiBroj"].ToString();
                    v.VrstaVozila = reader["VrstaVozila"].ToString();
                    v.PodvrstaVozila = reader["PodvrstaVozila"].ToString();
                    v.klasa = reader["klasa"].ToString();
                    v.oblikkaroserije = reader["oblikkaroserije"].ToString();
                    v.OblikKaroserijeSkraceno = reader["oblikkaroserijeSkraceno"].ToString();
                    v.specijalno = reader["specijalno"].ToString();
                    v.Marka = reader["Marka"].ToString();
                    v.Model = reader["Model"].ToString();
                    v.Tip = reader["Tip"].ToString();
                    v.GodinaProizvodnje = reader["GodinaProizvodnje"].ToString();
                    v.BrojSasije = reader["BrojSasije"].ToString();
                    v.BrojMotora = reader["BrojMotora"].ToString();
                    v.Boja = reader["Boja"].ToString();
                    v.DatumPrveRegostracije = reader["DatumPrveRegostracije"].ToString();
                    v.Odnos = reader["Odnos"].ToString();
                    v.Snaga = reader["Snaga"].ToString();
                    v.Zapremina = reader["Zapremina"].ToString();
                    v.NajvecaMasa = reader["NajvecaMasa"].ToString();
                    v.Tezina = reader["Tezina"].ToString();
                    v.Nosivost = reader["Nosivost"].ToString();
                    v.Sedenje = reader["Sedenje"].ToString();
                    v.Stajanje = reader["Stajanje"].ToString();
                    v.BrojOsovina = reader["BrojOsovina"].ToString();
                    v.Gorivo = reader["Gorivo"].ToString();
                    v.Homologacija = reader["Homologacija"].ToString();
                    v.ImeVL = reader["ImeVL"].ToString();
                    v.PrezimeVL = reader["PrezimeVL"].ToString();
                    v.MaticniBrojVL = reader["MaticniBrojVL"].ToString();
                    v.BRLKVL = reader["BRLKVL"].ToString();
                    v.BrojPosteVL = reader["BrojPosteVL"].ToString();
                    v.MestoVL = reader["MestoVL"].ToString();
                    v.UlicaVL = reader["UlicaVL"].ToString();
                    v.BrojVL = reader["BrojVL"].ToString();
                    v.TelefonaVL = reader["TelefonaVL"].ToString();
                    v.ZiroRacunVL = reader["ZiroRacunVL"].ToString();
                    v.ImeDL = reader["ImeDL"].ToString();
                    v.PrezimeDL = reader["PrezimeDL"].ToString();
                    v.MaticniBrojDL = reader["MaticniBrojDL"].ToString();
                    v.BRLKDL = reader["BRLKDL"].ToString();
                    v.BrojPosteDL = reader["BrojPosteDL"].ToString();
                    v.MestoDL = reader["MestoDL"].ToString();
                    v.UlicaDL = reader["UlicaDL"].ToString();
                    v.BrojDL = reader["BrojDL"].ToString();
                    v.TelefonaDL = reader["TelefonaDL"].ToString();
                    v.Napomena = reader["Napomena"].ToString();
                    v.Datum = reader["Datum"].ToString();
                    v.VremeOd = reader["VremeOd"].ToString();
                    v.VremeDo = reader["VremeDo"].ToString();
                    v.Kontrolor1 = reader["Kontrolor1"].ToString();
                    v.Kontrolor2 = reader["Kontrolor2"].ToString();
                    v.Predjenikm = reader["Predjenikm"].ToString();
                    v.PoslednjeUverenjeBroj = reader["PoslednjeUverenjeBroj"].ToString();
                    v.PoslednjeUverenjeDatum = reader["PoslednjeUverenjeDatum"].ToString();
                    v.DecimalnaOznaka = reader["DecimalnaOznaka"].ToString();
                    v.OblikKaroserijeSkraceno = reader["OblikKaroserijeSkraceno"].ToString();
                    v.BojaL = reader["BojaL"].ToString(); v.BojaD = reader["BojaD"].ToString();
                    v.BojaD = reader["BojaD"].ToString();
                    v.Ispravnost = reader["Ispravnost"].ToString();
                    v.BrSati = reader["BrSati"].ToString();
                    v.Kuka = reader["Kuka"].ToString();
                    v.OpterecenjeN1 = reader["OpterecenjeN1"].ToString();
                    v.OpterecenjeN2 = reader["OpterecenjeN2"].ToString();
                    v.OpterecenjeN3 = reader["OpterecenjeN3"].ToString();
                    v.OpterecenjeN4 = reader["OpterecenjeN4"].ToString();
                    v.OpterecenjeN5 = reader["OpterecenjeN5"].ToString();
                    v.PneumatikL1 = reader["PneumatikL1"].ToString();
                    v.PneumatikL2 = reader["PneumatikL2"].ToString();
                    v.PneumatikL3 = reader["PneumatikL3"].ToString();
                    v.PneumatikL4 = reader["PneumatikL4"].ToString();
                    v.PneumatikL5 = reader["PneumatikL5"].ToString();
                    v.PneumatikD1 = reader["PneumatikD1"].ToString();
                    v.PneumatikD2 = reader["PneumatikD2"].ToString();
                    v.PneumatikD3 = reader["PneumatikD3"].ToString();
                    v.PneumatikD4 = reader["PneumatikD4"].ToString();
                    v.PneumatikD5 = reader["PneumatikD5"].ToString();
                    v.zaustavljanje = reader["Zaustavljanje"].ToString();
                    v.terensko = reader["terensko"].ToString();
                    v.komentar = reader["komentar"].ToString();
                    v.rejting = reader["rejting"].ToString();
                    v.decplusoznaka = reader["Decplusoznaka"].ToString();
                    v.datumIsteka = reader["datumIsteka"].ToString();

                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return v;
        }

        //public bool proveriKorisnika(string u, string p)
        //{

        //    try
        //    {
        //        command.CommandText = "SELECT * FROM korisnik WHERE Ime='" + u + "' AND Sifra='" + p + "'";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        int count = 0;
        //        OleDbDataReader reader = command.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            count++;
        //        }
        //        if (count == 1)
        //        {
        //            return true;
        //        }
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return false;
        //}



        ////public OleDbCommand popunikorisnike()
        ////{
        ////    connection.Open();
        ////    OleDbCommand command = new OleDbCommand();
        ////    command.Connection = connection;
        ////    string query = "SELECT * FROM korisnik where status = 2";
        ////    command.CommandText = query;
        ////    connection.Close();

        ////    return command;
        ////}

        ////public OleDbCommand popuniodgledane(string usr)
        ////{
        ////    connection.Open();
        ////    OleDbCommand command = new OleDbCommand();
        ////    command.Connection = connection;
        ////    string query = "SELECT Naziv FROM odgledani where Username = '"+usr + "'";
        ////    command.CommandText = query;
        ////    connection.Close();

        ////    return command;
        ////}

        ////public void brisikorisnika(string usr) {
        ////    connection.Open();
        ////    OleDbCommand command = new OleDbCommand();
        ////    command.Connection = connection;
        ////    string query = "DELETE FROM korisnik WHERE Username = " + usr + "";
        ////    command.CommandText = query;

        ////    connection.Close();
        ////}

        //public artikal DodajArtikal(string bar)
        //{
        //    artikal a = new artikal();
        //    try
        //    {
        //        command.CommandText = "SELECT * FROM artikal INNER JOIN barKodovi ON barKodovi.Sifra = artikal.Sifra WHERE barKodovi.Bar = '" + bar + "'";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();


        //        OleDbDataReader reader = command.ExecuteReader();

        //        while (reader.Read())
        //        {

        //            a.sifra = long.Parse(reader["artikal.Sifra"].ToString());
        //            a.opis = reader["Opis"].ToString();
        //            a.cena = Convert.ToDouble(reader["Cena"].ToString());
        //            a.kolicina = Convert.ToDouble(reader["Kolicina"]);
        //            a.nabavna= Convert.ToDouble(reader["Nabavna"].ToString());
        //        }
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return a;
        //}

        //public artikal DodajArtikal(long sifra)
        //{
        //    artikal a = new artikal();
        //    try
        //    {
        //        command.CommandText = "SELECT * FROM artikal WHERE Sifra = " + sifra + ";";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();


        //        OleDbDataReader reader = command.ExecuteReader();

        //        while (reader.Read())
        //        {

        //            a.opis = reader["Opis"].ToString();
        //            a.cena = Convert.ToDouble(reader["Cena"].ToString());
        //            a.kolicina = Convert.ToDouble(reader["Kolicina"]);
        //            a.nabavna = Convert.ToDouble(reader["Nabavna"].ToString());
        //        }
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return a;
        //}



        //public OleDbCommand popuniaritkle()
        //{
        //    OleDbCommand command = new OleDbCommand();
        //    try
        //    {
        //        connection.Open();
        //        command.Connection = connection;
        //        string query = "SELECT * FROM artikal ORDER BY Sifra";
        //        command.CommandText = query;
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return command;
        //}

        //public OleDbCommand popuniBarKodove(long s)
        //{
        //    OleDbCommand command = new OleDbCommand();
        //    try
        //    {
        //        connection.Open();
        //        command.Connection = connection;
        //        string query = "SELECT Bar FROM barKodovi WHERE Sifra = " + s;
        //        command.CommandText = query;
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return command;
        //}

        //public void unesiBar(string b, long s)
        //{
        //    try
        //    {
        //        command.CommandText = "INSERT INTO barKodovi VALUES (" + s + ",'" + b + "')";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //        MessageBox.Show("Uspešno ste dodali novi bar kod");
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        //public void UnesiArtikal(artikal k)
        //{
        //    try
        //    {
        //        command.CommandText = "INSERT INTO artikal VALUES (" + k.sifra + ",'" + k.opis + "'," + k.cena + "," + k.kolicina + "," + k.nabavna + ")";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //        MessageBox.Show("Uspešno ste dodali novi artikal");
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        //public void izmeniArtikal(artikal a)
        //{
        //    try
        //    {
        //        command.CommandText = "UPDATE artikal SET Opis= '" + a.opis + "', Cena = " + a.cena + ", Kolicina = " + a.kolicina + ", Nabavna = " + a.nabavna + " WHERE Sifra= " + a.sifra + "";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //        MessageBox.Show("Uspešno ste izvršili izmenu artikla");
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        //public void brisiArtikal(long k)
        //{

        //    try
        //    {
        //        command.CommandText = "DELETE FROM artikal WHERE Sifra = " + k + "";
        //        command.CommandType = CommandType.Text;

        //        connection.Open();

        //        command.ExecuteNonQuery();

        //        command.CommandText = "DELETE FROM barKodovi WHERE Sifra = " + k + "";
        //        command.CommandType = CommandType.Text;
        //        command.ExecuteNonQuery();

        //        connection.Close();
        //        MessageBox.Show("Uspešno ste obrisali artikal");
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}


        //public int proveri(long a, string datum)
        //{
        //    int count = 0;
        //    try
        //    {
        //        command.CommandText = "SELECT * FROM prodaja WHERE Sifra= " + a + " AND Datum = '" + datum + "'";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();


        //        OleDbDataReader reader = command.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            count++;
        //        }
        //        connection.Close();

        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return count;
        //}

        //public void unesiprodaju(artikal x, string k, double iznos)
        //{
        //    try
        //    {
        //        command.CommandText = "INSERT INTO prodaja VALUES (" + x.sifra + ",'" + x.opis + "'," + x.kolicina + "," + x.cena + ",'" + k + "'," + iznos + ")";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        //public prodaja uzmiprodaju(long bar, string datum)
        //{
        //    prodaja p = new prodaja();
        //    p.Sifra = bar;

        //    try
        //    {
        //        command.CommandText = "SELECT * FROM prodaja WHERE Sifra = " + bar + " AND Datum = '" + datum + "'";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        OleDbDataReader reader = command.ExecuteReader();


        //        while (reader.Read())
        //        {
        //            p.Kolicina = Convert.ToInt32(reader["Kolicina"]);
        //            p.Cena = Convert.ToDouble(reader["Cena"]);

        //        }
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return p;

        //}

        //public void izmeniprodaju(prodaja stara, prodaja nova, string datum)
        //{

        //    try
        //    {
        //        command.CommandText = "UPDATE prodaja SET Kolicina =  " + nova.Kolicina + " , Iznos = " + nova.Iznos + " WHERE Sifra = " + stara.Sifra + " AND Datum = '" + datum + "'";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        internal List<string> nadjiKotrolore()
        {
            List<string> listaGrupa = new List<string>();
            try
            {
                command.CommandText = "SELECT * FROM kontrolor";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string a = string.Empty;
                    a = reader["imePrezime"].ToString();

                    listaGrupa.Add(a);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return listaGrupa;
        }

        internal void brisiRegistraciju(string text)
        {
            try
            {
                command.CommandText = "DELETE FROM registracija WHERE IDBroj = ('" + text + "')";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }

        internal void unesiVrstu(int s, string vrsta)
        {
            try
            {
                command.CommandText = "INSERT INTO vrstavozila(ID,NazivVrste) VALUES ( '" + s + "' ,  '" + vrsta + "');";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Uspešno ste izvrsili snimanje");
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }
        internal void unesiKaroseriju(int s, string vrsta)
        {
            try
            {
                command.CommandText = "INSERT INTO karoserije VALUES ( '" + vrsta + "' ,  '" + s.ToString("00") + "');";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Uspešno ste izvrsili snimanje");
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }
        internal void unesipodvrstu(int s, string vrsta)
        {
            try
            {
                command.CommandText = "INSERT INTO podvrste VALUES ('" + s.ToString("00") + "','" + vrsta + "');";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Uspešno ste izvrsili snimanje");
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }

        internal void unesiMarku(string vrsta)
        {
            try
            {
                command.CommandText = "INSERT INTO marke VALUES ('" + vrsta + "');";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("Uspešno ste izvrsili snimanje");
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }

        internal List<karoserije> lista_podVrs(string text)
        {
            List<karoserije> lista = new List<karoserije>();
            try
            {
                command.CommandText = "SELECT * FROM podvrste INNER JOIN vrstavozila ON vrstavozila.ID = podvrste.ID WHERE vrstavozila.NazivVrste = '" + text + "'";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    karoserije k = new karoserije();
                    k.KaroserijaNaziv = reader["NazivPodvrste"].ToString();
                    lista.Add(k);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return lista;
        }

        internal vozilo id(string va)
        {
            vozilo v = new vozilo();
            try
            {
                command.CommandText = "SELECT * FROM registracija WHERE IDBroj = '" + va + "';";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    v.IDBroj = reader["IDBroj"].ToString();
                    v.RegistarskiBroj = reader["RegistarskiBroj"].ToString();
                    v.VrstaVozila = reader["VrstaVozila"].ToString();
                    v.PodvrstaVozila = reader["PodvrstaVozila"].ToString();
                    v.klasa = reader["klasa"].ToString();
                    v.oblikkaroserije = reader["oblikkaroserije"].ToString();
                    v.OblikKaroserijeSkraceno = reader["oblikkaroserijeSkraceno"].ToString();
                    v.specijalno = reader["specijalno"].ToString();
                    v.Marka = reader["Marka"].ToString();
                    v.Model = reader["Model"].ToString();
                    v.Tip = reader["Tip"].ToString();
                    v.GodinaProizvodnje = reader["GodinaProizvodnje"].ToString();
                    v.BrojSasije = reader["BrojSasije"].ToString();
                    v.BrojMotora = reader["BrojMotora"].ToString();
                    v.Boja = reader["Boja"].ToString();
                    v.DatumPrveRegostracije = reader["DatumPrveRegostracije"].ToString();
                    v.Odnos = reader["Odnos"].ToString();
                    v.Snaga = reader["Snaga"].ToString();
                    v.Zapremina = reader["Zapremina"].ToString();
                    v.NajvecaMasa = reader["NajvecaMasa"].ToString();
                    v.Tezina = reader["Tezina"].ToString();
                    v.Nosivost = reader["Nosivost"].ToString();
                    v.Sedenje = reader["Sedenje"].ToString();
                    v.Stajanje = reader["Stajanje"].ToString();
                    v.BrojOsovina = reader["BrojOsovina"].ToString();
                    v.Gorivo = reader["Gorivo"].ToString();
                    v.Homologacija = reader["Homologacija"].ToString();
                    v.ImeVL = reader["ImeVL"].ToString();
                    v.PrezimeVL = reader["PrezimeVL"].ToString();
                    v.MaticniBrojVL = reader["MaticniBrojVL"].ToString();
                    v.BRLKVL = reader["BRLKVL"].ToString();
                    v.BrojPosteVL = reader["BrojPosteVL"].ToString();
                    v.MestoVL = reader["MestoVL"].ToString();
                    v.UlicaVL = reader["UlicaVL"].ToString();
                    v.BrojVL = reader["BrojVL"].ToString();
                    v.TelefonaVL = reader["TelefonaVL"].ToString();
                    v.ZiroRacunVL = reader["ZiroRacunVL"].ToString();
                    v.ImeDL = reader["ImeDL"].ToString();
                    v.PrezimeDL = reader["PrezimeDL"].ToString();
                    v.MaticniBrojDL = reader["MaticniBrojDL"].ToString();
                    v.BRLKDL = reader["BRLKDL"].ToString();
                    v.BrojPosteDL = reader["BrojPosteDL"].ToString();
                    v.MestoDL = reader["MestoDL"].ToString();
                    v.UlicaDL = reader["UlicaDL"].ToString();
                    v.BrojDL = reader["BrojDL"].ToString();
                    v.TelefonaDL = reader["TelefonaDL"].ToString();
                    v.Napomena = reader["Napomena"].ToString();
                    v.Datum = reader["Datum"].ToString();
                    v.VremeOd = reader["VremeOd"].ToString();
                    v.VremeDo = reader["VremeDo"].ToString();
                    v.Kontrolor1 = reader["Kontrolor1"].ToString();
                    v.Kontrolor2 = reader["Kontrolor2"].ToString();
                    v.Predjenikm = reader["Predjenikm"].ToString();
                    v.PoslednjeUverenjeBroj = reader["PoslednjeUverenjeBroj"].ToString();
                    v.PoslednjeUverenjeDatum = reader["PoslednjeUverenjeDatum"].ToString();
                    v.DecimalnaOznaka = reader["DecimalnaOznaka"].ToString();
                    v.OblikKaroserijeSkraceno = reader["OblikKaroserijeSkraceno"].ToString();
                    v.BojaL = reader["BojaL"].ToString(); v.BojaD = reader["BojaD"].ToString();
                    v.BojaD = reader["BojaD"].ToString();
                    v.Ispravnost = reader["Ispravnost"].ToString();
                    v.BrSati = reader["BrSati"].ToString();
                    v.Kuka = reader["Kuka"].ToString();
                    v.OpterecenjeN1 = reader["OpterecenjeN1"].ToString();
                    v.OpterecenjeN2 = reader["OpterecenjeN2"].ToString();
                    v.OpterecenjeN3 = reader["OpterecenjeN3"].ToString();
                    v.OpterecenjeN4 = reader["OpterecenjeN4"].ToString();
                    v.OpterecenjeN5 = reader["OpterecenjeN5"].ToString();
                    v.PneumatikL1 = reader["PneumatikL1"].ToString();
                    v.PneumatikL2 = reader["PneumatikL2"].ToString();
                    v.PneumatikL3 = reader["PneumatikL3"].ToString();
                    v.PneumatikL4 = reader["PneumatikL4"].ToString();
                    v.PneumatikL5 = reader["PneumatikL5"].ToString();
                    v.PneumatikD1 = reader["PneumatikD1"].ToString();
                    v.PneumatikD2 = reader["PneumatikD2"].ToString();
                    v.PneumatikD3 = reader["PneumatikD3"].ToString();
                    v.PneumatikD4 = reader["PneumatikD4"].ToString();
                    v.PneumatikD5 = reader["PneumatikD5"].ToString();
                    v.zaustavljanje = reader["Zaustavljanje"].ToString();
                    v.terensko = reader["terensko"].ToString();
                    v.komentar = reader["komentar"].ToString();
                    v.rejting = reader["rejting"].ToString();
                    v.decplusoznaka = reader["Decplusoznaka"].ToString();
                    v.datumIsteka = reader["datumIsteka"].ToString();
                    v.VlaznostKT = reader["VlaznostKT"].ToString();
                    v.DatumPrveRegostracije = reader["DatumPrveRegostracije"].ToString();
                    v.SerijskiBroj = reader["SerijskiBroj"].ToString();
                    v.BrojDokumenta = reader["BrojDokumenta"].ToString();
                    v.DatumIzdavanjaSaobracajne = reader["DatumIzdavanjaSaobracajne"].ToString();


                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return v;

        }

        internal string nadjiPoslednjiPregled()
        {
            string s = "";
            try
            {
                command.CommandText = "SELECT * FROM brojPregleda;";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {

                    s = reader["broj"].ToString();
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return s;
        }

        internal void uvecajPreglede(string brpregleda)
        {
            try
            {
                command.CommandText = "UPDATE brojPregleda SET broj = '" + brpregleda + "';";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }

        internal void promeniMargine(string levo, string desno)
        {
            try
            {
                command.CommandText = "UPDATE margine SET levo = '" + levo + "', desno = '" + desno + "';";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }
        internal void promeniBrojac(int br)
        {
            try
            {
                command.CommandText = "UPDATE brojacOtpremnica SET broj = " + br + ";";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }
        internal List<neispravnost> popuniNeispravnost(string text)
        {
            List<neispravnost> listaGrupa = new List<neispravnost>();
            DateTime t = DateTime.Now.Date;
            try
            {
                command.CommandText = "SELECT neispravnosti.ID as ID,neispravnosti.Deo as Deo,neispravnosti.Neispravnost as Neispravnost,neispravnostiPoVozilima.Stepen as Stepen FROM neispravnostiPoVozilima INNER JOIN neispravnosti ON neispravnostiPoVozilima.ID = neispravnosti.ID WHERE IDPregleda = '" + text + "'";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    neispravnost a = new neispravnost();
                    a.ID = reader["ID"].ToString();
                    a.Stepen = reader["Stepen"].ToString();
                    a.Deo = reader["Deo"].ToString();
                    a.Neispravnost = reader["Neispravnost"].ToString();
                    listaGrupa.Add(a);
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
            return listaGrupa;
        }

        internal void unesiNeispravnostiPregleda(neispravnost n, string text)
        {
            try
            {
                command.CommandText = "INSERT INTO neispravnostiPoVozilima VALUES ('" + n.ID + "','" + text + "','" + n.Stepen + "')";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }

        internal void brisiNeispravnostiPregleda(string text)
        {
            try
            {
                command.CommandText = "DELETE FROM neispravnostiPoVozilima WHERE IDPregleda = ('" + text + "')";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }

        internal void posaljiPoruku(string broj_otpremnice)
        {
            try
            {
                command.CommandText = "UPDATE registracija SET Poruka = 'Da' WHERE IDBroj = '" + broj_otpremnice + "'";
                command.CommandType = CommandType.Text;
                connection.Open();

                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }

        internal string izracunajPreglede(string jmbg)
        {

            string br = "0";
            try
            {
                command.CommandText = "SELECT count(*) as broj FROM registracija WHERE ((MaticniBrojVL = '" + jmbg + "' AND MaticniBrojDL = '') OR MaticniBrojDL = '" + jmbg + "') AND Ispravnost = 'Da' ";
                command.CommandType = CommandType.Text;
                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    br = reader["broj"].ToString();
                }
                connection.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

                if (connection != null)
                {
                    connection.Close();
                }
                br = "0";
            }
            return br;
        }

        //public OleDbCommand popuniprodaju(string t)
        //{
        //    OleDbCommand command = new OleDbCommand();
        //    try
        //    {
        //        connection.Open();
        //        command.Connection = connection;
        //        string query = "SELECT * FROM prodaja WHERE Datum = '" + t + "'";
        //        command.CommandText = query;
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return command;
        //}


        //public List<artikal> PronadjiNajprodavanije()
        //{
        //    List<artikal> listaGrupa = new List<artikal>();
        //    DateTime t = DateTime.Now.Date;
        //    try
        //    {
        //        command.CommandText = "SELECT * FROM prodaja WHERE Datum = '" + t.ToString("dd.MM.yyyy") + "' ORDER BY Kolicina DESC";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        OleDbDataReader reader = command.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            artikal a = new artikal();
        //            a.opis = reader["Opis"].ToString();
        //            a.cena = double.Parse(reader["Cena"].ToString());
        //            a.kolicina = double.Parse(reader["kolicina"].ToString());
        //            a.sifra = long.Parse(reader["Sifra"].ToString());

        //            listaGrupa.Add(a);
        //        }
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return listaGrupa;
        //}

        //internal List<artikal> PronadjOmiljene()
        //{
        //    List<artikal> listaGrupa = new List<artikal>();
        //    try
        //    {
        //        command.CommandText = "SELECT artikal.* FROM omiljeni INNER JOIN artikal ON artikal.Sifra = omiljeni.Sifra ORDER BY artikal.Sifra";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        OleDbDataReader reader = command.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            artikal a = new artikal();
        //            a.sifra = long.Parse(reader["Sifra"].ToString());
        //            a.opis = reader["Opis"].ToString();
        //            listaGrupa.Add(a);
        //        }
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return listaGrupa;
        //}

        //internal void povecaj_brojac(int broj_optremnice)
        //{
        //    try
        //    {
        //        command.CommandText = "UPDATE brojac SET br_Otpremnice = " + broj_optremnice;
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        //internal void izmeniKolicinuArtikal(artikal novi)
        //{
        //    try
        //    {
        //        command.CommandText = "UPDATE artikal SET Kolicina =  " + novi.kolicina + ", Cena = " + novi.cena + " WHERE Sifra = " + novi.sifra + "";

        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //    }

        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        //internal void UnesiOmeljine(long sifra)
        //{
        //    try
        //    {
        //        command.CommandText = "INSERT INTO omiljeni VALUES (" + sifra + ")";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}


        //internal void BrisiOmiljene(long sifra)
        //{
        //    try
        //    {
        //        command.CommandText = "DELETE FROM omiljeni WHERE Sifra = (" + sifra + ")";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        //internal void BrisiBarKod(string b)
        //{
        //    try
        //    {
        //        command.CommandText = "DELETE FROM barKodovi WHERE Bar = '" + b + "'";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //        MessageBox.Show("Uspešno ste obrisali bar kod");
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}

        //internal List<artikal> NadjiArtiklePoImenu(string text)
        //{
        //    List<artikal> listaGrupa = new List<artikal>();

        //    try
        //    {
        //        command.CommandText = "SELECT * FROM artikal WHERE Opis LIKE '%" + text + "%'";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        OleDbDataReader reader = command.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            artikal a = new artikal();
        //            a.sifra = long.Parse(reader["Sifra"].ToString());
        //            a.opis = reader["Opis"].ToString();
        //            a.cena = double.Parse(reader["Cena"].ToString());
        //            a.kolicina = double.Parse(reader["Kolicina"].ToString());

        //            listaGrupa.Add(a);
        //        }
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //    return listaGrupa;
        //}

        //internal void unesiPoseban(string naziv)
        //{
        //    try
        //    {
        //        command.CommandText = "INSERT INTO posebni(Naziv) VALUES ('" + naziv + "')";
        //        command.CommandType = CommandType.Text;
        //        connection.Open();

        //        command.ExecuteNonQuery();
        //        connection.Close();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        MessageBox.Show(Exc.Message(ex.Errors[0].SQLState));

        //        if (connection != null)
        //        {
        //            connection.Close();
        //        }
        //    }
        //}
    }
}
/*
public OleDbCommand popuniprodaju(string d)
{
    connection.Open();
    OleDbCommand command = new OleDbCommand();
    command.Connection = connection;
    string query = "SELECT * FROM prodaja WHERE Datum = '" + d + "'";
    command.CommandText = query;
    connection.Close();

    return command;
}

public List<string> PopuniGrupe()
{
    List<string> listaGrupa = new List<string>();

    try
    {
        command.CommandText = "SELECT * FROM rabat";
        command.CommandType = CommandType.Text;
        connection.Open();

        OleDbDataReader reader = command.ExecuteReader();

        while (reader.Read())
        {
            string s = reader["Tip"].ToString();

            listaGrupa.Add(s);
        }
        return listaGrupa;
    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}



public void UnesiKupca(kupci k)
{
    try
    {
        command.CommandText = "INSERT INTO kupci (Naziv,Adresa,Postanski_broj,Grad,PIB,RG) VALUES ('" + k.naziv + "','" + k.adresa + "','" + k.postanski_broj +"','" + k.Grad + "','" + k.PIB + "','" + k.RG + "')";
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();
    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}



public void IzmeniKupca(kupci starikupac, kupci novikupac)
{

    {

        try
        {
            command.CommandText = "UPDATE kupci SET Adresa= '" + novikupac.adresa + "', Postanski_broj ='" + novikupac.postanski_broj + "', Grad ='" + novikupac.Grad + "', PIB ='" + novikupac.PIB + "', RG ='" + novikupac.RG + "', Naziv ='" + novikupac.naziv + "' WHERE Naziv='" + starikupac.naziv + "'";
            command.CommandType = CommandType.Text;
            connection.Open();

            command.ExecuteNonQuery();

        }
        catch (Exception) { throw; }
        finally
        {
            if (connection != null)
            {
                connection.Close();
            }
        }
    }
}


public void BrisiKupca(kupci k)
{

    try
    {
        command.CommandText = "DELETE FROM kupci WHERE Naziv= '" + k.naziv + "'";
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();

    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}


public void UnesiArtikal(artikal k)
{
    try
    {
        command.CommandText = "INSERT INTO magacin (Bar,Opis,Cena,Kolicina,Kataloski_kod) VALUES ('" + k.bar + "','" + k.naziv + "'," + k.cena + "," + k.kolicina + ",'" +k.kataloski_kod+"')";
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();
    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}



public void izmeniArtikal(artikal a)
{

    {

        try
        {
            command.CommandText = "UPDATE magacin SET Opis= '" + a.naziv + "', Cena = " + a.cena +  ", Kolicina = "+ a.kolicina +", Kataloski_kod=  '" + a.kataloski_kod+"' WHERE Bar = " + a.bar;
            command.CommandType = CommandType.Text;
            connection.Open();

            command.ExecuteNonQuery();

        }
        catch (Exception) { throw; }
        finally
        {
            if (connection != null)
            {
                connection.Close();
            }
        }
    }
}


public void brisiArtikal(long k)
{

    try
    {
        command.CommandText = "DELETE FROM magacin WHERE Bar= " + k + "";
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();

    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}


public void dodajKolicinu(artikal a)
{

    {

        try
        {
            command.CommandText = "UPDATE magacin SET Kolicina = " + a.kolicina +" WHERE Sifra = " + a.bar;
            command.CommandType = CommandType.Text;
            connection.Open();

            command.ExecuteNonQuery();

        }
        catch (Exception) { throw; }
        finally
        {
            if (connection != null)
            {
                connection.Close();
            }
        }
    }
}






//brojac otpremnica da odredi koji je broj ispred /15
public void povecaj_brojac(int i)
{
    try
    {
        command.CommandText = "UPDATE brojac SET br_Otpremnice = " + i ;
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();
    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}

public int brojac_otpremnice()
{
    try
    {
        command.CommandText = "SELECT * FROM brojac";
        command.CommandType = CommandType.Text;
        connection.Open();

        OleDbDataReader reader = command.ExecuteReader();

        while (reader.Read())
        {

            return Convert.ToInt32(reader["br_Otpremnice"].ToString());

        }
    }
    catch (Exception) { throw; }
    finally
    {

        if (connection != null)
        {
            connection.Close();
        }
    }
    return 0;
}

public void povecaj_fiskalni(int i)
{
    try
    {
        command.CommandText = "UPDATE fiskalni SET bi = " + i;
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();
    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}

public int brojac_fiskalnog()
{
    try
    {
        command.CommandText = "SELECT * FROM fiskalni";
        command.CommandType = CommandType.Text;
        connection.Open();

        OleDbDataReader reader = command.ExecuteReader();

        while (reader.Read())
        {

            return Convert.ToInt32(reader["bi"].ToString());

        }
    }
    catch (Exception) { throw; }
    finally
    {

        if (connection != null)
        {
            connection.Close();
        }
    }
    return 0;
}

public int proveri(Int32 a,string datum)
{

    try
    {
        command.CommandText = "SELECT * FROM prodaja WHERE Bar= " + a + " AND Datum = '" + datum + "'";
        command.CommandType = CommandType.Text;
        connection.Open();

        int count = 0;
        OleDbDataReader reader = command.ExecuteReader();

        while (reader.Read())
        {

            count++;
        }

        return count;


    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}


public double izracunaj_dug(dugovanje a)
{
    double s = 0;

    try
    {
        command.CommandText = "SELECT * FROM uplate_kupaca WHERE Br_otpremnice = '" +   "'";
        command.CommandType = CommandType.Text;
        connection.Open();

        OleDbDataReader reader = command.ExecuteReader();


        while (reader.Read())
        {
            uplate f = new uplate();

            s += Convert.ToDouble(reader["Iznos"]);

        }
        return s;
    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}




public void unesiprodaju(artikal x,string k, double iznos)
{
    try
    {
        command.CommandText = "INSERT INTO prodaja VALUES (" + x.bar + ",'" + x.naziv+ "'," + x.kolicina + "," + x.cena + ",'"+ k + "'," + iznos +")";
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();
    }
    catch (Exception) { }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}

public void izmeniprodaju(prodaja stara, prodaja nova,string datum)
{

    try
    {
        command.CommandText = "UPDATE prodaja SET Kolicina =  " + nova.Kolicina + " , Iznos = " + nova.Iznos + " WHERE Bar = " + stara.Bar + " AND Datum = '" + datum+ "'";
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();

    }
    catch (Exception) {  }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}

public void izmenimagacin(artikal novi)
{

    try
    {
        command.CommandText = "UPDATE magacin SET Kolicina =  " + novi.kolicina + ", Cena = "+ novi.cena +", Kataloski_kod = '" + novi.kataloski_kod + "' WHERE Bar = " + novi.bar;
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();

    }
    catch (Exception) { }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }

}






public prodaja uzmiprodaju(int bar,string datum)
{
    prodaja p = new prodaja();
    p.Bar = bar;

    try
    {
        command.CommandText = "SELECT * FROM prodaja WHERE Bar = " + bar + " AND Datum = '" + datum + "'";
        command.CommandType = CommandType.Text;
        connection.Open();

        OleDbDataReader reader = command.ExecuteReader();


        while (reader.Read())
        {


            p.Kolicina = Convert.ToInt32(reader["Kolicina"]);
            p.Cena = Convert.ToDouble(reader["Cena"]);

        }
        return p;
    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}




public OleDbCommand popunidugovanja(kupci k)
{
    connection.Open();
    OleDbCommand command = new OleDbCommand();
    command.Connection = connection;
    string query = "SELECT * FROM dugovanja WHERE Naziv = '" + k.naziv + "'";
    command.CommandText = query;
    connection.Close();

    return command;
}

public void UnesiZaduzenje(dugovanje k)
{
    try
    {
        command.CommandText = "INSERT INTO dugovanja (Naziv,Broj_Dokumenta,Datum,Nalog,Duguje,Potrazuje,Saldo) VALUES ('" + k.naziv + "','" + k.broj_dokumenta + "','" + k.datum + "'," + k.nalog + "," + k.duguje + "," + k.potrazuje + "," + k.saldo + ")";
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();
    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }

}



public void brisidugovanje(int id)
{
    try
    {
        //command.CommandText = "DELETE FROM dugovanja WHERE  Naziv = '" + p + "' , Broj_dokumenta = '" + p_2 + "' , Potrazuje = "+ d;
        command.CommandText = "DELETE FROM dugovanja WHERE  ID = " + id;
        command.CommandType = CommandType.Text;
        connection.Open();

        command.ExecuteNonQuery();

    }
    catch (Exception) { throw; }
    finally
    {
        if (connection != null)
        {
            connection.Close();
        }
    }
}

public OleDbCommand popuniaritkle(string p)
{
    connection.Open();
    OleDbCommand command = new OleDbCommand();
    command.Connection = connection;
    string query = "SELECT * FROM magacin WHERE Opis LIKE '%" + p +"%' ORDER BY Bar DESC";
    command.CommandText = query;
    connection.Close();

    return command;
}


public OleDbCommand popuniaritklepokodu(string p)
{
connection.Open();
OleDbCommand command = new OleDbCommand();
command.Connection = connection;
string query = "SELECT * FROM magacin WHERE Kataloski_kod LIKE '%" + p + "%' ORDER BY Bar DESC";
command.CommandText = query;
connection.Close();

return command;
}
public OleDbCommand popuniaritkle(int p)
{
connection.Open();
OleDbCommand command = new OleDbCommand();
command.Connection = connection;
string query = "SELECT * FROM magacin WHERE Bar LIKE '%" + p + "%' ORDER BY Bar DESC";
command.CommandText = query;
connection.Close();

return command;
}

public void unesiprodajucrno(artikal x, string k, double iznos)
{
try
{
command.CommandText = "INSERT INTO prodaja_pom VALUES (" + x.bar + ",'" + x.naziv + "'," + x.kolicina + "," + x.cena + ",'" + k + "'," + iznos + ")";
command.CommandType = CommandType.Text;
connection.Open();

command.ExecuteNonQuery();
}
catch (Exception) { }
finally
{
if (connection != null)
{
    connection.Close();
}
}
}

public void izmeniprodajucrno(prodaja stara, prodaja nova, string datum)
{

try
{
command.CommandText = "UPDATE prodaja_pom SET Kolicina =  " + nova.Kolicina + " , Iznos = " + nova.Iznos + " WHERE Bar = " + stara.Bar + " AND Datum = '" + datum + "'";
command.CommandType = CommandType.Text;
connection.Open();

command.ExecuteNonQuery();

}
catch (Exception) { }
finally
{
if (connection != null)
{
    connection.Close();
}
}
}

public OleDbCommand popuniprodajucrno(string t)
{
connection.Open();
OleDbCommand command = new OleDbCommand();
command.Connection = connection;
string query = "SELECT * FROM prodaja_pom WHERE Datum = '" + t + "'";
command.CommandText = query;
connection.Close();

return command;
}

public int provericrno(int sifra, string pom)
{
try
{
command.CommandText = "SELECT * FROM prodaja_pom WHERE Bar= " + sifra + " AND Datum = '" + pom + "'";
command.CommandType = CommandType.Text;
connection.Open();

int count = 0;
OleDbDataReader reader = command.ExecuteReader();

while (reader.Read())
{

    count++;
}

return count;


}
catch (Exception) { throw; }
finally
{
if (connection != null)
{
    connection.Close();
}
}
}

public prodaja uzmiprodajucrno(int sifra, string pom)
{
prodaja p = new prodaja();
p.Bar = sifra;

try
{
command.CommandText = "SELECT * FROM prodaja_pom WHERE Bar = " + sifra + " AND Datum = '" + pom + "'";
command.CommandType = CommandType.Text;
connection.Open();

OleDbDataReader reader = command.ExecuteReader();


while (reader.Read())
{


    p.Kolicina = Convert.ToInt32(reader["Kolicina"]);
    p.Cena = Convert.ToDouble(reader["Cena"]);

}
return p;
}
catch (Exception) { throw; }
finally
{
if (connection != null)
{
    connection.Close();
}
}
}

public OleDbCommand popunidugovanjaDobavljaca(kupci k)
{
connection.Open();
OleDbCommand command = new OleDbCommand();
command.Connection = connection;
string query = "SELECT * FROM dugovanja_dobavljacima WHERE Naziv = '" + k.naziv + "'";
command.CommandText = query;
connection.Close();

return command;
}

public void UnesiZaduzenjeDobavljacu(zaduzenja k)
{
try
{
command.CommandText = "INSERT INTO dugovanja_dobavljacima (Naziv,Broj_Dokumenta,Datum,Nalog,Duguje,Potrazuje,Saldo) VALUES ('" + k.naziv + "','" + k.broj_dokumenta + "','" + k.datum + "','" + k.nalog + "'," + k.duguje + "," + k.potrazuje + "," + k.saldo + ")";
command.CommandType = CommandType.Text;
connection.Open();

command.ExecuteNonQuery();
}
catch (Exception) {  }
finally
{
if (connection != null)
{
    connection.Close();
}
}
}

public void brisidugovanjeDobavljacu(int id)
{
try
{
//command.CommandText = "DELETE FROM dugovanja WHERE  Naziv = '" + p + "' , Broj_dokumenta = '" + p_2 + "' , Potrazuje = "+ d;
command.CommandText = "DELETE FROM dugovanja_dobavljacima WHERE  ID = " + id;
command.CommandType = CommandType.Text;
connection.Open();

command.ExecuteNonQuery();

}
catch (Exception) { }
finally
{
if (connection != null)
{
    connection.Close();
}
}
}

public void IzmeniZaduzenjeDobavljacu(zaduzenja z)
{
try
{
command.CommandText = "UPDATE dugovanja_dobavljacima SET Naziv = '" + z.naziv + "' ,Broj_Dokumenta= '" + z.broj_dokumenta + "' ,Datum= '" + z.datum + "' ,Nalog= '" + z.nalog + "' ,Duguje=" + z.duguje + " ,Potrazuje= " + z.potrazuje + " ,Saldo= " + z.saldo + " WHERE ID=" + z.ID;
command.CommandType = CommandType.Text;
connection.Open();

command.ExecuteNonQuery();
}
catch (Exception) { }
finally
{
if (connection != null)
{
    connection.Close();
}
}
}

public List<string> nadjislicne(string p)
{
List<string> listaSlicnih = new List<string>();

try
{
command.CommandText = "SELECT * FROM slicni WHERE Kataloski_kod_1 = '" + p + "'";
command.CommandType = CommandType.Text;
connection.Open();

OleDbDataReader reader = command.ExecuteReader();

while (reader.Read())
{
    string f = reader["Kataloski_kod_2"].ToString();

    listaSlicnih.Add(f);
}
return listaSlicnih;
}
catch (Exception) { throw; }
finally
{
if (connection != null)
{
    connection.Close();
}
}
}



public void dodajSlicne(string p, string p_2)
{
try
{
command.CommandText = "INSERT INTO slicni (Kataloski_kod_1,Kataloski_kod_2) VALUES ('" +p + "','" +p_2 + "')";
command.CommandType = CommandType.Text;
connection.Open();

command.ExecuteNonQuery();
}
catch (Exception) { }
finally
{
if (connection != null)
{
    connection.Close();
}
}
}

public void BrisiSlicne(string p, string p_2)
{
try
{
command.CommandText = "DELETE * FROM slicni WHERE Kataloski_kod_1 = '" + p + "' AND Kataloski_kod_2 = '" + p_2 + "'";
command.CommandType = CommandType.Text;
connection.Open();

command.ExecuteNonQuery();
}
catch (Exception) { }
finally
{
if (connection != null)
{
    connection.Close();
}
}
}

public void unesinapomenu(artikal a)
{
{

try
{
    command.CommandText = "UPDATE magacin SET Napomena= '" + a.napomena + "' WHERE Bar = " + a.bar;
    command.CommandType = CommandType.Text;
    connection.Open();

    command.ExecuteNonQuery();

}
catch (Exception) { throw; }
finally
{
    if (connection != null)
    {
        connection.Close();
    }
}
}
}
}
}

*/
