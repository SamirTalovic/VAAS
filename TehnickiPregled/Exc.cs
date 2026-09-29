using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TehnickiPregled
{
    public static class Exc
    {
        public static string Message(string NumberOfError)
        {
            try
            {
                int n = Convert.ToInt32(NumberOfError);

                switch (n)
                {
                    case 3200:
                        return "Ne može se izvršiti brisanje jer je podatak vezan za neku drugu tabelu. Kontaktirajte administratora aplikacije!";
                        break; ;
                    case 3126:
                        return "";
                        break;
                    case 3346:
                        return "Broj unetih i destinacionih argumenata nije isti. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3000:
                        return "Došlo je do greške u upitu. Kontakirajte administratora aplikacije!";
                        break;
                    case 2420:
                        return "Sintaksna greška u zapisu broja. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2421:
                        return "Sintaksna greška u zapisu datuma. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2422:
                        return "Sintaksna greška u zapisu stringa. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2423:
                        return "Nedozvoljeno korišćenje ., ! ili (). Kontaktirajte administratora aplikacije!";
                        break;
                    case 2424:
                        return "Nepoznato ime. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2427:
                        return "Objektu nije dodeljenja vrednost. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2428:
                        return "Neispravno korišćenje argumenata na upitu. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2431:
                        return "Sintaksna greška u upitu nedostaje operator. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2432:
                        return "Sintaksna greška: nedostaje zarez u upitu. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2433:
                        return "Sintaksna greška u upitu. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2435:
                        return "Sintaksna greška u upitu postoji ) viška. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2436:
                        return "Sintaksna greška u upitu nedostaje ) ili ]. Kontaktirajte administratora aplikacije!";
                        break;
                    case 2447:
                        return "Nedozvoljeno korišćenje ., ! ili (). Kontaktirajte administratora aplikacije!";
                        break;
                    case 2448:
                        return "Nije moguće postaviti određenu vrednost u upitu. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3001:
                        return "Neispravni podaci za unos. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3011:
                        return "U bazi nije moguće naći traženi objekat iz upita. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3017:
                        return "Polje je previše dugačko. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3022:
                        return "Nije moguće uneti zapis, jer već postoji u bazi podataka. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3047:
                        return "Zapis je previše dug. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3060:
                        return "Pogrešan tip podataka za unos, proverite polja. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3078:
                        return "Sintaksna greška, u bazi ne postoji tabela sa traženim nazivom. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3106:
                        return "Nedostaje polja nad kojim se vrši izmena. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3130:
                        return "Sintaksna greška u queryju za brisanje. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3131:
                        return "Sintaksna greška u queryiju za selektovanje. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3134:
                        return "Sintaksna greška u queryiju za unos podataka. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3135:
                        return "Sintaksna greška u JOIN queryiju. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3144:
                        return "Sintaksna greška u queryiju za izmenu. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3145:
                        return "Sintaksna greška u SELECT WHERE queryju. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3283:
                        return "Sintaksna greška primarni ključ već postoji. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3309:
                        return "Sintaksna greška: vrednost polja je predugačka. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3314:
                        return "Niste popunili sva polja koja se zahtevaju. Kontaktirajte administratora aplikacije!";
                        break;
                    case 3201:
                        return "Ne možete menjati niti brisati polje jer je povezano sa nekom drugom tabelom. Kontaktirajte administratora aplikacije!";
                        break;
                    default:
                        return "Došlo je do greške. Kontaktirajte administratora aplikacije!";
                        break;
                }
            }
            catch { }
            return "";
        }
    }
}
