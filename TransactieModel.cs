using System.Globalization;

namespace FinancialAdministrator
{
    public class TransactieModel
    {
        public DateTime Boekingsdatum { get; set; }

        public double Bedrag { get; set; }

        public string Tegenrekening { get; set; }

        public string Omschrijving { get; set; }

        public string Detail { get; set; }

        public string Categorie { get; set; }

        public string[] boodschappen = { "jumbo", "carrefour", "colruyt", "aldi", "albert heijn", "delhaize", "co&go" };

        public string[] giften = { "mercy ships", "rode kruis" };

        public string[] auto = { "tinq", "q8 easy", "texaco", "shell", "parking" };

        public string[] ziekte = { "uz leuven", "goed farma", "apotheek", "facturenbureau k.u." };

        public string[] verzekeringen = { "verzekeringen", "verzekeringspremie", "zorgkas" };

        public string[] belastingvermindering = { "heyo vzw", "sporty creatief vzw", "tofsport", "katholiek onderwijs" };
        
        public TransactieModel(string boekingsdatum, string bedrag, string tegenrekening, string omschrijving, string detail)
        {
            Tegenrekening = tegenrekening;
            Boekingsdatum = Convert.ToDateTime(boekingsdatum);
            Bedrag = Convert.ToDouble(bedrag);
            Omschrijving = omschrijving;
            Detail = detail;
            Categorie = "None";

            if (giften.Any(omschrijving.ToLower().Contains) && Bedrag < 0)
            {
                Categorie = "Donations";
            }
            if (auto.Any(omschrijving.ToLower().Contains) && Bedrag < 0)
            {
                Categorie = "Car";
            }
            if (ziekte.Any(omschrijving.ToLower().Contains) && Bedrag < 0)
            {
                Categorie = "Sickness";
            }
            if (boodschappen.Any(omschrijving.ToLower().Contains) && Bedrag < 0)
            {
                Categorie = "Shoppings";
            }
            if (verzekeringen.Any(omschrijving.ToLower().Contains) && Bedrag < 0)
            {
                Categorie = "Insurance";
            }
            if (belastingvermindering.Any(omschrijving.ToLower().Contains) && Bedrag < 0)
            {
                Categorie = "TaxReduction";
            }
            if (tegenrekening == "NL26ASNB8888888888") // change to valid savings account
            {
                Categorie = "ToSavingsAccount";
            }
            if (Bedrag > 0)
            {
                Categorie = "Deposit";
            }
        }
    }
}
