using System.Globalization;
using System.Windows.Forms.Design;

namespace FinancialAdministrator
{
    public partial class financialAdministrator : Form
    {
        CsvReader csvReader { get; set; }

        string[] readFile { get; set; }

        int administrationYear { get; set; }
        public financialAdministrator()
        {
            InitializeComponent();
            progressBar1.Visible = false;
        }

        private void readFileButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string fileName = openFileDialog1.FileName;
            try
            {
                readFile = File.ReadAllLines(fileName);
                //todo handle files that are not csv
                administrationYear = Convert.ToInt16(Convert.ToDateTime(readFile[1].Split(';')[4], new CultureInfo("nl-BE")).Year);
                filePreview.Text = "De kolommen zijn: \n\n"; 
                foreach (string column in readFile[0].Split(';'))
                {
                    filePreview.Text += column + "\n";
                }
                    
                csvReader = new CsvReader(readFile);
            } catch (Exception)
            {
                filePreview.Text = "Het gekozen bestand is in gebruik!\nSluit het en probeer het opnieuw.";
            }           
        }

        private void generateXcelButton_Click(object sender, EventArgs e)
        {
            if (readFile == null)
            {
                filePreview.Text = "Selecteer eerst een valide CSV bestand";
            }
            else
            {
                saveFileDialog1.ShowDialog();
                string fileName = saveFileDialog1.FileName + ".xlsx";
                filePreview.Text = "Schrijven naar: " + fileName;
                List<TransactieModel> result = csvReader.parseFile();
                ExcelWriter writer = new ExcelWriter(fileName, result, administrationYear, progressBar1);
                writer.writeFile();
            }
            
        }
    }
}
