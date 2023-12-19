using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows.Forms.Design;
using System.Resources;
using System.Reflection;

namespace FinancialAdministrator
{
    public partial class FinancialAdministrator : Form
    {
        CsvReader? CsvReader { get; set; }

        string[]? ReadFile { get; set; }

        int AdministrationYear { get; set; }

        ResourceManager ResourceManager { get; set; }
        public FinancialAdministrator()
        {
            InitializeComponent();
            progressBar1.Visible = false;
            ResourceManager = new ResourceManager("FinancialAdministratorLight.Resources.Strings", Assembly.GetExecutingAssembly());
            Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("nl-BE");
            filePreview.Text = ResourceManager.GetString("NoFile");
        }

        private void ReadFileButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            FileInfo file = new FileInfo(openFileDialog1.FileName);
           
            if (file.Extension != ".csv")
            {
                filePreview.Text = ResourceManager.GetString("CSVError1");
            }
            else
            {
                try
                {
                    ReadFile = File.ReadAllLines(file.FullName);

                    if (ReadFile != null && ReadFile.Length > 1)
                    {
                        var splitResult = ReadFile[1].Split(';');
                        if (splitResult != null && splitResult.Length > 4)
                        {
                            try
                            {
                                AdministrationYear = Convert.ToInt16(Convert.ToDateTime(splitResult[4], new CultureInfo("nl-BE")).Year);
                                filePreview.Text = ResourceManager.GetString("Columns");

                                var secondSplitResult = ReadFile[0].Split(';');
                                if (secondSplitResult != null && secondSplitResult.Length > 10)
                                {

                                    foreach (string column in ReadFile[0].Split(';'))
                                    {
                                        filePreview.Text += column + "\n";
                                    }

                                    CsvReader = new CsvReader(ReadFile);
                                }
                                else
                                {
                                    filePreview.Text += ResourceManager.GetString("NonValid");
                                }
                            }
                            catch (FormatException)
                            {
                                filePreview.Text = ResourceManager.GetString("CSVError2");
                            }
                            catch (Exception)
                            {
                                filePreview.Text = ResourceManager.GetString("SomethingWrong");
                            }
                        }
                        else
                        {
                            filePreview.Text = ResourceManager.GetString("CSVError2");
                        }
                    }
                    else
                    {
                        filePreview.Text = ResourceManager.GetString("CSVError2");
                    }
                }
                catch (IOException)
                {
                    filePreview.Text = ResourceManager.GetString("FileInUse");
                }
                catch (Exception)
                {
                    filePreview.Text = ResourceManager.GetString("SomethingWrong");
                }
            }
        }

        private async void GenerateXcelButton_Click(object sender, EventArgs e)
        {
            if (ReadFile == null)
            {
                filePreview.Text = ResourceManager.GetString("CSVError3");
            }
            else
            {
                saveFileDialog1.ShowDialog();
                try
                {
                    FileInfo file = new(saveFileDialog1.FileName);
                    filePreview.Text = ResourceManager.GetString("WritingTo") + file.FullName + (file.Extension == ".xlsx" ? "" : ".xlsx");
                    List<TransactieModel> result = await CsvReader.parseFileAsync();
                    ExcelWriter excelWriter = new(file, result, AdministrationYear, progressBar1, filePreview);
                    await excelWriter.WriteFileAsync();
                }
                catch (ArgumentException)
                {
                    filePreview.Text = ResourceManager.GetString("NoFileOutput");
                }
                catch (Exception)
                {
                    filePreview.Text = ResourceManager.GetString("SomethingWrong");
                }
            }
            
        }
    }
}
