using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections;
using System.Diagnostics;
using System.Drawing.Text;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Resources;
using System.Windows.Forms.Design;

namespace FinancialAdministrator
{
    internal class ExcelWriter
    {
        public FileInfo File { get; set; }

        public List<TransactieModel> Data { get; set; }

        public int AdministrationYear { get; set; }

        public Hashtable NumberOfRowsList = [];

        readonly ProgressBar progressBar1;

        readonly RichTextBox FilePreview;
        ResourceManager ResourceManager { get; set; }

        readonly string[,]? Months;
        

        public ExcelWriter(FileInfo file, List<TransactieModel> result, int administrationYear, ProgressBar progressBar1, RichTextBox filePreview)
        {
            File = file;
            Data = result;
            AdministrationYear = administrationYear;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            this.progressBar1 = progressBar1;
            FilePreview = filePreview;

            ResourceManager = new ResourceManager("FinancialAdministratorLight.Resources.Strings", Assembly.GetExecutingAssembly());
            Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("nl-BE");
            Months = new string[,] {
                                    { "January", "31" },
                                    { "February", (AdministrationYear % 4 == 0) ? "29" : "28" },
                                    { "March", "31" },
                                    { "April", "30" },
                                    { "May", "31" },
                                    { "June", "30" },
                                    { "July", "31" },
                                    { "August", "31" },
                                    { "September", "30" },
                                    { "October", "31" },
                                    { "November", "30" },
                                    { "December", "31" }
                                    };            
        }   

        public async Task WriteFileAsync()
        {
            DeleteIfExists();
            string fileName = File.FullName;

            if (File.Extension != ".xlsx")
            {
                fileName += ".xlsx";
            }

            using (var package = new ExcelPackage(fileName))
            {
                progressBar1.Visible = true;
                for (int i = 0; i < Months.GetLength(0); i++)
                {
                    try
                    {
                        await Task.Run(() => CreateMonthSheet(Months[i, 0], i + 1, Convert.ToInt16(Months[i, 1]), package));
                    }
                    catch (InvalidOperationException)
                    {
                        FilePreview.Text = ResourceManager.GetString("FileInUse");
                    }
                    catch (Exception)
                    {
                        FilePreview.Text = ResourceManager.GetString("SomethingWrong");
                    }
                    progressBar1.Value =  100/12 * i;
                }

                await package.SaveAsync();
                progressBar1.Value = 100;

                if (Type.GetTypeFromProgID("Excel.Application", false) != null)
                {
                    System.Diagnostics.Process.Start(new ProcessStartInfo { FileName = File.FullName, UseShellExecute = true });
                }
              
                progressBar1.Visible = false;
            }
       }

        private void CreateMonthSheet(string monthName, int monthNumber, int endDay, ExcelPackage package)
        {
            try {
                var workSheet = package.Workbook.Worksheets.Add(ResourceManager.GetString(monthName));
            
                var monthData = Data.Where(transactie => 
                    transactie.Boekingsdatum >= new DateTime(AdministrationYear, monthNumber, 01)
                    && transactie.Boekingsdatum <= new DateTime(AdministrationYear, monthNumber, endDay))
                    .OrderBy(transactie => transactie.Boekingsdatum);

                int numberOfRows = 0;

                for (int i = 0; i < monthData.Count(); i++)
                {
                    switch (monthData.ElementAt(i).Categorie) {
                        case "Shoppings":
                            workSheet.Cells["H" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                            break;
                        case "Deposit":
                            workSheet.Cells["C" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                            break;
                        case "Donations":
                            workSheet.Cells["G" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                            break;
                        case "Car":
                            workSheet.Cells["I" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                            break;
                        case "Sickness":
                            workSheet.Cells["F" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                            break;
                        case "Insurance":
                            workSheet.Cells["E" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                            break;
                        case "TaxReduction":
                            workSheet.Cells["L" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                            break;
                        case "ToSavingsAccount":
                            workSheet.Cells["B" + (i + 12)].Value = Math.Abs(monthData.ElementAt(i).Bedrag);
                            workSheet.Cells["K" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                            break;
                        default:
                            workSheet.Cells["K" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                            break;
                    }
                    workSheet.Cells["M" + (i + 12)].Value = monthData.ElementAt(i).Tegenrekening;
                    workSheet.Cells["N" + (i + 12)].Value = monthData.ElementAt(i).Omschrijving;
                    workSheet.Cells["O" + (i + 12)].Value = monthData.ElementAt(i).Boekingsdatum;
                    workSheet.Cells["P" + (i + 12)].Value = monthData.ElementAt(i).Detail;
                    numberOfRows = (i + 12);
                }

                ApplyTemplate(workSheet, numberOfRows, monthNumber);
            }
            catch (MissingManifestResourceException)
            {
                FilePreview.Text = $"Months are not defined in resource for nl-BE";
            }
            catch (Exception)
            {
                FilePreview.Text = ResourceManager.GetString("SomethingWrong");
            }
        }

        private void ApplyTemplate(ExcelWorksheet workSheet, int numberOfRows, int monthNumber)
        {
            workSheet.Column(15).Style.Numberformat.Format = "dd-MM-yyyy";
            workSheet.Cells["A1"].Value = ResourceManager.GetString("FinAdmin") + " " + AdministrationYear;
            workSheet.Cells["A1:B1"].Merge = true;
            workSheet.Row(1).Style.Font.Size = 14;

            workSheet.Cells["B9"].Value = ResourceManager.GetString("IncomeTo");
            workSheet.Cells["B9:C9"].Merge = true;
            workSheet.Cells["B9:C9"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["E9:P9"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["A10:P10"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            workSheet.Cells["A11:P11"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["A11:P11"].Style.Fill.SetBackground(Color.LightGray);

            workSheet.Cells["D9"].Value = ResourceManager.GetString("ExpensesOf");
            workSheet.Cells["E9"].Value = ResourceManager.GetString("CheckingAccountExpenses");
            workSheet.Cells["E9:L9"].Merge = true;
            workSheet.Cells["A10"].Value = ResourceManager.GetString("ManualDescription");
            workSheet.Cells["A11"].Value = ResourceManager.GetString("MainAccount");
            workSheet.Cells["B10"].Value = ResourceManager.GetString("SavingsAcc");
            workSheet.Cells["C10"].Value = ResourceManager.GetString("CheckingAcc");
            workSheet.Cells["D10"].Value = ResourceManager.GetString("SavingsAccount");
            workSheet.Cells["E10"].Value = ResourceManager.GetString("Insurance");
            workSheet.Cells["F10"].Value = ResourceManager.GetString("Sickness");
            workSheet.Cells["G10"].Value = ResourceManager.GetString("Donations");
            workSheet.Cells["H10"].Value = ResourceManager.GetString("Shoppings");
            workSheet.Cells["I10"].Value = ResourceManager.GetString("Car");
            workSheet.Cells["J10"].Value = ResourceManager.GetString("Work");
            workSheet.Cells["K10"].Value = ResourceManager.GetString("Other");
            workSheet.Cells["L10"].Value = ResourceManager.GetString("TaxReduction");
            workSheet.Cells["M10"].Value = ResourceManager.GetString("OffsetAccount");
            workSheet.Cells["N10"].Value = ResourceManager.GetString("Description");
            workSheet.Cells["O10"].Value = ResourceManager.GetString("Date");
            workSheet.Cells["P10"].Value = ResourceManager.GetString("Details");
            workSheet.Cells["A9:P10"].Style.Font.Bold = true;

            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

            workSheet.Cells["B9:B" + numberOfRows].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["C9:C" + numberOfRows].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["D9:D" + numberOfRows].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = ResourceManager.GetString("Subtotals");
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;

            workSheet.Cells["B" + numberOfRows].Formula = "SUM(B12:B" + (numberOfRows - 1) + ")";
            workSheet.Cells["B" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["C" + numberOfRows].Formula = "SUM(C12:C" + (numberOfRows - 1) + ")";
            workSheet.Cells["C" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["D" + numberOfRows].Formula = "SUM(D12:D" + (numberOfRows - 1) + ")";
            workSheet.Cells["D" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["E" + numberOfRows].Formula = "SUM(E12:E" + (numberOfRows - 1) + ")";
            workSheet.Cells["E" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["F" + numberOfRows].Formula = "SUM(F12:F" + (numberOfRows - 1) + ")";
            workSheet.Cells["F" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["G" + numberOfRows].Formula = "SUM(G12:G" + (numberOfRows - 1) + ")";
            workSheet.Cells["G" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["H" + numberOfRows].Formula = "SUM(H12:H" + (numberOfRows - 1) + ")";
            workSheet.Cells["H" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["I" + numberOfRows].Formula = "SUM(I12:I" + (numberOfRows - 1) + ")";
            workSheet.Cells["I" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["J" + numberOfRows].Formula = "SUM(J12:J" + (numberOfRows - 1) + ")";
            workSheet.Cells["J" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["K" + numberOfRows].Formula = "SUM(K12:K" + (numberOfRows - 1) + ")";
            workSheet.Cells["K" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["L" + numberOfRows].Formula = "SUM(L12:L" + (numberOfRows - 1) + ")";
            workSheet.Cells["L" + numberOfRows].Style.Font.Bold = true;

            numberOfRows += 2;
            workSheet.Cells["A" + numberOfRows].Value = ResourceManager.GetString("TotalIncome");
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["B" + numberOfRows].Formula = "B" + (numberOfRows - 2) + "+ C" + (numberOfRows - 2);
            workSheet.Cells["B" + numberOfRows].Style.Font.Bold = true;

            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = ResourceManager.GetString("TotalExpenses");
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["B" + numberOfRows].Formula = "SUM(D" + (numberOfRows - 3) + ":L" + (numberOfRows - 3) + ")";
            workSheet.Cells["B" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["A" + numberOfRows + ":B" + numberOfRows].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = ResourceManager.GetString("MonthlyOverview");
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["A" + numberOfRows].Style.Font.Color.SetColor(Color.Red);
            workSheet.Cells["B" + numberOfRows].Formula = "SUM(B" + (numberOfRows - 2) + ":B" + (numberOfRows - 1) + ")";
            workSheet.Cells["B" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["B" + numberOfRows].Style.Font.Color.SetColor(Color.Red);

            numberOfRows += 3;
            workSheet.Cells["A" + numberOfRows].Value = ResourceManager.GetString("YearlyOverview");
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["A" + numberOfRows].Style.Font.Color.SetColor(Color.Red);
            workSheet.Cells["A" + numberOfRows + ":B" + numberOfRows].Style.Border.BorderAround(ExcelBorderStyle.Thick);
            workSheet.Cells["B" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["B" + numberOfRows].Style.Font.Color.SetColor(Color.Red);

            NumberOfRowsList.Add(monthNumber, numberOfRows);
            
            switch (monthNumber)
            {
                case 1:
                    workSheet.Cells["B" + numberOfRows].Formula = "B" + (numberOfRows - 3);
                    break;
                case 2:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("January") + "!B" + NumberOfRowsList[1] + "+B" + (numberOfRows - 3);
                    break;
                case 3:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("February") + "!B" + NumberOfRowsList[2] + "+B" + (numberOfRows - 3);
                    break;
                case 4:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("March") + "!B" + NumberOfRowsList[3] + "+B" + (numberOfRows - 3);
                    break;
                case 5:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("April") + "!B" + NumberOfRowsList[4] + "+B" + (numberOfRows - 3);
                    break;
                case 6:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("May") + "!B" + NumberOfRowsList[5] + "+B" + (numberOfRows - 3);
                    break;
                case 7:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("June") + "!B" + NumberOfRowsList[6] + "+B" + (numberOfRows - 3);
                    break;
                case 8:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("July") + "!B" + NumberOfRowsList[7] + "+B" + (numberOfRows - 3);
                    break;
                case 9:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("August") + "!B" + NumberOfRowsList[8] + "+B" + (numberOfRows - 3);
                    break;
                case 10:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("September") + "!B" + NumberOfRowsList[9] + "+B" + (numberOfRows - 3);
                    break;
                case 11:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("October") + "!B" + NumberOfRowsList[10] + "+B" + (numberOfRows - 3);
                    break;
                case 12:
                    workSheet.Cells["B" + numberOfRows].Formula = ResourceManager.GetString("November") + "!B" + NumberOfRowsList[11] + "+B" + (numberOfRows - 3);
                    break;
            }

            var range = workSheet.Cells["A9:P22"];
            range.AutoFitColumns();
        }
        public void DeleteIfExists()
        {
            try
            {
                if (System.IO.File.Exists(File.FullName))
                {
                    System.IO.File.Delete(File.FullName);
                }
            }
            catch (IOException)
            {
                FilePreview.Text = ResourceManager.GetString("OpenErrorPart1") + File.FullName + ResourceManager.GetString("OpenErrorPart2");
            }
            catch (Exception)
            {
                FilePreview.Text = ResourceManager.GetString("SomethingWrong");
            }
        }
    }
}