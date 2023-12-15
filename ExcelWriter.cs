using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.NetworkInformation;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Style;
using System.IO.Packaging;
using System.Runtime.CompilerServices;
using System.Collections;

namespace FinancialAdministrator
{
    internal class ExcelWriter
    {
        public string fileName { get; set; }

        public List<TransactieModel> data { get; set; }

        public int administrationYear { get; set; }

        public Hashtable numberOfRowsList = new Hashtable();

        ProgressBar progressBar1;

        public ExcelWriter(string filename, List<TransactieModel> result, int administrationyear, ProgressBar progressBar1)
        {
            fileName = filename;
            data = result;
            administrationYear = administrationyear;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            this.progressBar1 = progressBar1;
        }

        public async Task writeFile()
        {
            await saveExcelFile(data, fileName);
        }

        public async Task saveExcelFile(List<TransactieModel> data, string fileName)
        {
            deleteIfExists(fileName);
            
            using (var package = new ExcelPackage(fileName))
            {
                await createMonthSheet("Januari", 1, 31, package);
                if (administrationYear % 4 == 0)
                {
                    await createMonthSheet("Februari", 2, 29, package);
                }
                else
                {
                    await createMonthSheet("Februari", 2, 28, package);
                }
                progressBar1.Visible = true;
                await createMonthSheet("Maart", 3, 31, package);
                progressBar1.Value = 10;
                await createMonthSheet("April", 4, 30, package);
                await createMonthSheet("Mei", 5, 31, package);
                await createMonthSheet("Juni", 6, 30, package);
                await createMonthSheet("Juli", 7, 31, package);
                await createMonthSheet("Augustus", 8, 31, package);
                progressBar1.Value = 50;
                await createMonthSheet("September", 9, 30, package);
                await createMonthSheet("Oktober", 10, 31, package);
                await createMonthSheet("November", 11, 30, package);
                await createMonthSheet("December", 12, 31, package);

                await package.SaveAsync();
                progressBar1.Value = 100;

                if (Type.GetTypeFromProgID("Excel.Application", false) != null)
                {
                    System.Diagnostics.Process.Start(new ProcessStartInfo { FileName = fileName, UseShellExecute = true } );
                }
                progressBar1.Value = 0;
                progressBar1.Visible = false;
            }
        }

        private async Task createMonthSheet(string monthName, int monthNumber, int endDay, ExcelPackage package)
        {
            var workSheet = package.Workbook.Worksheets.Add(monthName);
            
            var monthData = data.Where(transactie => 
                transactie.Boekingsdatum >= new DateTime(administrationYear, monthNumber, 01)
                && transactie.Boekingsdatum < new DateTime(administrationYear, monthNumber, endDay))
                .OrderBy(transactie => transactie.Boekingsdatum);

            int numberOfRows = 0;

            for (int i = 0; i < monthData.Count(); i++)
            {
                switch (monthData.ElementAt(i).Categorie) {
                    case "Boodschappen":
                        workSheet.Cells["H" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                        break;
                    case "Storting":
                        workSheet.Cells["C" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                        break;
                    case "Giften":
                        workSheet.Cells["G" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                        break;
                    case "Auto":
                        workSheet.Cells["I" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                        break;
                    case "Ziekte":
                        workSheet.Cells["F" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                        break;
                    case "Verzekeringen":
                        workSheet.Cells["E" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
                        break;
                    case "Belastingvermindering":
                        workSheet.Cells["L" + (i + 12)].Value = monthData.ElementAt(i).Bedrag;
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

            await applyTemplate(workSheet, numberOfRows, monthName);          
        }

        private async Task applyTemplate(ExcelWorksheet workSheet, int numberOfRows, string monthName)
        {
            workSheet.Column(15).Style.Numberformat.Format = "dd-MM-yyyy";
            workSheet.Cells["A1"].Value = "Financiele Administratie " + administrationYear;
            workSheet.Cells["A1:B1"].Merge = true;
            workSheet.Row(1).Style.Font.Size = 14;

            workSheet.Cells["B9"].Value = "Inkomsten naar";
            workSheet.Cells["B9:C9"].Merge = true;
            workSheet.Cells["B9:C9"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["E9:P9"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["A10:P10"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            workSheet.Cells["A11:P11"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["A11:P11"].Style.Fill.SetBackground(Color.LightGray);

            workSheet.Cells["D9"].Value = "Uitgaven van";
            workSheet.Cells["E9"].Value = "Uitgaven van betaalrekening";
            workSheet.Cells["E9:L9"].Merge = true;
            workSheet.Cells["A10"].Value = "Omschrijving";
            workSheet.Cells["A11"].Value = "ING Lion Account Zichtrekening";
            workSheet.Cells["B10"].Value = "spaarrek.";
            workSheet.Cells["C10"].Value = "betaalrek.";
            workSheet.Cells["D10"].Value = "spaarrekening";
            workSheet.Cells["E10"].Value = "Verzekeringen";
            workSheet.Cells["F10"].Value = "Ziekte";
            workSheet.Cells["G10"].Value = "Giften";
            workSheet.Cells["H10"].Value = "Boodschappen";
            workSheet.Cells["I10"].Value = "Auto";
            workSheet.Cells["J10"].Value = "Werk";
            workSheet.Cells["K10"].Value = "Overig";
            workSheet.Cells["L10"].Value = "Belastingvermindering";
            workSheet.Cells["M10"].Value = "Tegenrekening";
            workSheet.Cells["N10"].Value = "Omschrijving";
            workSheet.Cells["O10"].Value = "Datum";
            workSheet.Cells["P10"].Value = "Details";
            workSheet.Cells["A9:P10"].Style.Font.Bold = true;

            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = "ING BELGIË Groen Boekje";
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Fill.SetBackground(Color.LightGray);
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            
            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = "rente";

            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = "ING BELGIË Lion Deposit";
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Fill.SetBackground(Color.LightGray);
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = "-";

            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = "ASN NEDERLAND";
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Fill.SetBackground(Color.LightGray);
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            numberOfRows += 10;
            workSheet.Cells["A" + numberOfRows].Value = "ING NEDERLAND";
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Fill.SetBackground(Color.LightGray);
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            numberOfRows += 5;
            workSheet.Cells["A" + numberOfRows + ":P" + numberOfRows].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

            workSheet.Cells["B9:B" + numberOfRows].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["C9:C" + numberOfRows].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["D9:D" + numberOfRows].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = "Subtotalen";
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
            workSheet.Cells["A" + numberOfRows].Value = "Totalen inkomsten";
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["B" + numberOfRows].Formula = "B" + (numberOfRows - 2) + "+ C" + (numberOfRows - 2);
            workSheet.Cells["B" + numberOfRows].Style.Font.Bold = true;

            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = "Totalen uitgaven";
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["B" + numberOfRows].Formula = "SUM(D" + (numberOfRows - 3) + ":L" + (numberOfRows - 3) + ")";
            workSheet.Cells["B" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["A" + numberOfRows + ":B" + numberOfRows].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

            numberOfRows++;
            workSheet.Cells["A" + numberOfRows].Value = "Maandoverzicht";
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["A" + numberOfRows].Style.Font.Color.SetColor(Color.Red);
            workSheet.Cells["B" + numberOfRows].Formula = "SUM(B" + (numberOfRows - 2) + ":B" + (numberOfRows - 1) + ")";
            workSheet.Cells["B" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["B" + numberOfRows].Style.Font.Color.SetColor(Color.Red);

            numberOfRows += 3;
            workSheet.Cells["A" + numberOfRows].Value = "Jaaroverzicht";
            workSheet.Cells["A" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["A" + numberOfRows].Style.Font.Color.SetColor(Color.Red);
            workSheet.Cells["A" + numberOfRows + ":B" + numberOfRows].Style.Border.BorderAround(ExcelBorderStyle.Thick);
            workSheet.Cells["B" + numberOfRows].Style.Font.Bold = true;
            workSheet.Cells["B" + numberOfRows].Style.Font.Color.SetColor(Color.Red);

            numberOfRowsList.Add(monthName, numberOfRows);

            switch (monthName)
            {
                case "Januari":
                    workSheet.Cells["B" + numberOfRows].Formula = "B" + (numberOfRows - 3); 
                    break;
                case "Februari":
                    workSheet.Cells["B" + numberOfRows].Formula = "Januari!B" + numberOfRowsList["Januari"] + "+B" + (numberOfRows - 3);
                    break;
                case "Maart":
                    workSheet.Cells["B" + numberOfRows].Formula = "Februari!B" + numberOfRowsList["Februari"] + "+B" + (numberOfRows - 3);
                    break;
                case "April":
                    workSheet.Cells["B" + numberOfRows].Formula = "Maart!B" + numberOfRowsList["Maart"] + "+B" + (numberOfRows - 3);
                    break;
                case "Mei":
                    workSheet.Cells["B" + numberOfRows].Formula = "April!B" + numberOfRowsList["April"] + "+B" + (numberOfRows - 3);
                    break;
                case "Juni":
                    workSheet.Cells["B" + numberOfRows].Formula = "Mei!B" + numberOfRowsList["Mei"] + "+B" + (numberOfRows - 3);
                    break;
                case "Juli":
                    workSheet.Cells["B" + numberOfRows].Formula = "Juni!B" + numberOfRowsList["Juni"] + "+B" + (numberOfRows - 3); 
                    break;
                case "Augustus":
                    workSheet.Cells["B" + numberOfRows].Formula = "Juli!B" + numberOfRowsList["Juli"] + "+B" + (numberOfRows - 3);
                    break;
                case "September":
                    workSheet.Cells["B" + numberOfRows].Formula = "Augustus!B" + numberOfRowsList["Augustus"] + "+B" + (numberOfRows - 3);
                    break;
                case "Oktober":
                    workSheet.Cells["B" + numberOfRows].Formula = "September!B" + numberOfRowsList["September"] + "+B" + (numberOfRows - 3);
                    break;
                case "November":
                    workSheet.Cells["B" + numberOfRows].Formula = "Oktober!B" + numberOfRowsList["Oktober"] + "+B" + (numberOfRows - 3);
                    break;
                case "December":
                    workSheet.Cells["B" + numberOfRows].Formula = "November!B" + numberOfRowsList["November"] + "+B" + (numberOfRows - 3);
                    break;
            }

            var range = workSheet.Cells["A9:P22"];
            range.AutoFitColumns();
        }

        private void deleteIfExists(string fileName)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
        }
    }
}
