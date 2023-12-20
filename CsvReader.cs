using System.Collections.Generic;

namespace FinancialAdministrator
{
    internal class CsvReader(string[] readfile)
    {
        public async Task<List<TransactieModel>> parseFileAsync()
        {
            List<TransactieModel> result = new();
            for (int i = 1; i < readfile.Length; i++)
            {
                string[] rowData = readfile[i].Split(';');
                await Task.Run(() => result.Add(new TransactieModel(rowData[4], rowData[6], rowData[2], rowData[8], rowData[9])));
            }
            return result;
        }
    }
}
