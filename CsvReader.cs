using System.Collections.Generic;

namespace FinancialAdministrator
{
    internal class CsvReader
    {
        string[] readFile { get; set; }
        public CsvReader(string[] readfile)
        {
            readFile = readfile;
        }

        public List<TransactieModel> parseFile()
        {
            List<TransactieModel> result = new List<TransactieModel>();
            for (int i = 1; i < readFile.Length; i++)
            {
                string[] rowData = readFile[i].Split(';');
                result.Add(new TransactieModel(rowData[4], rowData[6], rowData[2], rowData[8], rowData[9]));
            }
            return result;
        }

    }
}
