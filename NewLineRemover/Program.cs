using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.IO;
using System.Data;

namespace NewLineRemover
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "C:/Users/MarlonCopeland/Dup_Raw_SII_D_wk20_062418.XLSX";
            //string filePath = "C:/Users/MarlonCopeland/Dup_Raw_SII_C_wk20_060818.XLSX";

            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            ////1. Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            StringBuilder sb = new StringBuilder();

            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            string completeLine = "";
            //5. Data Reader methods
            int b = 0;
            while (excelReader.Read())
            {
                string fixedLine = "";


                for (int i = 0; i < 90; i++)
                {
                    string outputLine = excelReader[i] == null ? " " : excelReader[i].ToString();



                    if (excelReader[0].ToString() == "contract_id")
                    {
                        //do for first column
                        sb.Append((outputLine == null ? " " : outputLine) + "|");

                    }
                    else
                    {
                        if (excelReader[i] != null)
                        {
                            if (excelReader[i].ToString().Contains(Environment.NewLine))
                            {

                                outputLine = excelReader[i].ToString().Replace(Environment.NewLine, " ");

                            }
                        }
                        //remove time and date from columns
                        if (outputLine.Contains("12:00:00 AM") || outputLine.Contains("12/31/1899"))
                        {
                            outputLine = outputLine.Replace("12:00:00 AM", "").Replace("12/31/1899", "").Trim();
                        }

                        sb.Append(outputLine + "|");

                    }
                }
                sb.Append(Environment.NewLine);
                if (b > 57000)
                { break; }
                b++;

                //excelReader.GetInt32(0);
            }
            using (StreamWriter outputFile = new StreamWriter("C:/Users/MarlonCopeland/textoutput.txt"))
            {
                outputFile.WriteLine(sb.ToString());
            }
            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();
        }
    }
}
