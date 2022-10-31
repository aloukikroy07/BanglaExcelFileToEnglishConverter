using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Office.Interop.Excel;
using _excel = Microsoft.Office.Interop.Excel;

namespace Excel_File_Bangla_To_English
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Please enter path of the file with filename:");
            string filePath = Console.ReadLine();
            Console.Write("Please enter path where you want to save the file:");
            string outputPath = Console.ReadLine();

            Console.WriteLine("Please wait and please dont close the window");

            Application excelApp = new Application();


            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }

            Workbook excelBook = excelApp.Workbooks.Open(filePath);

            string fileName = excelBook.Name;

            _Application excel = new _excel.Application(); ;
            Workbook workBook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet); 

            int count = excelBook.Sheets.Count;

            for(int x = 1; x<=count; x++)
            {
                _Worksheet excelSheet = excelBook.Sheets[x];
                string sheetName = excelSheet.Name;

                Worksheet worksheet = workBook.Worksheets.Add();
                worksheet.Name=sheetName;

                Range excelRange = excelSheet.UsedRange;

                int rows = excelRange.Rows.Count;
                int cols = excelRange.Columns.Count;

                for (int i = 1; i <= rows; i++)
                {
                    //create new line
                    //Console.Write("\r\n");
                    for (int j = 1; j <= cols; j++)
                    {
                        if (excelRange.Cells[i, j].Value2 == null)
                        {
                            
                        }

                        else{
                            //string word1 = excelRange.Cells[i, j].Value2.ToString();
                            //string[] wordList = word1.Split(new char[] { ',', ';', ' ' });
                            //string sentence = null;
                            //foreach (string wd in wordList)
                            //{
                            //    if(Regex.IsMatch(wd, @"[a-zA-Z0-9]"))
                            //    {
                            //        sentence = sentence+" "+ wd;
                            //    }

                            //    else
                            //    {
                            //        string word = Translate(word1);
                            //        sentence = sentence + " " + word;
                            //    }
                            //}
                            //worksheet.Cells[i, j] = sentence;

                            string word1 = excelRange.Cells[i, j].Value2.ToString();
                            string word = Translate(word1);
                            worksheet.Cells[i, j] = word;
                        }
                    }
                }
            }

            workBook.SaveAs(outputPath+@"\"+fileName);
            //after reading, relaase the excel project
            
            excelBook.Close();
            workBook.Close();
            excelApp.Quit();
            Console.Clear();
            Console.WriteLine("Successfully converted and saved to the path. If you want to convert again then press enter. If you want to exit then just close the window.");
            Console.ReadKey();
            Console.Clear();
            Main(null);
        }

        public static string Translate(string word)
        {
            var toLanguage = "en";//English
            var fromLanguage = "bn";//Deutsch
            var url = $"https://translate.googleapis.com/translate_a/single?client=gtx&sl={fromLanguage}&tl={toLanguage}&dt=t&q={HttpUtility.UrlEncode(word)}";
            var webClient = new WebClient
            {
                Encoding = System.Text.Encoding.UTF8
            };
            var result = webClient.DownloadString(url);
            try
            {
                result = result.Substring(4, result.IndexOf("\"", 4, StringComparison.Ordinal) - 4);
                return result;
            }
            catch
            {
                return "Error";
            }
        }
    }
}
