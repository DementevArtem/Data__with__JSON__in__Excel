using System;
using System.IO;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace FarmData
{
    class Data
    {
        internal string login { get; set; }
        internal string password { get; set; }
        internal string name { get; set; }
        internal string birthDay { get; set; }
        internal string cookie { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {

            string pathTextFile = "C:\\Users\\" + Environment.UserName + "\\Desktop\\DataFarm\\DataAutoReg.txt";
            string pathExelFile = "C:\\Users\\" + Environment.UserName + "\\Desktop\\DataFarm\\DataFarm.xlsx";

            var ex = new Excel.Application();

            Excel.Workbook wb = ex.Workbooks.Add(1);
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

            bool failed = false;

            using (StreamReader sr = new StreamReader(pathTextFile, Encoding.Default))
            {
                string dataString;
                int i = 1;
                while ((dataString = sr.ReadLine()) != null)
                {
                    string[] dataArray = dataString.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);

                    var data = new Data();
                    if(dataArray.Length == 7)
                    {
                        data.login = dataArray[0];
                        data.password = dataArray[1];
                        data.name = dataArray[3];
                        data.birthDay = dataArray[4];
                        data.cookie = dataArray[6];
                    }
                    else
                    {
                        data.login = dataArray[0];
                        data.password = dataArray[1];
                        data.name = dataArray[2];
                        data.birthDay = dataArray[3];
                        data.cookie = dataArray[5];
                    }
                    
                    var json = Encoding.UTF8.GetString(Convert.FromBase64String(data.cookie));
                    
                    do
                    {
                        try
                        {
                            sheet.Name = "DataFarm";
                            sheet.Cells[i, 1] = data.login;
                            sheet.Cells[i, 2] = data.password;
                            sheet.Cells[i, 4] = data.name;
                            sheet.Cells[i, 5] = data.birthDay;
                            sheet.Cells[i, 6] = json;

                            failed = false;
                        }
                        catch (System.Runtime.InteropServices.COMException e)
                        {
                            failed = true;
                        }
                        System.Threading.Thread.Sleep(10);
                    } while (failed);
                    i++;
                }
            }

            wb.SaveAs(@"C:\Users\Артем\Desktop\DataFarm\DataFarm.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);
            wb.Close(false, Type.Missing, Type.Missing);

            Console.WriteLine("Success!");
            Console.ReadLine();
        }
    }
}
