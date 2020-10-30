using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Avinch.RTFToTextConverter;

namespace Sakov.Evgeni1
{
    class Programm
    {

       static readonly string txtFile = @"text.txt";
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
                ofd.ShowDialog();
                
                string text = RTFToText.converting().rtfFromFile(ofd.FileName);

                FileStream Filetxt = File.Create(txtFile);
                Filetxt.Close();
                string FirstData = "Регистрационный номер сделки";
                string SecondData = "Номер договора";
                string ThirdData = "Счет контрагента";
                string FouthData = "Адрес контрагента";
                string FifthData = "Наименование договора";

                StreamWriter SW = new StreamWriter(txtFile);
                SW.Write(text);
                SW.Close();
                
                StreamReader SR = new StreamReader(txtFile);
                string line;

                // Создаём экземпляр нашего приложения
                Excel.Application excelApp = new Excel.Application();
                // Создаём экземпляр рабочий книги Excel
                Excel.Workbook workBook;
                // Создаём экземпляр листа Excel
                Excel.Worksheet workSheet;

                workBook = excelApp.Workbooks.Add();
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                // Заполняем первую строку числами от 1 до 10
                // Открываем созданный excel-файл
                int i = 1;
                while ((line = SR.ReadLine()) != null)
                {
                    if (line.StartsWith(FirstData) && line != "")
                    {
                        workSheet.Cells[i, 1] = FirstData;
                        FirstData = line.Remove(0, FirstData.Length + 1);
                        FirstData.Trim();
                        workSheet.Cells[i, 2] = FirstData;
                        i += 1;
                    }
                    else if (line.StartsWith(SecondData) && line != "")
                    {
                        workSheet.Cells[i, 1] = SecondData;
                        SecondData = line.Remove(0, SecondData.Length + 1);
                        SecondData.Trim();
                        workSheet.Cells[i, 2] = SecondData;
                        i += 1;
                    }
                    else if (line.StartsWith(ThirdData) && line != "")
                    {
                        workSheet.Cells[i, 1] = ThirdData;
                        ThirdData = line.Remove(0, ThirdData.Length + 1);
                        ThirdData.Trim();
                        workSheet.Cells[i, 2] = ThirdData;
                        i += 1;
                    }
                    else if (line.StartsWith(FouthData) && line != "")
                    {
                        workSheet.Cells[i, 1] = FouthData;
                        FouthData = line.Remove(0, FouthData.Length + 1);
                        FouthData.Trim();
                        workSheet.Cells[i, 2] = FouthData;
                        i += 1;
                    }
                    else if (line.StartsWith(FifthData) && line != "")
                    {
                        workSheet.Cells[i, 1] = FifthData;
                        FifthData = line.Remove(0, FifthData.Length + 1);
                        FifthData.Trim();
                        workSheet.Cells[i, 2] = FifthData;
                        i += 1;
                    }
                }
                SR.Close();
                File.Delete(txtFile);
                Excel.Range range = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[5, 2]];
                range.EntireColumn.AutoFit();
                range.EntireRow.AutoFit();
                excelApp.Visible = true;
                excelApp.UserControl = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

    }
}
