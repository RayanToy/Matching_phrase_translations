using System;
using System.Collections.Generic;
using System.IO; 
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
namespace TestZad2
{
    class Program
    {
        static void Main(string[] args)
    {       
            //Запуск эксель приложеня
            Application excelApp = new Application();
            Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\User\source\repos\TestZad2\TestZad2\ViewerMessages.xlsx");
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;
            // Подсчёт строк и столбцов со значениями из эксель файла
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;
            string[] Rus = new string[1 + rows];
            string[] Eng = new string[1 + rows];
            //Записываю в стринг массивы из экселя русские слова и их английские варианты
            for (int i = 1; i <= rows; i++)
            {
                Eng[i] = excelRange.Cells[i, 1].Value2.ToString();
                Rus[i] = excelRange.Cells[i, 2].Value2.ToString();
            }
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            var document = XDocument.Load(@"C:\Users\User\source\repos\TestZad2\TestZad2\ViewerMessages.xml");
            for(int i = 1; i < rows + 1; i++)
            {
                // Выбираю из root node с русским вариантом для каждого id
                var variantRus =
               (
                   from m in document.Root.Elements("message")
                   where (int)m.Attribute("id") == i
                   from v in m.Elements("variant")
                   where (string)v.Attribute("language") == "ru_RU"
                   select v
               ).FirstOrDefault();
                // Выбираю из root node с английским вариантом
                var variantEng =
                (
                    from m in document.Root.Elements("message")
                    where (int)m.Attribute("id") == i
                    from v in m.Elements("variant")
                    where (string)v.Attribute("language") == "en_US"
                    select v
                ).FirstOrDefault();
                // Сравниваю английские слова из node с английскими словами из excel и если есть совпадение, то записываю в node c аттрибутом ru_RU русский вариант
                for (int j = 1; j < rows + 1; j++)
                {
                    if (variantEng.Value == Eng[j])
                    {
                        variantRus.SetValue(Rus[j]);
                    }
                }
            }
            document.Save(@"C:\Users\User\source\repos\TestZad2\TestZad2\ViewerMessages.xml");
        }
    }
}


