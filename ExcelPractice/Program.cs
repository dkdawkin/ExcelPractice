using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelPractice
{
    class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            Console.Write(program.readExcel("C:/Users/Delante/Desktop/Schedule.xlsx").ToString());
            Console.WriteLine("test successful");

        }
        public int readExcel(String fileName)
        {
            int workSheets = 0;
            try
            {
                ExcelPackage excel = new ExcelPackage(new System.IO.FileInfo(fileName));
                ExcelWorksheets wrkSheets = excel.Workbook.Worksheets;
                foreach (ExcelWorksheet wrkSheet in wrkSheets)
                {
                    for (int i = wrkSheet.Dimension.Start.Row; i <= wrkSheet.Dimension.End.Row;i++)
                    {
                        for (int j = wrkSheet.Dimension.Start.Column;j <= wrkSheet.Dimension.End.Column;j++)
                        {
                            object cellValue = wrkSheet.Cells[j,i].Value;
                            Console.Write(" "+cellValue.ToString()+" ");
                        }
                        Console.WriteLine("");
                    }
                }

            }
            catch (Exception ex)
            {

            }

            return workSheets;
        }
    }
}
