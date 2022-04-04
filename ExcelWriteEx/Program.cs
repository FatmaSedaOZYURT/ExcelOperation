using ExcelReaderEx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriteEx
{
    class Program
    {
        static void Main(string[] args)
        {
            int i = 1;
            List<FileData> fileDatas = new List<FileData>(); //Dolu geldiğini düşünürsek.
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            if (app != null)
            {
                Microsoft.Office.Interop.Excel.Workbook myexcelWorkbook = app.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet myexcelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)myexcelWorkbook.Sheets.Add();

                myexcelWorksheet.Cells[1, 1] = "Başlık";
                myexcelWorksheet.Cells[1, 2] = "Açıklama";

                foreach (var item in fileDatas)
                {
                    i++;
                    myexcelWorksheet.Cells[i, 1] = item.Title;
                    myexcelWorksheet.Cells[i, 2] = item.Description;
                }

                app.ActiveWorkbook.SaveAs(@"C:\Users\Public\abc.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);

                myexcelWorkbook.Close();
                app.Quit();
            }
        }
    }
}
