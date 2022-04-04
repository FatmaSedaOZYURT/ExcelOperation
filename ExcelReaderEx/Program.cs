using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderEx
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\file.xlsx";

            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            IExcelDataReader excelReader;
            List<FileData> fileDatas = new List<FileData>();
            int counter = 0;

            if (Path.GetExtension(filePath).ToUpper() == ".XLS")
            {
                //Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                //Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }


            //Veriler okunmaya başlıyor.
            while (excelReader.Read())
            {
                counter++;

                //ilk satır başlık olduğu için 2.satırdan okumaya başlıyorum.
                if (counter > 1)
                {
                    fileDatas.Add(new FileData() { Title = excelReader.GetValue(0).ToString(), Description = excelReader.GetValue(1).ToString() });
                }
            }
            // Okuma bitiriliyor.
            excelReader.Close();
        }
    }
}
