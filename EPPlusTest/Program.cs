using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //Install-Package EPPlus -Version 4.5.1

            using (var package = new ExcelPackage())
            {
                var path = "C:\\Users\\chn3\\Desktop\\KT on OSIS\\testExcel.xlsx";
                var workSheet = package.Workbook.Worksheets.Add("TestWorksheet");
                var workSheetHidden = package.Workbook.Worksheets.Add("TestWorksheetHidden");


                //styling
                var modelTable = workSheet.Cells["A1:Z30"];
                modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;




                //1.Static Data
                workSheetHidden.Cells[1, 1].Value = 2.56;
                workSheetHidden.Cells[1, 2].Value = 2.1;
                workSheetHidden.Cells[2, 1].Value = 3.26;
                workSheetHidden.Cells[2, 2].Value = 3.1;

                workSheetHidden.Cells["A1:Z100"].Style.Locked = true;
                workSheet.Cells["A1:z100"].Style.Locked = true;
                workSheetHidden.Hidden = eWorkSheetHidden.VeryHidden;
                workSheet.Protection.SetPassword("abc");
                workSheetHidden.Protection.SetPassword("abc");
                
                var list=new List<decimal>() {1,2.3M,5.6M,6.6M};
                workSheet.Cells[2, 2].LoadFromCollection(list.Select(x => new  {Data= x}));

                workSheet.Cells[5, 1].Formula = "=SUM(TestWorksheetHidden!A1*TestWorksheetHidden!B1+TestWorksheetHidden!A2*TestWorksheetHidden!B2)";


                package.Workbook.Protection.LockStructure = true;
                package.Workbook.Protection.SetPassword("abc");


                workSheet.Column(1).Width = 50;
                workSheet.Column(5).Width = 60;


                //2.Collection of Data
                //string[] list = { "Sheldon", "Chandler", "Howard", "Frank" };
                //var range = workSheet.Cells["A0"].LoadFromCollection(list);


                //3.Dropdown From list
                //var cell = workSheet.Cells[15, 1];
                //var dataValidation = workSheet.DataValidations.AddListValidation(cell.Address);
                //dataValidation.Formula.ExcelFormula = "=" + range.FullAddressAbsolute;
                //dataValidation.AllowBlank = true;


                //4.Read Data
                //Console.WriteLine(ReadExcel(path));
                //Console.ReadLine();



                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

               

                File.Delete(path);
                using (var fileStream = File.Create(path))
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                }


                System.Diagnostics.Process.Start(path);
            }
        }

        private static string ReadExcel(string path)
        {
            string result = "";
            var fileInfo = new FileInfo(path);
            using (var package = new ExcelPackage(fileInfo))
            {
                var workSheet = package.Workbook
                    .Worksheets
                    .Single(ws => ws.Name == "TestWorksheet");
                result = workSheet.Cells[15, 1].Value.ToString();
            }

            return result;
        }
    }
}
