using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.IO;
using Spire.Xls;
using Spire.Pdf;

namespace downxl
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet wsSheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
            using (ExcelRange Rng = wsSheet1.Cells[1, 1, 4, 4])
            {
                Rng.Value = "Report Notification Config";
                wsSheet1.Cells[2, 1].Value = "Reporttype";

                //Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#B7DEE8");

                //coloums
                wsSheet1.Cells[3, 1].Value = "Name";
                wsSheet1.Cells[3, 1].Style.Font.Bold = true;
                // wsSheet1.Cells[3, 1][3, 2][3, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                // wsSheet1.Cells[3, 1][3, 2][3, 3].Style.Fill.BackgroundColor.SetColor(colFromHex);
                wsSheet1.Cells[3, 2].Value = "EmailAddresses";
                wsSheet1.Cells[3, 2].Style.Font.Bold = true;
                wsSheet1.Cells[3, 3].Value = "PhoneNumber";
                wsSheet1.Cells[3, 3].Style.Font.Bold = true;

                //data
                wsSheet1.Cells[4, 1].Value = "sai";
                wsSheet1.Cells[4, 2].Value = "skputta@deltaintech.com";
                wsSheet1.Cells[4, 3].Value = "7036355827";

                Rng.Merge = true;
                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }
            wsSheet1.Protection.IsProtected = false;
            wsSheet1.Protection.AllowSelectLockedCells = false;

            Workbook wb = new Workbook();
            wb.LoadFromFile(@"D:\New5.xlsx");
            wb.ConverterSetting.SheetFitToPage = true;
            Worksheet sheet = wb.Worksheets[0];
            for (int i = 1; i < sheet.Columns.Length; i++)
            {
                sheet.AutoFitColumn(i);
            }
            sheet.SaveToPdf(@"D:\toPDF.pdf");

            //Workbook workbook = new Workbook();
            //workbook.LoadFromFile(@"D:\New5.xlsx", ExcelVersion.Version2010);
            //workbook.SaveToFile(@"D:\result.pdf", Spire.Xls.FileFormat.PDF);

            //excelPackage.SaveAs(new FileInfo(@"D:\New5.xlsx"));
            Console.WriteLine("File has been Downloaded...!");
        }

    }
}
