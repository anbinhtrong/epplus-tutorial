using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace EpplusDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            ExportFile();
        }

        static void ExportFile()
        {
            var stream = new MemoryStream();
            using (var xlPackage = new ExcelPackage(stream))
            {
                #region sheet Images
                var worksheetImage = xlPackage.Workbook.Worksheets.Add("Images");
                var properties = new List<string>
                {
                    "Id",
                    "FileName"
                };
                for (var i = 0; i < properties.Count; i++)
                {
                    worksheetImage.Cells[1, i + 1].Value = properties[i];
                    worksheetImage.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheetImage.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                    worksheetImage.Cells[1, i + 1].Style.Font.Bold = true;
                }
                worksheetImage.View.FreezePanes(2, 1);
                #endregion
                worksheetImage.View.FreezePanes(2, 1);
                //todo insert body here
                worksheetImage.Cells.AutoFitColumns(0);
                    

                xlPackage.Save();
            }
            var file = new FileStream("export.xlsx", FileMode.Create, FileAccess.Write);            
            byte[] bytes = stream.ToArray();           
            file.Write(bytes, 0, bytes.Length);
            file.Dispose();

        }
    }
}