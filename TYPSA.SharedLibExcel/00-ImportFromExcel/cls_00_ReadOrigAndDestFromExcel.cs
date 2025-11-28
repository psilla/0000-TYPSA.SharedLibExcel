using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ReadOrigAndDestFromExcel
    {
        public static List<(string Origin, string Destination)> ReadOriginDestinationExcel(string filePath)
        {
            List<(string Origin, string Destination)> list =
                new List<(string Origin, string Destination)>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = package.Workbook.Worksheets[0];

                int row = 2;

                while (ws.Cells[row, 2].Value != null && ws.Cells[row, 3].Value != null)
                {
                    string origin = ws.Cells[row, 2].Text.Trim();
                    string dest = ws.Cells[row, 3].Text.Trim();

                    list.Add((origin, dest));
                    row++;
                }
            }
            // return
            return list;
        }


    }
}
