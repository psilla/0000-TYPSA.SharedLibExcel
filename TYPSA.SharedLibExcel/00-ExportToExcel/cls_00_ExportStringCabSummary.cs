using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportStringCabSummary
    {
        public static void ExportStringCabSummary(
            string excelPath,
            string sheetName,
            Dictionary<string, double> summary,
            string title,
            int startColumn
        )
        {
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var ws = package.Workbook.Worksheets[sheetName];
                if (ws == null) return;

                int row = 1;

                // Título
                ws.Cells[row, startColumn].Value = title;
                ws.Cells[row, startColumn].Style.Font.Bold = true;
                row++;

                // Headers
                ws.Cells[row, startColumn].Value = "Group";
                ws.Cells[row, startColumn + 1].Value = "Total Cable Length";
                ws.Cells[row, startColumn, row, startColumn + 1].Style.Font.Bold = true;
                row++;

                // Datos
                foreach (var kvp in summary.OrderBy(k => k.Key))
                {
                    ws.Cells[row, startColumn].Value = kvp.Key;
                    ws.Cells[row, startColumn + 1].Value = kvp.Value;
                    row++;
                }

                ws.Cells[1, startColumn, row, startColumn + 1].AutoFitColumns();

                package.Save();
            }
        }
    }
}
