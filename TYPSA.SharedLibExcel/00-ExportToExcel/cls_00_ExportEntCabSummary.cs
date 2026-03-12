using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportEntCabSummary
    {
        public static void ExportEntCabSummaryWithPhases(
            string filePath,
            string sheetName,
            Dictionary<string, (double cableLength, double cableLengthCorrected, double cableLengthCorrectedTotal, double totalInstalledCableLength)> summary,
            string title,
            int startColumn
        )
        {

            // EPPlus setup
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = package.Workbook.Worksheets[sheetName];
                if (ws == null) return;

                int row = 1;

                int headerCount = 5;
                int lastColumn = startColumn + headerCount - 1;
                // Titulo
                ws.Cells[row, startColumn, row, lastColumn].Merge = true;
                ws.Cells[row, startColumn].Value = title;
                ws.Cells[row, startColumn].Style.Font.Bold = true;
                ws.Cells[row, startColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, startColumn].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Row(row).Height = 20;
                row++;

                // Headers
                ws.Cells[row, startColumn].Value = "Label";
                ws.Cells[row, startColumn + 1].Value = "CableLength";
                ws.Cells[row, startColumn + 2].Value = "CableLengthCorrected";
                ws.Cells[row, startColumn + 3].Value = "CableLengthCorrectedTotal";
                ws.Cells[row, startColumn + 4].Value = "TotalInstalledCableLength";
                ws.Cells[row, startColumn, row, startColumn + 4].Style.Font.Bold = true;
                row++;

                // Datos
                foreach (var kvp in summary.OrderBy(k => k.Key))
                {
                    ws.Cells[row, startColumn].Value = kvp.Key;
                    ws.Cells[row, startColumn + 1].Value = Math.Round(kvp.Value.cableLength, 2);
                    ws.Cells[row, startColumn + 2].Value = Math.Round(kvp.Value.cableLengthCorrected, 2);
                    ws.Cells[row, startColumn + 3].Value = Math.Round(kvp.Value.cableLengthCorrectedTotal, 2);
                    ws.Cells[row, startColumn + 4].Value = Math.Round(kvp.Value.totalInstalledCableLength, 2);
                    row++;
                }

                // Formato
                ws.Cells[1, startColumn, row, startColumn + 4].AutoFitColumns();

                // Guardar
                package.Save();
            }
        }
    }
}
