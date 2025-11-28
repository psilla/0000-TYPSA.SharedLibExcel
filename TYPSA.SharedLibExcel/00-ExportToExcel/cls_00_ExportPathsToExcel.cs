using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportPathsToExcel
    {
        public static void ExportPathsToExcel(
            string excelPath,
            List<(string Terminal, List<string> Path, double Length, bool IsComplete)> results
        )
        {
            // ======================================
            // 4. EXPORTAR EN LA MISMA HOJA DEL EXCEL
            // ======================================
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var ws = package.Workbook.Worksheets[0];

                // Detectar número máximo de segmentos usados por cualquier terminal
                int maxSegments = results.Max(r => r.Path.Count);

                int startCol = 4; // Columna D

                // -------------------------------------
                // ENCABEZADOS DINÁMICOS
                // -------------------------------------
                for (int i = 0; i < maxSegments; i++)
                    ws.Cells[1, startCol + i].Value = $"T {i + 1}";

                ws.Cells[1, startCol + maxSegments].Value = "TotalLength";
                ws.Cells[1, startCol + maxSegments + 1].Value = "IsComplete";

                // -------------------------------------
                // ESCRIBIR DATOS
                // -------------------------------------
                int row = 2;
                foreach (var r in results)
                {
                    // Segments
                    for (int i = 0; i < r.Path.Count; i++)
                        ws.Cells[row, startCol + i].Value = r.Path[i];

                    // Total length
                    ws.Cells[row, startCol + maxSegments].Value = r.Length;

                    // Final reached? (true/false)
                    ws.Cells[row, startCol + maxSegments + 1].Value = r.IsComplete;

                    row++;
                }

                // Guardar cambios
                package.Save();
            }
        }





    }
}
