using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportToExcelAdjacency_OpenXml
    {
        public static void ExportAdjacencyDictToExcel(
            Dictionary<string, string> adjDict,
            List<string> headers,
            string sheetName,
            string tableName
        )
        {
            try
            {
                // EPPlus setup
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add(sheetName);

                    if (headers == null || headers.Count < 2)
                        throw new ArgumentException("Debe haber al menos dos encabezados.");

                    // ------------------------
                    // HEADERS (dinámicos)
                    // ------------------------
                    for (int col = 0; col < headers.Count; col++)
                    {
                        ws.Cells[1, col + 1].Value = headers[col];
                        ws.Cells[1, col + 1].Style.Font.Bold = true;
                    }

                    int row = 2;

                    // ------------------------
                    // DATA ROWS
                    // ------------------------
                    foreach (var kvp in adjDict)
                    {
                        ws.Cells[row, 1].Value = kvp.Key;
                        ws.Cells[row, 2].Value = kvp.Value;
                        row++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    // ------------------------
                    // CONVERTIR A TABLA
                    // ------------------------
                    int totalRows = adjDict.Count + 1;
                    int totalCols = headers.Count;

                    string tableRange =
                        $"{ws.Cells[1, 1].Address}:{ws.Cells[totalRows, totalCols].Address}";

                    var tbl = ws.Tables.Add(ws.Cells[tableRange], tableName);

                    tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                    tbl.ShowFilter = true;

                    // ----------------------------------------------------
                    // 4. Guardar archivo temporal
                    // ----------------------------------------------------
                    // Guardar archivo temporal
                    using (MemoryStream ms = new MemoryStream())
                    {
                        package.SaveAs(ms);
                        string path = Path.Combine(
                            Path.GetTempPath(),
                            $"{sheetName}_{Guid.NewGuid()}.xlsx"
                        );

                        File.WriteAllBytes(path, ms.ToArray());
                        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERROR al exportar adyacencias:\n{ex}");
            }
        }


    }
}
