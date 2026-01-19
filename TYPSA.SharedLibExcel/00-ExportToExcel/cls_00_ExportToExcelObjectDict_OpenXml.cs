using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportToExcelObjectDict_OpenXml
    {
        public static void ExportObjectDictToExcel(
            Dictionary<string, object> dict,
            List<string> headers,           
            string sheetName,
            string tableName
        )
        {
            try
            {
                // EPPlus setup
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add(sheetName);

                    if (dict == null || dict.Count == 0)
                        throw new ArgumentException("No hay datos para exportar.");

                    // ---------------------------
                    // 1. ESCRIBIR ENCABEZADOS
                    // ---------------------------
                    for (int col = 0; col < headers.Count; col++)
                    {
                        ws.Cells[1, col + 1].Value = headers[col];
                        ws.Cells[1, col + 1].Style.Font.Bold = true;
                    }

                    int row = 2;

                    // ---------------------------
                    // 2. ESCRIBIR FILAS
                    // ---------------------------
                    foreach (var kvp in dict)
                    {
                        object obj = kvp.Value;
                        Type objType = obj.GetType();

                        // Para cada encabezado buscar la propiedad con ese nombre
                        for (int col = 0; col < headers.Count; col++)
                        {
                            string propertyName = headers[col];

                            var prop = objType.GetProperty(propertyName);
                            if (prop != null)
                            {
                                object value = prop.GetValue(obj);
                                ws.Cells[row, col + 1].Value = value;
                            }
                            else
                            {
                                // Propiedad NO existe → celda vacía
                                ws.Cells[row, col + 1].Value = "";
                            }
                        }

                        row++;
                    }

                    // Ajustar columnas
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    // ----------------------------------------------------
                    // 3. CREAR TABLA EPPLUS
                    // ----------------------------------------------------
                    int totalRows = dict.Count + 1; // +1 header
                    int totalCols = headers.Count;

                    var tblRange = ws.Cells[1, 1, totalRows, totalCols];
                    var table = ws.Tables.Add(tblRange, tableName);

                    table.ShowFilter = true;
                    table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

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
                MessageBox.Show(
                    $"ERROR al exportar a Excel:\n{ex.Message}",
                    "Error de Exportación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }



    }
}
