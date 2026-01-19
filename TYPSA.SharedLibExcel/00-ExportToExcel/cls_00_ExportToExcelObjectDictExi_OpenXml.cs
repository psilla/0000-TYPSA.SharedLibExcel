using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportToExcelObjectDictExi_OpenXml
    {
        public static void ExportObjectDictToExcelExi(
            string excelPath,                 
            Dictionary<string, object> dict,
            List<string> headers,
            string sheetName,
            string tableName,
            bool overwriteSheet = true,
            bool openAfterExport = false
        )
        {
            try
            {
                // EPPlus setup
                System.Text.Encoding.RegisterProvider(
                    System.Text.CodePagesEncodingProvider.Instance
                );
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                if (dict == null || dict.Count == 0)
                    throw new ArgumentException("No hay datos para exportar.");

                FileInfo fi = new FileInfo(excelPath);
                if (!fi.Exists)
                    throw new FileNotFoundException(
                        "El archivo Excel no existe.", excelPath
                    );

                using (var package = new ExcelPackage(fi))
                {
                    ExcelWorksheet ws;

                    // ----------------------------------
                    // 1. CREAR / REEMPLAZAR HOJA
                    // ----------------------------------
                    var existingSheet = package.Workbook.Worksheets[sheetName];
                    if (existingSheet != null)
                    {
                        if (!overwriteSheet)
                            throw new InvalidOperationException(
                                $"La hoja '{sheetName}' ya existe."
                            );

                        package.Workbook.Worksheets.Delete(existingSheet);
                    }

                    ws = package.Workbook.Worksheets.Add(sheetName);

                    // ----------------------------------
                    // 2. ENCABEZADOS
                    // ----------------------------------
                    for (int col = 0; col < headers.Count; col++)
                    {
                        ws.Cells[1, col + 1].Value = headers[col];
                        ws.Cells[1, col + 1].Style.Font.Bold = true;
                    }

                    int row = 2;

                    // ----------------------------------
                    // 3. FILAS
                    // ----------------------------------
                    foreach (var kvp in dict)
                    {
                        object obj = kvp.Value;
                        Type objType = obj.GetType();

                        for (int col = 0; col < headers.Count; col++)
                        {
                            string propName = headers[col];
                            var prop = objType.GetProperty(propName);

                            ws.Cells[row, col + 1].Value =
                                prop != null ? prop.GetValue(obj) : "";
                        }

                        row++;
                    }

                    // Autoajustar columnas
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    // ----------------------------------
                    // 4. TABLA
                    // ----------------------------------
                    int totalRows = dict.Count + 1;
                    int totalCols = headers.Count;

                    var range = ws.Cells[1, 1, totalRows, totalCols];
                    var table = ws.Tables.Add(range, tableName);

                    table.ShowFilter = true;
                    table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

                    // ----------------------------------
                    // 5. GUARDAR
                    // ----------------------------------
                    package.Save();
                }

                // ----------------------------------
                // 6. ABRIR EXCEL (OPCIONAL)
                // ----------------------------------
                if (openAfterExport)
                {
                    Process.Start(new ProcessStartInfo(excelPath)
                    {
                        UseShellExecute = true
                    });
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
