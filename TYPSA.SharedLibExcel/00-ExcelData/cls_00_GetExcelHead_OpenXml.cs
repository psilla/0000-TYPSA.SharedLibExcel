using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Windows.Forms;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_GetExcelHead_OpenXml
    {
        public static List<string> GetExcelHeaders(
            string filePath,
            string sheetName,
            int startColumn
        )
        {
            List<string> headers = new List<string>();

            // try
            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // Obtener hoja
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
                    // Validamos
                    if (worksheet != null)
                    {
                        // Mensaje
                        MessageBox.Show($"⚠️ The sheet '{sheetName}' does not exist in the file.", "Sheet Error");
                        // Finalizamos
                        return null;
                    }

                    // Obtener el número total de columnas
                    int colCount = worksheet.Dimension.Columns;

                    // Verificar que la columna inicial esté dentro del rango de columnas disponibles
                    if (startColumn > colCount)
                    {
                        // Mensaje
                        MessageBox.Show($"⚠️ The sheet does not have column number {startColumn} or beyond.", "Column Warning");
                        // Finalizamos
                        return null;
                    }

                    // Leer los encabezados desde la columna indicada (argumento startColumn)
                    for (int col = startColumn; col <= colCount; col++)
                    {
                        // Leer el encabezado de la primera fila
                        string header = worksheet.Cells[1, col].Text;
                        headers.Add(header);
                    }

                    // Ordenar los encabezados alfabéticamente (sin distinguir mayúsculas/minúsculas)
                    List<string> orderedHeaders = headers
                        .OrderBy(h => h, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    // Mensaje
                    MessageBox.Show(
                        $"✅ Headers found in '{sheetName}' (starting from column {startColumn}):\n\n" +
                        $"{string.Join("\n", orderedHeaders)}",
                        "Excel Headers"
                    );
                }
            }
            // catch
            catch (Exception ex)
            {
                MessageBox.Show($"❌ ERROR while reading Excel headers:\n{ex.Message}", "Error");
                return null;
            }

            // return
            return headers;
        }

    }
}
