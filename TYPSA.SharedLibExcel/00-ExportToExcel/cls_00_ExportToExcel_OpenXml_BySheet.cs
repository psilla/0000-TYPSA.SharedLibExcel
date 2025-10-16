using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using OfficeOpenXml.Table;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportToExcel_OpenXml_BySheet
    {
        public static void ExportToExcel_OpenXml_BySheet(
            Dictionary<string,
            List<List<object>>> data
        )
        {
            // try
            try
            {
                // Necesario en .NET Core/.NET 5+ para que EPPlus pueda manejar ciertas codificaciones como IBM437.
                System.Text.Encoding.RegisterProvider(
                    System.Text.CodePagesEncodingProvider.Instance
                );

                // Se establece el contexto de licencia para uso no comercial. Obligatorio desde EPPlus 5.
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Crear un nuevo archivo Excel en memoria
                using (var package = new ExcelPackage())
                {
                    List<string> invalidTableNames = new List<string>();
                    // Cada entrada representa una hoja de Excel
                    foreach (var entry in data)
                    {
                        // Obtener nombre de la hoja
                        // Excel no permite nombres de hoja mayores a 31 caracteres, así que los truncamos en ese caso.
                        string sheetName = entry.Key.Length > 31 ? entry.Key.Substring(0, 31) : entry.Key;

                        // Creamos una hoja llamada con ese nombre
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                        // Obtener las filas y columnas asociadas a esta hoja
                        var rows = entry.Value;
                        // Validamos
                        if (rows.Count == 0) continue;

                        // Contamos
                        int totalRows = rows.Count;
                        int totalCols = rows[0].Count;

                        // Iterar sobre cada fila (i representa la fila actual)
                        for (int i = 0; i < totalRows; i++)
                        {
                            var row = rows[i];

                            // // Iterar sobre cada columna (j representa la columna actual)
                            for (int j = 0; j < totalCols; j++)
                            {
                                // // Asignar el valor a la celda (evitar valores nulos)
                                // Si el valor es null, se reemplaza por "NotFound"
                                object value = row[j] ?? "NotFound";
                                // Se escribe en Excel (i+1, j+1) porque Excel usa índices desde 1.
                                worksheet.Cells[i + 1, j + 1].Value = value.ToString();
                            }
                        }

                        // Crear tabla si hay al menos una fila de encabezado
                        if (totalRows > 1 && totalCols > 0)
                        {
                            string rangeAddress = $"A1:{ExcelCellBase.GetAddress(totalRows, totalCols)}";
                            string tableName = $"Table_{sheetName.Replace(" ", "_")}";

                            // Validar nombre antes de crear la tabla
                            if (IsValidExcelTableName(tableName))
                            {
                                try
                                {
                                    var table = worksheet.Tables.Add(worksheet.Cells[rangeAddress], tableName);
                                    table.TableStyle = TableStyles.Medium2;
                                    table.ShowFilter = true;
                                }
                                catch
                                {
                                    // Fallback en caso de excepción interna de EPPlus
                                    invalidTableNames.Add(sheetName);
                                }
                            }
                            else
                            {
                                // Si el nombre no es válido, registrar advertencia
                                invalidTableNames.Add(sheetName);
                            }
                        }
                    }

                    // Mensaje
                    if (invalidTableNames.Count > 0)
                    {
                        MessageBox.Show(
                            "The following sheet names could not be used for Excel Tables, " +
                            "so they were exported without table formatting:\n\n" +
                            string.Join("\n", invalidTableNames),
                            "Warning - Table Names",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                    }

                    // Guardar archivo en memoria
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        // Se guarda el archivo en el stream de memoria
                        package.SaveAs(memoryStream);
                        // Asegurar que los datos sean escritos completamente
                        memoryStream.Flush();

                        // FilePath
                        string tempFilePath =
                            Path.Combine(Path.GetTempPath(), $"TempExcel_{Guid.NewGuid()}.xlsx");

                        // Guardar en memoria
                        File.WriteAllBytes(tempFilePath, memoryStream.ToArray());
                        // Validamos
                        if (File.Exists(tempFilePath))
                        {
                            Process.Start(
                                new ProcessStartInfo(tempFilePath) { UseShellExecute = true }
                            );
                        }
                        else
                        {
                            // Mensaje
                            MessageBox.Show(
                                $"Could not create the Excel file.\nPath: {tempFilePath}",
                                "File Creation Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                    }
                }
            }
            // catch
            catch (Exception ex)
            {
                // Mensaje
                MessageBox.Show(
                    $"Error exporting with EPPlus: {ex.Message}\n{ex.StackTrace}",
                    "Export Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private static bool IsValidExcelTableName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            if (name.Length > 255)
                return false;

            // No puede contener caracteres no válidos
            char[] invalidChars = { '\\', '/', '*', '[', ']', ':', '?' };
            if (name.IndexOfAny(invalidChars) >= 0)
                return false;

            // No puede empezar con un número
            if (char.IsDigit(name[0]))
                return false;

            return true;
        }










    }
}


