using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportToExcel_OpenXml_OneSheet
    {

        public static void ExportToExcel_OpenXml_OneSheet(
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
                    // Creamos una hoja llamada con ese nombre
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Data");

                    // Headers fijos
                    HashSet<string> uniqueHeaders =
                        new HashSet<string> { "FileName", "Handle", "Layer", "ObjectType" };

                    // Dict para agrupar 
                    Dictionary<(string FileName, string Handle), Dictionary<string, string>> groupedData =
                        new Dictionary<(string, string), Dictionary<string, string>>();

                    // Iteramos
                    foreach (var sheetData in data.Values)
                    {
                        // Saltar si no tiene encabezados y valores
                        if (sheetData.Count < 2) continue;

                        // Primera fila contiene encabezados
                        List<string> headers = sheetData[0].Select(h => h.ToString()).ToList();

                        // Filas de datos
                        for (int i = 1; i < sheetData.Count; i++)
                        {
                            List<object> row = sheetData[i];

                            string fileName = row[0]?.ToString() ?? "Unknown";
                            string handle = row[1]?.ToString() ?? "Unknown";
                            string layer = row[2]?.ToString() ?? "Unknown";
                            string objectType = row[3]?.ToString() ?? "Unknown";

                            var key = (fileName, handle);

                            // Si el elemento aún no está en el diccionario, inicializarlo
                            if (!groupedData.ContainsKey(key))
                            {
                                groupedData[key] = new Dictionary<string, string>
                                {
                                    { "FileName", fileName },
                                    { "Handle", handle },
                                    { "Layer", layer },
                                    { "ObjectType", objectType }
                                };
                            }

                            // Agregar valores dinámicos
                            for (int j = 4; j < row.Count; j++)
                            {
                                string propertyName = headers[j];
                                string propertyValue = row[j]?.ToString() ?? "NotFound";

                                uniqueHeaders.Add(propertyName); // Agregar encabezado único
                                groupedData[key][propertyName] = propertyValue;
                            }
                        }
                    }

                    // Convertir dict en tabla Excel
                    List<string> headersList = uniqueHeaders.ToList();
                    int colCount = headersList.Count;

                    // Escribir encabezados en la primera fila
                    for (int col = 0; col < colCount; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = headersList[col];
                    }

                    // Escribir filas de datos
                    int currentRow = 2;
                    foreach (var item in groupedData)
                    {
                        var rowData = item.Value;

                        for (int col = 0; col < colCount; col++)
                        {
                            string header = headersList[col];
                            worksheet.Cells[currentRow, col + 1].Value =
                                rowData.ContainsKey(header) ? rowData[header] : "NotFound";
                        }
                        currentRow++;
                    }

                    // Crear tabla con filtros
                    string tableRange =
                        $"A1:{ExcelCellBase.GetAddress(currentRow - 1, headersList.Count)}";
                    ExcelTable table =
                        worksheet.Tables.Add(worksheet.Cells[tableRange], "ExportedDataTable");
                    table.ShowFilter = true;
                    table.TableStyle = TableStyles.Medium2;

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





    }
}


