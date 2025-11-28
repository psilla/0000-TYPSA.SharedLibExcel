using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml.Table;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportLabelsToExcel_OpenXml
    {
        public static void ExportLabelsToExcel_OpenXml(
            List<List<string>> etiquetas
        )
        {
            try
            {
                // Necesario en .NET Core/.NET 5+ para EPPlus
                System.Text.Encoding.RegisterProvider(
                    System.Text.CodePagesEncodingProvider.Instance
                );
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Labels");

                    if (etiquetas == null || etiquetas.Count == 0)
                        throw new ArgumentException("No hay datos para exportar.");

                    // Determinar el número máximo de columnas reales en todas las filas
                    int columnCount = etiquetas.Max(e => e.Count);

                    // Generar encabezados dinámicos
                    for (int col = 0; col < columnCount; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = $"Columna {col + 1}";
                        worksheet.Cells[1, col + 1].Style.Font.Bold = true;
                    }

                    // Escribir datos dinámicamente
                    int currentRow = 2;
                    foreach (var sublist in etiquetas)
                    {
                        for (int col = 0; col < sublist.Count; col++)
                        {
                            worksheet.Cells[currentRow, col + 1].Value = sublist[col];
                        }
                        currentRow++;
                    }

                    // Ajustar ancho de columnas automáticamente
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Crear tabla con estilo
                    string tableRange =
                        $"A1:{ExcelCellBase.GetAddress(etiquetas.Count + 1, columnCount)}";
                    ExcelTable table =
                        worksheet.Tables.Add(worksheet.Cells[tableRange], "LabelsTable");
                    table.ShowFilter = true;
                    table.TableStyle = TableStyles.Medium2;

                    // Guardar en archivo temporal
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        package.SaveAs(memoryStream);
                        memoryStream.Flush();

                        string tempFilePath =
                            Path.Combine(Path.GetTempPath(), $"Labels_{Guid.NewGuid()}.xlsx");

                        File.WriteAllBytes(tempFilePath, memoryStream.ToArray());

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
                                $"No se pudo crear el archivo Excel.\nRuta: {tempFilePath}",
                                "Error de creación de archivo",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error al exportar etiquetas con EPPlus: {ex.Message}\n{ex.StackTrace}",
                    "Error de Exportación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }


    }
}
