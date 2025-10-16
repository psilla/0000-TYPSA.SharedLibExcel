using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportReportToExcel
    {
        public static void ExportReportToExcel(
            StringBuilder sb,
            string reportName,
            string carpetaDestino = null
        )
        {
            // try
            try
            {
                if (string.IsNullOrWhiteSpace(carpetaDestino))
                    carpetaDestino = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                if (!Directory.Exists(carpetaDestino))
                    Directory.CreateDirectory(carpetaDestino);

                string nombreArchivo = $"Report_{DateTime.Now:yyyyMMdd_HHmmss}_{reportName}.xlsx";
                string rutaCompleta = Path.Combine(carpetaDestino, nombreArchivo);

                // Inicializar EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Report");

                    // Separar por líneas y escribir cada una en una fila
                    var lineas = sb.ToString().Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                    for (int i = 0; i < lineas.Length; i++)
                    {
                        worksheet.Cells[i + 1, 1].Value = lineas[i];
                    }

                    package.SaveAs(new FileInfo(rutaCompleta));
                }
            }
            // catch
            catch (Exception ex)
            {
                // Mensaje
                MessageBox.Show(
                   $"❌ No se pudo exportar el reporte {reportName}:\n{ex.Message}",
                   "Error de exportación",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error
               );
            }
        }





    }
}
