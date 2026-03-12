using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_GetExcelSheetNames_OpenXml
    {
        public static List<string> GetExcelSheetNames(string filePath)
        {
            // EPPlus setup
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Verificar si el archivo existe
            if (!File.Exists(filePath))
            {
                // Validamos
                MessageBox.Show(
                    "❌ ERROR: The file does not exist at the specified path.",
                    "File Error"
                );
                // Finalizamos
                return null;
            }
            // try
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    // Validamos hojas
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        // Mensaje
                        MessageBox.Show(
                            "⚠️ The Excel file contains no sheets.",
                            "Warning"
                        );
                        // Finalizamos
                        return null; 
                    }
                    // return
                    return package.Workbook.Worksheets.Select(ws => ws.Name).ToList();
                }
            }
            // catch
            catch (Exception ex)
            {
                // Mensaje
                MessageBox.Show(
                    $"❌ ERROR while retrieving Excel sheets:\n{ex.Message}",
                    "Error"
                );
                // Finalizamos
                return null;
            }
        }



    }
}
