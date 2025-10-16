using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Windows.Forms;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_GetExcelSheetNames_OpenXml
    {
        public static List<string> GetExcelSheetNames(string filePath)
        {
            List<string> sheetNames = new List<string>();

            // try
            try
            {
                // Verificar si el archivo existe
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("❌ ERROR: The file does not exist at the specified path.", "File Error");
                    return null;
                }

                FileInfo fileInfo = new FileInfo(filePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // Verificar si hay hojas en el archivo
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        MessageBox.Show("⚠️ The Excel file contains no sheets.", "Warning");
                        return null;
                    }

                    // Agregar nombres de hojas a la lista
                    foreach (var sheet in package.Workbook.Worksheets)
                    {
                        sheetNames.Add(sheet.Name);
                    }
                }

                // Mostrar las hojas encontradas en un MessageBox
                MessageBox.Show(
                    $"✅ Sheets found:\n\n{string.Join("\n", sheetNames.OrderBy(h => h, StringComparer.OrdinalIgnoreCase).ToList())}", "Excel Sheets");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ ERROR while retrieving Excel sheets:\n{ex.Message}", "Error");
                return null;
            }

            // return
            return sheetNames;
        }


    }
}
