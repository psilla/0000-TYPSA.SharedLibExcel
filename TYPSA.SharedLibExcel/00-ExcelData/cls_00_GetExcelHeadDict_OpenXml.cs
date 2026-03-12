using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_GetExcelHeadDict_OpenXml
    {
        public static Dictionary<string, List<string>> GetExcelHeadDict(
            string filePath,
            string sheetName,
            int headerRow = 1,
            int firstPsetColumn = 1,
            int firstValueRow = 2
        )
        {
            var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            // try
            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // Obtener hoja
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                    int lastRow = worksheet.Dimension.End.Row;
                    int lastCol = worksheet.Dimension.End.Column;
                    // Recorremos columnas: cada columna representa un PropertySet (encabezado)
                    for (int col = firstPsetColumn; col <= lastCol; col++)
                    {
                        string psetName = worksheet.Cells[headerRow, col].Text?.Trim();
                        // Obviamos Columna sin encabezado 
                        if (string.IsNullOrWhiteSpace(psetName)) continue;

                        List<string> values = new List<string>();
                        // Leer valores tal cual aparecen
                        for (int row = firstValueRow; row <= lastRow; row++)
                        {
                            string val = worksheet.Cells[row, col].Text?.Trim();
                            // Validamos
                            if (string.IsNullOrWhiteSpace(val)) continue;
                            // Evitar duplicados manteniendo primer orden
                            if (!values.Contains(val, StringComparer.OrdinalIgnoreCase))
                                values.Add(val);
                        }
                        // Almacenamos
                        result[psetName] = values;
                    }
                }
                // return
                return result;
            }
            // catch
            catch (Exception ex)
            {
                // Mensaje
                MessageBox.Show(
                    $"❌ ERROR while reading custom order:\n{ex.Message}\n{ex.StackTrace}",
                    "Excel Read Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                // Finalizamos
                return null;
            }
        }

        public static Dictionary<string, List<string>> GetExcelHeadDict_FromPairedColumns(
            string filePath,
            string sheetName,
            int headerRow = 1,
            int firstPsetColumn = 1,
            int firstValueRow = 2,
            int blockWidth = 3
        )
        {
            var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                    int lastRow = worksheet.Dimension.End.Row;
                    int lastCol = worksheet.Dimension.End.Column;

                    // recorrer columnas de 2 en 2
                    for (int col = firstPsetColumn; col <= lastCol; col += blockWidth)
                    {
                        int flagCol = col + 1;
                        if (flagCol > lastCol) break;

                        string psetName = worksheet.Cells[headerRow, col].Text?.Trim();
                        if (string.IsNullOrWhiteSpace(psetName)) continue;

                        List<string> props = new List<string>();

                        for (int row = firstValueRow; row <= lastRow; row++)
                        {
                            string propName = worksheet.Cells[row, col].Text?.Trim();
                            if (string.IsNullOrWhiteSpace(propName)) continue;

                            string flag = worksheet.Cells[row, flagCol].Text?.Trim();

                            bool include =
                                flag.Equals("Yes", StringComparison.OrdinalIgnoreCase) ||
                                flag.Equals("True", StringComparison.OrdinalIgnoreCase) ||
                                flag.Equals("1");

                            if (include && !props.Contains(propName, StringComparer.OrdinalIgnoreCase))
                                props.Add(propName);
                        }

                        result[psetName] = props;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ ERROR while reading custom order:\n{ex.Message}\n{ex.StackTrace}",
                    "Excel Read Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return null;
            }
        }



    }
}
