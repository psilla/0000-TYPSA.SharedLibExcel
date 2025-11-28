using System.Collections.Generic;
using System.Globalization;
using System.IO;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ReadCoordFromExcel
    {
        public static Dictionary<string, List<(double X, double Y)>> ReadCoordinatesFromExcel(string filePath)
        {
            var result = new Dictionary<string, List<(double X, double Y)>>();

            // Aseguramos la licencia
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Leemos primera hoja
                var ws = package.Workbook.Worksheets[0];
                int rowCount = ws.Dimension.End.Row;
                // Leemos a partir de fila 2
                for (int row = 2; row <= rowCount; row++) 
                {
                    // Registramos las columnas
                    string xStr = ws.Cells[row, 1].Text?.Trim();
                    string yStr = ws.Cells[row, 2].Text?.Trim();
                    string name = ws.Cells[row, 3].Text?.Trim();
                    // Validamos
                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(xStr) || string.IsNullOrWhiteSpace(yStr)) continue;

                    // Normalizar el separador decimal
                    xStr = xStr.Replace(',', '.');
                    yStr = yStr.Replace(',', '.');
                    // Validamos
                    if (double.TryParse(xStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double x) &&
                        double.TryParse(yStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double y))
                    {
                        // Validamos
                        if (!result.ContainsKey(name))
                            result[name] = new List<(double X, double Y)>();
                        // Almacenamos
                        result[name].Add((x, y));
                    }
                }
            }
            // return
            return result;
        }
    }
}
