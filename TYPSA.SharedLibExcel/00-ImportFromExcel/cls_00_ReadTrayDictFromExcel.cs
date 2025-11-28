using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ReadTrayDictFromExcel
    {
        public static Dictionary<string, object> ReadTrayDictFromExcel(string excelPath)
        {
            Dictionary<string, object> trayDict = new Dictionary<string, object>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var ws = package.Workbook.Worksheets.FirstOrDefault();
                if (ws == null)
                    throw new Exception("No worksheets found.");

                // ============================
                // Buscar columnas por encabezado
                // ============================
                int headerRow = 1;
                Dictionary<string, int> cols = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    string header = ws.Cells[headerRow, col].Text.Trim();
                    if (!string.IsNullOrWhiteSpace(header))
                        cols[header] = col;
                }

                // Validar encabezados obligatorios
                string[] requiredHeaders = { "Tag", "Layer", "StartX", "StartY", "EndX", "EndY", "Length", "Handle" };

                foreach (var h in requiredHeaders)
                {
                    if (!cols.ContainsKey(h))
                        throw new Exception($"Header '{h}' not found in Excel.");
                }

                // ============================
                // Leer filas
                // ============================
                for (int row = headerRow + 1; row <= ws.Dimension.End.Row; row++)
                {
                    string handle = ws.Cells[row, cols["Handle"]].Text.Trim();
                    if (string.IsNullOrWhiteSpace(handle))
                        continue;

                    Dictionary<string, object> rowDict = new Dictionary<string, object>();

                    foreach (string h in requiredHeaders)
                    {
                        string raw = ws.Cells[row, cols[h]].Text.Trim();

                        // Intentar parsear numérico
                        if (double.TryParse(raw.Replace(",", "."), System.Globalization.NumberStyles.Any,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            out double num))
                        {
                            rowDict[h] = num;
                        }
                        else
                        {
                            rowDict[h] = raw; // string
                        }
                    }

                    // Insertar en diccionario principal por HANDLE
                    trayDict[handle] = rowDict;
                }
            }
            // return
            return trayDict;
        }


    }
}
