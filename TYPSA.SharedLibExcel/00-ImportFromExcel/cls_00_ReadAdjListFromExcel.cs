using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ReadAdjListFromExcel
    {
        public static Dictionary<string, List<string>> ReadAdjListFromExcel(string excelPath)
        {
            // Inicializamos diccionario de salida
            Dictionary<string, List<string>> adjList = new Dictionary<string, List<string>>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var ws = package.Workbook.Worksheets.FirstOrDefault();
                if (ws == null)
                    throw new Exception("The Excel file does not contain any worksheet.");

                // Buscar encabezados por nombre
                int headerRow = 1;
                int nodeCol = -1;
                int neighborsCol = -1;

                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    string header = ws.Cells[headerRow, col].Text.Trim();

                    if (header.Equals("Node", StringComparison.OrdinalIgnoreCase))
                        nodeCol = col;

                    if (header.Equals("Neighbors", StringComparison.OrdinalIgnoreCase))
                        neighborsCol = col;
                }

                if (nodeCol == -1 || neighborsCol == -1)
                    throw new Exception("Required columns 'Node' and 'Neighbors' not found.");

                // Leer filas
                for (int row = headerRow + 1; row <= ws.Dimension.End.Row; row++)
                {
                    string node = ws.Cells[row, nodeCol].Text.Trim();
                    string neighRaw = ws.Cells[row, neighborsCol].Text.Trim();

                    if (string.IsNullOrWhiteSpace(node))
                        continue;

                    List<string> neighList = new List<string>();

                    if (!string.IsNullOrWhiteSpace(neighRaw))
                    {
                        neighList = neighRaw
                            .Split(',')
                            .Select(s => s.Trim())
                            .Where(s => !string.IsNullOrWhiteSpace(s))
                            .ToList();
                    }

                    adjList[node] = neighList;
                }
            }

            return adjList;
        }


    }
}
