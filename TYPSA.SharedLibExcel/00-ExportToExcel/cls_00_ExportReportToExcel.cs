using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportReportToExcel
    {
        public static void SaveExcelReport(
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
                   $"❌ Failed to export the report {reportName}:\n{ex.Message}",
                   "Error de exportación",
                   MessageBoxButtons.OK, MessageBoxIcon.Error
               );
            }
        }

        private static string GetExcelReportFullPath(
            string projectCode,
            string reportName,
            string rootFolderName = "AztecHrAutomations",
            bool exportToDesktop = true,
            string customBasePath = null
        )
        {
            string basePath;
            // True
            if (exportToDesktop)
            {
                // Ruta del Escritorio
                basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            // False
            else
            {   // Validamos
                if (string.IsNullOrWhiteSpace(customBasePath))
                    throw new ArgumentException("Custom base path must be provided when exportToDesktop is false.");
                // Invertimos
                basePath = customBasePath;
                // Validamos
                if (!Directory.Exists(basePath))
                    throw new DirectoryNotFoundException(
                        $"Base path does not exist: {basePath}"
                    );
            }

            // Carpeta raíz
            string rootFolderPath = Path.Combine(basePath, rootFolderName);
            // Validamos
            if (!Directory.Exists(rootFolderPath))
            {
                // Creamos carpeta
                Directory.CreateDirectory(rootFolderPath);
            }
                
            // Carpeta del proyecto
            string projectFolderPath = Path.Combine(rootFolderPath, projectCode);
            // Validamos
            if (!Directory.Exists(projectFolderPath))
            {
                // Creamos carpeta
                Directory.CreateDirectory(projectFolderPath);
            }
                
            // Nombre del archivo
            string fileName = $"{reportName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            // Full path final
            return Path.Combine(projectFolderPath, fileName);
        }

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Sheet";

            // Caracteres no permitidos en Excel
            char[] invalidChars = { '[', ']', '*', '?', '/', '\\', ':' };

            foreach (char c in invalidChars)
                name = name.Replace(c.ToString(), "");

            // Máx 31 caracteres
            return name.Length > 31
                ? name.Substring(0, 31)
                : name;
        }

        public static void SaveExcelReportInFolder(
            StringBuilder sb,
            string projectCode,
            string reportName,
            string rootFolderName = "ByDefault",
            bool exportToDesktop = true,
            string customBasePath = null
        )
        {
            // try
            try
            {
                // Obtenemos ruta
                string fullPath = GetExcelReportFullPath(
                    projectCode, reportName, rootFolderName, exportToDesktop, customBasePath
                );

                // Inicializar EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Report");

                    // Separar por líneas y escribir cada línea en una fila
                    var lines = sb.ToString().Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                    for (int i = 0; i < lines.Length; i++)
                    {
                        worksheet.Cells[i + 1, 1].Value = lines[i];
                    }

                    // Ajuste automático de columna
                    worksheet.Column(1).AutoFit();

                    package.SaveAs(new FileInfo(fullPath));
                }
            }
            // catch
            catch (Exception ex)
            {
                // Mensaje
                MessageBox.Show(
                   $"❌ Failed to export the report {reportName}:\n{ex.Message}",
                   "Error de exportación",
                   MessageBoxButtons.OK, MessageBoxIcon.Error
               );
            }
        }

        private static void AddRegionReportWorksheet(
            ExcelPackage package,
            string rawSheetName,
            StringBuilder content
        )
        {
            // Excel constraints
            string sheetName = SanitizeSheetName(rawSheetName);

            // Crear hoja
            var ws = package.Workbook.Worksheets.Add(sheetName);

            // Volcar contenido línea a línea
            int row = 1;
            foreach (string line in content
                .ToString()
                .Split(new[] { Environment.NewLine }, StringSplitOptions.None))
            {
                ws.Cells[row, 1].Value = line;
                row++;
            }

            // Auto-ajuste
            ws.Column(1).AutoFit();
        }

        public static void SaveExcelReportInFolder_BySheet(
            Dictionary<string, StringBuilder> regionReports,
            string projectCode,
            string reportName,
            string rootFolderName = "AztecHrAutomations",
            bool exportToDesktop = true,
            string customBasePath = null
        )
        {
            // try
            try 
            {
                // Obtenemos ruta
                string fullPath = GetExcelReportFullPath(
                    projectCode, reportName, rootFolderName, exportToDesktop, customBasePath
                );

                // Inicializar EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    // Iteramos
                    foreach (var kvp in regionReports)
                    {
                        // Procesamos hoja
                        AddRegionReportWorksheet(package, kvp.Key, kvp.Value);
                    }

                    // Guardar archivo
                    File.WriteAllBytes(fullPath, package.GetAsByteArray());
                }
            }
            // catch
            catch (Exception ex)
            {
                // Mensaje
                MessageBox.Show(
                    $"❌ Failed to export Excel report:\n\n{ex.Message}",
                    "Excel Export Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        public static void SaveExcelReportInFolder_BySheet_Report(
            Dictionary<string, StringBuilder> regionReports,
            List<Dictionary<string, List<List<object>>>> exportDataList,
            string projectCode,
            string reportName,
            string rootFolderName = "AztecHrAutomations",
            bool exportToDesktop = true,
            string customBasePath = null
        )
        {
            try
            {
                string fullPath = GetExcelReportFullPath(
                    projectCode, reportName, rootFolderName, exportToDesktop, customBasePath
                );

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    // =========================
                    // HOJA REPORT → SOLO exportData
                    // =========================

                    var wsReport = package.Workbook.Worksheets.Add("Report");

                    int rowCursor = 1;

                    if (exportDataList != null)
                    {
                        foreach (var exportData in exportDataList)
                        {
                            if (exportData == null) continue;

                            foreach (var entry in exportData)
                            {
                                // =========================
                                // Título PropertySet
                                // =========================

                                wsReport.Cells[rowCursor, 1].Value = entry.Key;
                                wsReport.Cells[rowCursor, 1].Style.Font.Bold = true;
                                rowCursor++;

                                var rows = entry.Value;
                                if (rows == null || rows.Count == 0) continue;

                                int totalRows = rows.Count;
                                int totalCols = rows[0].Count;

                                int tableStartRow = rowCursor;

                                // =========================
                                // Volcar datos
                                // =========================

                                int writeRowOffset = 0;

                                for (int i = 0; i < totalRows; i++)
                                {
                                    var row = rows[i];

                                    for (int j = 0; j < totalCols; j++)
                                    {
                                        var cellValue = row[j];

                                        var cell = wsReport.Cells[rowCursor + writeRowOffset, j + 1];

                                        cell.Value = cellValue?.ToString() ?? "NotFound";

                                        // Colorear según valor True / False
                                        bool? boolValue = null;

                                        if (cellValue is bool b)
                                        {
                                            boolValue = b;
                                        }
                                        else if (cellValue is string s)
                                        {
                                            if (string.Equals(s, "True", StringComparison.OrdinalIgnoreCase))
                                                boolValue = true;
                                            else if (string.Equals(s, "False", StringComparison.OrdinalIgnoreCase))
                                                boolValue = false;
                                        }

                                        if (boolValue.HasValue)
                                        {
                                            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                                            if (boolValue.Value)
                                                cell.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                                            else
                                                cell.Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                                        }
                                    }

                                    writeRowOffset++;
                                }

                                int tableEndRow = rowCursor + writeRowOffset - 1;

                                // =========================
                                // Crear tabla EPPlus
                                // =========================

                                string rangeAddress =
                                    $"A{tableStartRow}:{ExcelCellBase.GetAddress(tableEndRow, totalCols)}";

                                string tableName =
                                    $"ReportTable_{Guid.NewGuid().ToString("N").Substring(0, 8)}";

                                var table = wsReport.Tables.Add(wsReport.Cells[rangeAddress], tableName);
                                table.TableStyle = TableStyles.Medium2;
                                table.ShowFilter = true;

                                rowCursor = tableEndRow + 3;
                            }
                        }
                    }

                    wsReport.Cells.AutoFitColumns();

                    // =========================
                    // HOJAS regionReports 
                    // =========================

                    foreach (var kvp in regionReports)
                    {
                        AddRegionReportWorksheet(package, kvp.Key, kvp.Value);
                    }

                    File.WriteAllBytes(fullPath, package.GetAsByteArray());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Failed to export Excel report:\n\n{ex.Message}",
                    "Excel Export Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
        
        public static void SaveExcelReportInFolder_BySheet_Report_Count(
            Dictionary<string, StringBuilder> regionReports,
            List<Dictionary<string, List<List<object>>>> exportDataList,
            StringBuilder sbAmpacityReport,
            string projectCode,
            string reportName,
            string rootFolderName = "AztecHrAutomations",
            bool exportToDesktop = true,
            string customBasePath = null
        )
        {
            try
            {
                string fullPath = GetExcelReportFullPath(
                    projectCode, reportName, rootFolderName, exportToDesktop, customBasePath
                );

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    // -----------------------------=============================
                    // Hoja 1: Report
                    // -----------------------------=============================

                    var wsReport = package.Workbook.Worksheets.Add("Report");

                    int rowCursor = 1;

                    if (exportDataList != null)
                    {
                        foreach (var exportData in exportDataList)
                        {
                            if (exportData == null) continue;

                            foreach (var entry in exportData)
                            {
                                // Título PropertySet
                                wsReport.Cells[rowCursor, 1].Value = entry.Key;
                                wsReport.Cells[rowCursor, 1].Style.Font.Bold = true;
                                rowCursor++;

                                var rows = entry.Value;
                                if (rows == null || rows.Count == 0) continue;

                                int totalRows = rows.Count;
                                int totalCols = rows[0].Count;

                                // Fila para totales
                                int totalsRow = rowCursor;

                                // Fila donde empieza tabla (encabezado incluido en datos)
                                int tableStartRow = rowCursor + 1;

                                int[] falseCounts = new int[totalCols];

                                // Volcar datos
                                for (int i = 0; i < totalRows; i++)
                                {
                                    var row = rows[i];

                                    for (int j = 0; j < totalCols; j++)
                                    {
                                        var cellValue = row[j];

                                        var cell = wsReport.Cells[tableStartRow + i, j + 1];

                                        cell.Value = cellValue?.ToString() ?? "NotFound";

                                        bool? boolValue = null;

                                        if (cellValue is bool b)
                                            boolValue = b;
                                        else if (cellValue is string s)
                                        {
                                            if (string.Equals(s, "True", StringComparison.OrdinalIgnoreCase))
                                                boolValue = true;
                                            else if (string.Equals(s, "False", StringComparison.OrdinalIgnoreCase))
                                                boolValue = false;
                                        }

                                        if (boolValue.HasValue)
                                        {
                                            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                                            if (boolValue.Value)
                                            {
                                                cell.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                                            }
                                            else
                                            {
                                                cell.Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                                                falseCounts[j]++;
                                            }
                                        }
                                    }
                                }

                                // Escribir conteo de FALSE encima del header
                                for (int j = 0; j < totalCols; j++)
                                {
                                    if (falseCounts[j] > 0)
                                    {
                                        var cell = wsReport.Cells[totalsRow, j + 1];
                                        cell.Value = $"{falseCounts[j]}";
                                        cell.Style.Font.Bold = true;
                                    }
                                }

                                int tableEndRow = tableStartRow + totalRows - 1;

                                string rangeAddress =
                                    $"A{tableStartRow}:{ExcelCellBase.GetAddress(tableEndRow, totalCols)}";

                                string tableName =
                                    $"ReportTable_{Guid.NewGuid().ToString("N").Substring(0, 8)}";

                                var table = wsReport.Tables.Add(wsReport.Cells[rangeAddress], tableName);
                                table.TableStyle = TableStyles.Medium2;
                                table.ShowFilter = true;

                                rowCursor = tableEndRow + 3;
                            }
                        }
                    }

                    wsReport.Cells.AutoFitColumns();

                    // -----------------------------=============================
                    // Hoja 2: Ampacity Report
                    // -----------------------------=============================

                    if (sbAmpacityReport != null && sbAmpacityReport.Length > 0)
                    {
                        var wsAmp = package.Workbook.Worksheets.Add("Ampacity Report");

                        int row = 1;
                        // Contenido
                        string[] lines = sbAmpacityReport.ToString()
                            .Split(new[] { Environment.NewLine }, StringSplitOptions.None);

                        foreach (var line in lines)
                        {
                            wsAmp.Cells[row, 1].Value = line;
                            row++;
                        }

                        wsAmp.Cells.AutoFitColumns();
                    }

                    // -----------------------------=============================
                    // Hojas 3+: Report by Skid
                    // -----------------------------=============================

                    if (regionReports != null)
                    {
                        foreach (var kvp in regionReports)
                        {
                            AddRegionReportWorksheet(package, kvp.Key, kvp.Value);
                        }
                    }

                    // -----------------------------=============================
                    // Guardar
                    // -----------------------------=============================

                    File.WriteAllBytes(fullPath, package.GetAsByteArray());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Failed to export Excel report:\n\n{ex.Message}",
                    "Excel Export Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

















    }
}
