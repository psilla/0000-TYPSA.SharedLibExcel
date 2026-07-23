using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportFilteredParamCheckToExcel
    {
        public static void ExportFilteredParamCheckToExcel(
            List<Dictionary<string, object>> dataJsonByModelFiltered,
            List<string> headers,
            string fileNameKey,
            string civilParamCheckKey,
            string psetNameKey,
            string psetDataKey,
            string propNameKey,
            string dataTypeKey
        )
        {
            try
            {
                // Necesario para EPPlus AutoFitColumns en algunas versiones/runtime
                System.Text.Encoding.RegisterProvider(
                    System.Text.CodePagesEncodingProvider.Instance
                );

                // EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Validamos
                if (dataJsonByModelFiltered == null || dataJsonByModelFiltered.Count == 0)
                {
                    return;
                }

                // Archivo temporal
                string excelPath = Path.Combine(
                    Path.GetTempPath(), $"ParamCheckExport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                );

                FileInfo fileInfo = new FileInfo(excelPath);

                // Eliminamos si existe
                if (fileInfo.Exists)
                {
                    fileInfo.Delete();
                }

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add("ParamCheck");

                    // -----------------------------
                    // Headers
                    // -----------------------------

                    for (int col = 0; col < headers.Count; col++)
                    {
                        ws.Cells[1, col + 1].Value = headers[col];
                    }

                    int row = 2;

                    // -----------------------------
                    // Datos
                    // -----------------------------

                    foreach (Dictionary<string, object> fileData in dataJsonByModelFiltered)
                    {
                        string fileName = fileData.ContainsKey(fileNameKey) && fileData[fileNameKey] != null
                            ? fileData[fileNameKey].ToString()
                            : "Unknown";
                        // Validamos
                        if (!fileData.ContainsKey(civilParamCheckKey) || fileData[civilParamCheckKey] == null)
                        {
                            continue;
                        }

                        List<Dictionary<string, object>> psets = fileData[civilParamCheckKey] as List<Dictionary<string, object>>;
                        // Validamos
                        if (psets == null)
                        {
                            continue;
                        }

                        // Iteramos
                        foreach (Dictionary<string, object> pset in psets)
                        {
                            string psetName = pset.ContainsKey(psetNameKey) && pset[psetNameKey] != null
                                ? pset[psetNameKey].ToString()
                                : "Unknown";
                            // Validamos
                            if (!pset.ContainsKey(psetDataKey) || pset[psetDataKey] == null)
                            {
                                continue;
                            }

                            // Obtenemos
                            List<Dictionary<string, object>> psetData = pset[psetDataKey] as List<Dictionary<string, object>>;
                            // Validamos
                            if (psetData == null)
                            {
                                continue;
                            }

                            // Iteramos
                            foreach (Dictionary<string, object> propData in psetData)
                            {
                                string propName = propData.ContainsKey(propNameKey) && propData[propNameKey] != null
                                    ? propData[propNameKey].ToString()
                                    : "Unknown";

                                string dataType = propData.ContainsKey(dataTypeKey) && propData[dataTypeKey] != null
                                    ? propData[dataTypeKey].ToString()
                                    : "Unknown";

                                List<object> values = new List<object>
                            {
                                fileName, psetName, propName, dataType
                            };
                                // Iteramos
                                for (int col = 0; col < values.Count; col++)
                                {
                                    ws.Cells[row, col + 1].Value = values[col];
                                }

                                row++;
                            }
                        }
                    }

                    // -----------------------------
                    // Crear tabla
                    // -----------------------------

                    if (row > 2)
                    {
                        ExcelRange tableRange = ws.Cells[1, 1, row - 1, 4];

                        ExcelTable table = ws.Tables.Add(
                            tableRange,
                            "ParamCheckTable"
                        );

                        // Estilo de tabla
                        table.TableStyle = TableStyles.Light1;

                        // Mostrar filtros
                        table.ShowFilter = true;
                    }

                    // Encabezados en negrita
                    ws.Cells[1, 1, 1, 4].Style.Font.Bold = true;

                    // -----------------------------
                    // Formato
                    // -----------------------------

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    ws.View.FreezePanes(2, 1);

                    // Guardamos
                    package.Save();
                }

                // -----------------------------
                // Abrir Excel
                // -----------------------------

                Process.Start(new ProcessStartInfo
                {
                    FileName = excelPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    ex.ToString(),
                    "Error exporting Param Check",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }


    }
}
