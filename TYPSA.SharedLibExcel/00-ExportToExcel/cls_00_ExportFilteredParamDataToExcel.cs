using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportFilteredParamDataToExcel
    {
        public static void ExportParamDataToExcel(
            Dictionary<string, object> dictDataByFileToJson,
            List<string> headers,
            string dataByFileNameKey,
            string fileNameKey,
            string civilParamDataKey,
            string handleKey,
            string layerKey,
            string objectTypeKey,
            string propertySetInfoKey,
            string parametersKey,
            string propNameKey,
            string propValueKey
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

                // -----------------------------
                // Validar diccionario principal
                // -----------------------------

                if (dictDataByFileToJson == null || dictDataByFileToJson.Count == 0 ||
                    !dictDataByFileToJson.ContainsKey(dataByFileNameKey) || dictDataByFileToJson[dataByFileNameKey] == null
                ) return;
                

                // Obtenemos los datos por archivo
                List<Dictionary<string, object>> dataByFileName = 
                    dictDataByFileToJson[dataByFileNameKey] as List<Dictionary<string, object>>;
                // Validamos
                if (dataByFileName == null || dataByFileName.Count == 0) return;
               

                // -----------------------------
                // Obtener todos los PropName
                // -----------------------------

                HashSet<string> propertyNamesSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                // Iteramos
                foreach (Dictionary<string, object> fileData in dataByFileName)
                {
                    // Validamos
                    if (!fileData.ContainsKey(civilParamDataKey) || fileData[civilParamDataKey] == null) continue;
                    
                    List<Dictionary<string, object>> civilParamData = 
                        fileData[civilParamDataKey] as List<Dictionary<string, object>>;
                    // Validamos
                    if (civilParamData == null) continue;
                   
                    // Iteramos
                    foreach (Dictionary<string, object> entityData in civilParamData)
                    {
                        // Validamos
                        if (!entityData.ContainsKey(propertySetInfoKey) || entityData[propertySetInfoKey] == null) continue;
                      
                        List<Dictionary<string, object>> propertySetInfo = 
                            entityData[propertySetInfoKey] as List<Dictionary<string, object>>;
                        // Validamos
                        if (propertySetInfo == null) continue;
                        
                        // Iteramos
                        foreach (Dictionary<string, object> psetData in propertySetInfo)
                        {
                            // Validamos
                            if (!psetData.ContainsKey(parametersKey) || psetData[parametersKey] == null) continue;
                           
                            List<Dictionary<string, object>> parameters = psetData[parametersKey] as List<Dictionary<string, object>>;
                            // Validamos
                            if (parameters == null) continue;
                           
                            // Iteramos
                            foreach (Dictionary<string, object> parameter in parameters)
                            {
                                string propName = parameter.ContainsKey(propNameKey) && parameter[propNameKey] != null
                                    ? parameter[propNameKey].ToString()
                                    : null;
                                // Validamos
                                if (!string.IsNullOrWhiteSpace(propName))
                                {
                                    propertyNamesSet.Add(propName);
                                }
                            }
                        }
                    }
                }

                // Ordenamos
                List<string> propertyNames = propertyNamesSet.OrderBy(x => x).ToList();

                // -----------------------------
                // Preparar encabezados finales
                // -----------------------------

                // Creamos una copia para no modificar la lista recibida
                List<string> finalHeaders = new List<string>(headers.Take(4));

                // Añadimos los encabezados a las propiedades
                finalHeaders.AddRange(propertyNames);

                // -----------------------------
                // Archivo temporal
                // -----------------------------

                string excelPath = Path.Combine(
                    Path.GetTempPath(),
                    $"ParamDataExport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                );

                FileInfo fileInfo = new FileInfo(excelPath);

                // Eliminamos si existe
                if (fileInfo.Exists)
                {
                    fileInfo.Delete();
                }

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add("ParamData");

                    // -----------------------------
                    // Headers
                    // -----------------------------

                    for (int col = 0; col < finalHeaders.Count; col++)
                    {
                        ws.Cells[1, col + 1].Value = finalHeaders[col];
                    }

                    int row = 2;

                    // -----------------------------
                    // Datos
                    // -----------------------------

                    foreach (Dictionary<string, object> fileData in dataByFileName)
                    {
                        string fileName = fileData.ContainsKey(fileNameKey) && fileData[fileNameKey] != null
                            ? fileData[fileNameKey].ToString()
                            : "Unknown";
                        // Validamos
                        if (!fileData.ContainsKey(civilParamDataKey) || fileData[civilParamDataKey] == null) continue;

                        List<Dictionary<string, object>> civilParamData = 
                            fileData[civilParamDataKey] as List<Dictionary<string, object>>;
                        // Validamos
                        if (civilParamData == null) continue;
                        
                        // Iteramos
                        foreach (Dictionary<string, object> entityData in civilParamData)
                        {
                            string handle = entityData.ContainsKey(handleKey) && entityData[handleKey] != null
                                ? entityData[handleKey].ToString()
                                : "Unknown";

                            string layer = entityData.ContainsKey(layerKey) && entityData[layerKey] != null
                                ? entityData[layerKey].ToString()
                                : "Unknown";

                            string objectType = entityData.ContainsKey(objectTypeKey) && entityData[objectTypeKey] != null
                                ? entityData[objectTypeKey].ToString()
                                : "Unknown";

                            Dictionary<string, object> propertyValues = new Dictionary<string, object>(
                                StringComparer.OrdinalIgnoreCase
                            );

                            // -----------------------------
                            // Obtener valores de propiedades
                            // -----------------------------

                            if (entityData.ContainsKey(propertySetInfoKey) && entityData[propertySetInfoKey] != null)
                            {
                                List<Dictionary<string, object>> propertySetInfo = 
                                    entityData[propertySetInfoKey] as List<Dictionary<string, object>>;
                                // Validamos
                                if (propertySetInfo != null)
                                {
                                    // Iteramos
                                    foreach (Dictionary<string, object> psetData in propertySetInfo
                                    )
                                    {
                                        // Validamos
                                        if (!psetData.ContainsKey(parametersKey) || psetData[parametersKey] == null) continue;
                                       
                                        List<Dictionary<string, object>> parameters = 
                                            psetData[parametersKey] as List<Dictionary<string, object>>;
                                        // Validamos
                                        if (parameters == null) continue;
                                        
                                        // Iteramos
                                        foreach (Dictionary<string, object> parameter in parameters)
                                        {
                                            string propName = parameter.ContainsKey(propNameKey) && parameter[propNameKey] != null
                                                ? parameter[propNameKey].ToString()
                                                : null;
                                            // Validamos
                                            if (string.IsNullOrWhiteSpace(propName)) continue;
                                           
                                            object propValue = parameter.ContainsKey(propValueKey)
                                                ? parameter[propValueKey]
                                                : null;

                                            propertyValues[propName] = propValue;
                                        }
                                    }
                                }
                            }

                            // -----------------------------
                            // Escribir fila
                            // -----------------------------

                            int col = 1;

                            ws.Cells[row, col++].Value = fileName;
                            ws.Cells[row, col++].Value = handle;
                            ws.Cells[row, col++].Value = layer;
                            ws.Cells[row, col++].Value = objectType;

                            // A partir del quinto encabezado se escriben las propiedades
                            foreach (string propertyName in propertyNames)
                            {
                                object propertyValue = null;

                                propertyValues.TryGetValue(propertyName, out propertyValue);

                                ws.Cells[row, col++].Value = propertyValue ?? string.Empty;
                            }

                            row++;
                        }
                    }

                    // -----------------------------
                    // Crear tabla
                    // -----------------------------

                    if (row > 2)
                    {
                        ExcelRange tableRange = ws.Cells[1, 1, row - 1, finalHeaders.Count];

                        ExcelTable table = ws.Tables.Add( tableRange, "ParamDataTable");

                        // Estilo de tabla
                        table.TableStyle = TableStyles.Light1;
                        // Mostrar filtros
                        table.ShowFilter = true;
                    }

                    // Encabezados en negrita
                    ws.Cells[1, 1, 1, finalHeaders.Count].Style.Font.Bold = true;

                    // -----------------------------
                    // Formato
                    // -----------------------------

                    if (ws.Dimension != null)
                    {
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    }

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
                    "Error exporting Param Data",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }

        



    }
}
