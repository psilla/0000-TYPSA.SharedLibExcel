using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_ExportModCheckToExcel_OpenXml
    {
        private static void ApplyBoolColor(
            ExcelRange cell,
            bool value
        )
        {
            cell.Style.Font.Bold = true;

            if (value)
            {
                cell.Style.Font.Color.SetColor(System.Drawing.Color.DarkGreen);
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(
                    System.Drawing.Color.FromArgb(198, 239, 206)
                );
            }
            else
            {
                cell.Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(
                    System.Drawing.Color.FromArgb(255, 199, 206)
                );
            }
        }

        public static string GetSafeSheetName(string name)
        {
            string safe = name.Length > 31 ? name.Substring(0, 31) : name;

            char[] invalid = new[] { '[', ']', '*', '?', '/', '\\', ':' };
            foreach (var ch in invalid)
            {
                safe = safe.Replace(ch.ToString(), "");
            }

            return safe;
        }

        private static string CleanExcelString(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";

            // eliminar caracteres no válidos
            var cleaned = new string(
                input.Where(c => !char.IsControl(c)).ToArray()
            );

            // opcional: limitar longitud
            if (cleaned.Length > 300)
                cleaned = cleaned.Substring(0, 300);

            return cleaned;
        }

        private static void CreateSheetFromObjects<T>(
            ExcelPackage pck,
            string sheetName,
            List<T> data
        )
        {
            // Validamos
            if (data == null || data.Count == 0) return;

            // -----------------------------
            // Split Nombre Hoja
            // -----------------------------

            string baseName = sheetName.Split(':')[0].Trim();

            // -----------------------------
            // Limpiar Nombre Hoja
            // -----------------------------

            string safeSheetName = GetSafeSheetName(baseName);

            // -----------------------------
            // Crear Nombre Hoja
            // -----------------------------

            ExcelWorksheet ws = pck.Workbook.Worksheets.Add(safeSheetName);

            var props = typeof(T).GetProperties();

            // -----------------------------
            // Obtener headers
            // -----------------------------

            for (int c = 0; c < props.Length; c++)
            {
                ws.Cells[1, c + 1].Value = props[c].Name;
            }

            int r = 2;
            // Iteramos
            foreach (var item in data)
            {
                // Iterar propiedades
                for (int c = 0; c < props.Length; c++)
                {
                    // -----------------------------
                    // Obtener valor
                    // -----------------------------

                    var value = props[c].GetValue(item);
                    var cell = ws.Cells[r, c + 1];

                    // try
                    try
                    {
                        if (value == null)
                        {
                            cell.Value = "";
                        }
                        else if (value is string s)
                        {
                            cell.Value = CleanExcelString(s);
                        }
                        else if (
                            value is int ||
                            value is double ||
                            value is bool ||
                            value is DateTime
                        )
                        {
                            cell.Value = value;
                        }
                        else
                        {
                            cell.Value = CleanExcelString(value.ToString());
                        }
                    }
                    // catch
                    catch (Exception ex)
                    {
                        throw new Exception(
                            $"Error in sheet '{sheetName}', row {r}, col {c}, property '{props[c].Name}'\n\n{ex}"
                        );
                    }

                    // -----------------------------
                    // Dar color True/False
                    // -----------------------------

                    if (value != null)
                    {
                        string text = value.ToString().ToLower();

                        if (value is bool b)
                        {
                            ApplyBoolColor(cell, b);
                        }
                        else if (text == "true" || text == "verdadero")
                        {
                            ApplyBoolColor(cell, true);
                        }
                        else if (text == "false" || text == "falso")
                        {
                            ApplyBoolColor(cell, false);
                        }
                    }
                }

                r++;
            }

            // -----------------------------
            // Generar Nombre Tabla
            // -----------------------------

            string tableBase = baseName;
            // limpiar caracteres inválidos
            foreach (char c in new[] { '[', ']', '*', '/', '\\', '?', ':' })
            {
                tableBase = tableBase.Replace(c.ToString(), "");
            }

            // quitar espacios
            tableBase = tableBase.Replace(" ", "");

            // limitar longitud
            if (tableBase.Length > 20)
                tableBase = tableBase.Substring(0, 20);

            string safeTableName = tableBase + "_" + Guid.NewGuid().ToString("N").Substring(0, 6);

            // -----------------------------
            // Crear Tabla
            // -----------------------------

            var range = ws.Cells[1, 1, r - 1, props.Length];
            var table = ws.Tables.Add(range, safeTableName);
            table.ShowHeader = true;
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

            ws.Cells.AutoFitColumns();
        }

        private static void CreateSheetFromMixedObjects(
            ExcelPackage pck,
            string sheetName,
            List<object> data
        )
        {
            if (data == null || data.Count == 0) return;

            string baseName = sheetName.Split(':')[0].Trim();
            string safeSheetName = GetSafeSheetName(baseName);

            ExcelWorksheet ws = pck.Workbook.Worksheets.Add(safeSheetName);

            var props = data
                .Where(x => x != null)
                .SelectMany(x => x.GetType().GetProperties())
                .GroupBy(p => p.Name)
                .Select(g => g.First())
                .ToList();

            for (int c = 0; c < props.Count; c++)
            {
                ws.Cells[1, c + 1].Value = props[c].Name;
            }

            int r = 2;

            foreach (var item in data)
            {
                if (item == null) continue;

                Type itemType = item.GetType();

                for (int c = 0; c < props.Count; c++)
                {
                    string propName = props[c].Name;

                    var itemProp = itemType.GetProperty(propName);
                    var cell = ws.Cells[r, c + 1];

                    object value = itemProp != null
                        ? itemProp.GetValue(item)
                        : null;

                    if (value == null)
                    {
                        cell.Value = "";
                    }
                    else if (value is string s)
                    {
                        cell.Value = CleanExcelString(s);
                    }
                    else if (
                        value is int ||
                        value is double ||
                        value is bool ||
                        value is DateTime
                    )
                    {
                        cell.Value = value;
                    }
                    else
                    {
                        cell.Value = CleanExcelString(value.ToString());
                    }

                    if (value != null)
                    {
                        string text = value.ToString().ToLower();

                        if (value is bool b)
                        {
                            ApplyBoolColor(cell, b);
                        }
                        else if (text == "true" || text == "verdadero")
                        {
                            ApplyBoolColor(cell, true);
                        }
                        else if (text == "false" || text == "falso")
                        {
                            ApplyBoolColor(cell, false);
                        }
                    }
                }

                r++;
            }

            string tableBase = baseName;

            foreach (char c in new[] { '[', ']', '*', '/', '\\', '?', ':' })
            {
                tableBase = tableBase.Replace(c.ToString(), "");
            }

            tableBase = tableBase.Replace(" ", "");

            if (tableBase.Length > 20)
                tableBase = tableBase.Substring(0, 20);

            string safeTableName = tableBase + "_" + Guid.NewGuid().ToString("N").Substring(0, 6);

            var range = ws.Cells[1, 1, r - 1, props.Count];
            var table = ws.Tables.Add(range, safeTableName);
            table.ShowHeader = true;
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

            ws.Cells.AutoFitColumns();
        }

        public static void ExportDataToExcel(
            Dictionary<string, object> exportData
        )
        {
            // try
            try
            {
                // Necesario para EPPlus AutoFitColumns en algunas versiones/runtime
                System.Text.Encoding.RegisterProvider(
                    System.Text.CodePagesEncodingProvider.Instance
                );

                // Activar licencia EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Creamos Excel
                using (var pck = new ExcelPackage())
                {
                    // -----------------------------
                    // Crear hojas
                    // -----------------------------

                    foreach (var kvp in exportData)
                    {
                        // Validamos
                        if (kvp.Value == null) continue;

                        // Convertimos
                        var enumerable = kvp.Value as System.Collections.IEnumerable;

                        // Validamos
                        if (enumerable == null) continue;

                        // -----------------------------
                        // Convertimos enumerable
                        // -----------------------------

                        var data = enumerable.Cast<object>().ToList();

                        // Validamos
                        if (!data.Any()) continue;

                        bool hasMixedTypes = data
                            .Select(x => x.GetType()).Distinct().Count() > 1;
                        // Validamos
                        if (hasMixedTypes)
                        {
                            CreateSheetFromMixedObjects(
                                pck, kvp.Key, data
                            );
                        }
                        else
                        {
                            // -----------------------------
                            // Obtener tipo real
                            // -----------------------------

                            Type itemType = data.First().GetType();

                            // -----------------------------
                            // Crear List<T> REAL
                            // -----------------------------

                            Type listType = typeof(List<>).MakeGenericType(itemType);

                            var typedList = Activator.CreateInstance(listType);

                            var addMethod = listType.GetMethod("Add");

                            foreach (var item in data)
                            {
                                addMethod.Invoke(
                                    typedList,
                                    new[] { item }
                                );
                            }

                            // -----------------------------
                            // Obtener metodo generico
                            // -----------------------------

                            var method = typeof(cls_00_ExportModCheckToExcel_OpenXml).GetMethod(
                                nameof(CreateSheetFromObjects),
                                System.Reflection.BindingFlags.NonPublic |
                                System.Reflection.BindingFlags.Static
                            );

                            if (method == null) continue;

                            // -----------------------------
                            // Construir metodo generico
                            // -----------------------------

                            var genericMethod = method.MakeGenericMethod(itemType);

                            // -----------------------------
                            // Invocar
                            // -----------------------------

                            genericMethod.Invoke(
                                null,
                                new object[]
                                {
                                    pck, kvp.Key, typedList
                                }
                            );
                        }

                    }

                    // -----------------------------
                    // Activamos primera hoja
                    // -----------------------------

                    if (pck.Workbook.Worksheets.Count > 0)
                    {
                        pck.Workbook.Worksheets[0]
                            .View.TabSelected = true;
                    }

                    // -----------------------------
                    // Ruta temporal
                    // -----------------------------

                    var temp = Path.Combine(
                        Path.GetTempPath(), $"{DateTime.Now:yyyyMMdd}_AteneaModelChecker.xlsx"
                    );

                    // -----------------------------
                    // Guardar Excel
                    // -----------------------------

                    File.WriteAllBytes(temp, pck.GetAsByteArray());

                    // -----------------------------
                    // Abrir Excel
                    // -----------------------------

                    Process.Start(
                        new ProcessStartInfo(temp)
                        {
                            UseShellExecute = true
                        }
                    );
                }
            }
            // catch
            catch (Exception ex)
            {
                // Mensaje
                MessageBox.Show(
                    "ERROR EXPORTING EXCEL\n\n" +
                    "MESSAGE:\n" + ex.Message + "\n\n" +
                    "STACK:\n" + ex.StackTrace,
                    "EXPORT ERROR"
                );
            }
        }



    }
}
