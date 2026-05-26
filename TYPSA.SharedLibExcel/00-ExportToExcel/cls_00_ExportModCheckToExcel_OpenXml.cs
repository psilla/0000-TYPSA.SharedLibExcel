using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TYPSA.SharedLib.Excel
{
    public static class AteneaModelCheckerOptionsLocalized
    {
        public static string ProjectUnits(bool es) =>
            es
                ? "Unidades del proyecto: Comprueba las unidades del dibujo (longitud, ángulos, etc.)"
                : "Project Units: Checks drawing units (length, angles, etc.)";

        public static string LayersInUse(bool es) =>
            es
                ? "Capas en uso: Analiza qué capas están en uso y cuántas entidades contienen"
                : "Layers in Use: Analyzes which layers are used and entity count per layer";

        public static string LayerZero(bool es) =>
            es
                ? "Entidades en capa 0: Detecta entidades dibujadas en la capa 0"
                : "Entities in Layer 0: Detects entities drawn on layer 0";

        public static string Version(bool es) =>
            es
                ? "Versión del software: Obtiene la versión de AutoCAD del archivo"
                : "Software Version: Retrieves AutoCAD file version";

        public static string Xrefs(bool es) =>
            es
                ? "Referencias externas: Lista las referencias externas y su estado"
                : "External References: Lists external references and their status";

        //public static string CoordSystem(bool es) =>
        //    es
        //        ? "Sistema de coordenadas: Obtiene el sistema de coordenadas del dibujo"
        //        : "Coordinates System: Retrieves drawing coordinate system";

        public static string PaperTextFont(bool es) =>
            es
                ? "Fuente de texto en layout: Analiza las fuentes de texto utilizadas en layouts"
                : "Text Font in Layout: Analyzes text fonts used in layouts";

        public static string ByLayerProperties(bool es) =>
            es
                ? "Propiedades por capa: Verifica si color, tipo de línea y grosor están por capa"
                : "ByLayer Properties: Checks if color, linetype and lineweight are ByLayer";

        public static string RevisionClouds(bool es) =>
            es
                ? "Nubes de revisión: Detecta nubes de revisión en layouts"
                : "Revision Clouds: Detects revision clouds in layouts";

        public static string BlockAttributes(bool es) =>
            es
                ? "Atributos de bloques: Extrae atributos de bloques en layouts (title blocks, etc.)"
                : "Block Attributes: Extracts block attributes in layouts (title blocks, etc.)";

        public static string PlotTag(bool es) =>
            es
                ? "Ruta de impresión: Extrae información de Ritningsfil, Plottdatum y Plottad av"
                : "Plot Information: Extracts Ritningsfil, Plottdatum and Plottad av information";

        public static List<string> GetAllOptions(bool isSpanish)
        {
            return new List<string>
                {
                    ProjectUnits(isSpanish),
                    LayersInUse(isSpanish),
                    LayerZero(isSpanish),
                    Version(isSpanish),
                    Xrefs(isSpanish),
                    //CoordSystem(isSpanish),
                    PaperTextFont(isSpanish),
                    ByLayerProperties(isSpanish),
                    RevisionClouds(isSpanish),
                    BlockAttributes(isSpanish),
                    PlotTag(isSpanish)
                };
        }

        public static List<string> GetDefaultSelectedOptions(bool isSpanish)
        {
            return GetAllOptions(isSpanish);
        }
    }

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

        private static string GetSafeSheetName(string name)
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

        public static void ExportDataToExcel(
            Dictionary<string, object> exportData
        )
        {
            // try
            try
            {
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

                        // Añadimos elementos
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

                        //var method = typeof(cls_00_ExportModCheckToExcel_OpenXml).GetMethod(
                        //    "CreateSheetFromObjects",
                        //    System.Reflection.BindingFlags.NonPublic |
                        //    System.Reflection.BindingFlags.Static
                        //);
                        var method = typeof(cls_00_ExportModCheckToExcel_OpenXml).GetMethod(
                            nameof(CreateSheetFromObjects),
                            System.Reflection.BindingFlags.NonPublic |
                            System.Reflection.BindingFlags.Static
                        );

                        // Validamos
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
