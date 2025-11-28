using System.Windows.Forms;
using TYPSA.SharedLib.UserForms;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_SelectExcelDirectory
    {
        public static string SelectExcelDirectory()
        {
            using (var form = new ExcelPathEntry())
            {
                // Validamos
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // return
                    return form.ExcelPath;
                }
                // En caso de no validar
                else
                {
                    // Mensaje
                    MessageBox.Show(
                        "No directory path provided. Process aborted.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error
                    );
                    // Finalizamos
                    return null;
                }
            }
        }




    }
}
