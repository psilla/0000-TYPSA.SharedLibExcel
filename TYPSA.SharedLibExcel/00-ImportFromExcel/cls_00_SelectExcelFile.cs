using System.Windows.Forms;

namespace TYPSA.SharedLib.Excel
{
    public class cls_00_SelectExcelFile
    {
        public static string SelectExcelFile(string initialDirectory)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = initialDirectory,
                Title = "Select Excel File to Analyze",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                Multiselect = false
            })
            {
                // Validamos
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // return
                    return openFileDialog.FileName;
                }
                // En caso de no validar
                else
                {
                    // Mensaje
                    MessageBox.Show(
                        "No Excel file selected. Process aborted.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error
                    );
                    // Finalizamos
                    return null;
                }
            }
        }


    }
}
