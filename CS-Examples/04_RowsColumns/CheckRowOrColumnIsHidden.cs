using Spire.Xls;
using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CheckRowOrColumnIsHidden
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an Excel file into the workbook
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CheckRowOrColumnIsHidden.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a StringBuilder to store the result
            StringBuilder result = new StringBuilder();

            // Specify the row and column index to check
            int rowIndex = 2;
            int columnIndex = 2;

            // Check if the second row is hidden
            bool rowIsHide = sheet.GetRowIsHide(rowIndex);
            if (rowIsHide)
            {
                result.AppendLine("The second row is hidden.");
            }
            else
            {
                result.AppendLine("The second row is not hidden.");
            }

            // Check if the second column is hidden
            bool columnIsHide = sheet.GetColumnIsHide(columnIndex);
            if (columnIsHide)
            {
                result.AppendLine("The second column is hidden.");
            }
            else
            {
                result.AppendLine("The second column is not hidden.");
            }

            // Save the result to a text file
            File.WriteAllText("CheckRowOrColumnIsHidden_result.txt", result.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // View the text file
            FileViewer("CheckRowOrColumnIsHidden_result.txt");
        }
        
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
