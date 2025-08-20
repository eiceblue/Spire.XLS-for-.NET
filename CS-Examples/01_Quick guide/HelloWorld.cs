using Spire.Xls;
using System;
using System.Windows.Forms;

namespace HelloWorld
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Get the reference to the first sheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set the cell value of cell A1 to "Hello World"
            sheet.Range["A1"].Text = "Hello World";

            // Auto-fit the columns to adjust their width based on the content
            sheet.Range["A1"].AutoFitColumns();

            // Specify the filename for the resulting Excel file
            String result = "HelloWorld.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object
            workbook.Dispose();

            // View the document using a file viewer
            FileViewer(result);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
