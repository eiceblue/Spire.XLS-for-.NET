using Spire.Xls;
using System;
using System.Windows.Forms;

namespace OpenExistingFile
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

            // Load an existing Excel file from the specified path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\templateAz2.xlsx");

            // Add a new sheet with the name "MySheet"
            Worksheet sheet = workbook.Worksheets.Add("MySheet");

            // Set the value of cell A1 to "Hello World"
            sheet.Range["A1"].Text = "Hello World";

            // Specify the name for the resulting Excel file
            String result = "OpenExistingFile_result.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object
            workbook.Dispose();

            // Launch the document
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
