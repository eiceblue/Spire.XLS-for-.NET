using Spire.Xls;
using System;
using System.Windows.Forms;

namespace InsertImageInWPSCell
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

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Embed an image into the first cell
              worksheet.Cells[0].InsertOrUpdateCellImage(@"..\..\..\..\..\..\Data\SpireXls.png", true);
        
            // Specify the filename for the resulting Excel file
            String result = "output.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object
            workbook.Dispose();
         
            // View the document using a file viewer
            FileViewer(result);

            this.Close();
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
