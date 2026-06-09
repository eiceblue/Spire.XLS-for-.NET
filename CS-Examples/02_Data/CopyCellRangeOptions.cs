using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using System;
using System.IO;
using System.Windows.Forms;

namespace CopyCellRangeOptions
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

            // Load the workbook from the specified file path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CopyCellRangeOptions.xlsx");

            // Get the reference to the first sheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Configure copy options
            CopyRangeOptions options = CopyRangeOptions.Transpose | CopyRangeOptions.All;

            // Copy ranges to destination with the specified options
            sheet["A1:C4"].Copy(sheet["D2:G3"], options);
            sheet["A1:B5"].Copy(sheet["D5"], options);

            // Specify the output file name for the EXCEL result
            string result = "CopyCellRangeOptions_out.xlsx";

            //Save the Excel file
            workbook.SaveToFile(result, FileFormat.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
