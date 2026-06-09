using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using System;
using System.IO;
using System.Windows.Forms;

namespace InsertEmbedCheckBox
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

            // Get the cell range at position A1
            XlsRange range = sheet.Range["A1"];

            // Insert an embedded checkbox into the specified range
            range.InsertEmbedCheckBox();

            // Set the checkbox state to checked (true = checked, false = unchecked)
            range.SetEmbedCheckBoxCheckState(true);

            // Specify the output file name for the EXCEL result
            string result = "InsertEmbedCheckBox_out.xlsx";

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
