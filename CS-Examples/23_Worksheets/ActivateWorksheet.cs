using System;
using System.Windows.Forms;
using Spire.Xls;

namespace ActivateWorksheet
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

            // Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorksheetSample2.xlsx");

            // Get the second worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[1];

            // Activate the sheet
            sheet.Activate();

            // Specify the output filename for the workbook
            string output = "ActivateWorksheet.xlsx";

            // Save the modified workbook to a file
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
            ExcelDocViewer(output);
		}
        private void ExcelDocViewer(string fileName)
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
