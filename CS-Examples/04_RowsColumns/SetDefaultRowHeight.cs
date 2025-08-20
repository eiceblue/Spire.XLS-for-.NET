using System;
using System.Windows.Forms;
using Spire.Xls;

namespace SetDefaultRowHeight
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

            // Load an existing document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CommonTemplate.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set the default row height to 30 units
            sheet.DefaultRowHeight = 30;

            // Specify the output file name
            string output = "SetDefaultRowHeight.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
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
