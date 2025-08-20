using System;
using System.Windows.Forms;
using Spire.Xls;

namespace ShowOrHideGridLine
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            // Create a workbook
			Workbook workbook = new Workbook();

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorksheetSample2.xlsx");

            // Get the first and second worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            Worksheet sheet2 = workbook.Worksheets[1];

            // Hide grid line in the first worksheet
            sheet1.GridLinesVisible = false;

            //Show grid line in the first worksheet
            sheet2.GridLinesVisible = true;

            // Save the document
            string output = "ShowOrHideGridLine.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the file
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
