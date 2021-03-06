using System;
using System.Windows.Forms;
using Spire.Xls;

namespace AutoFitRowInRange
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            //Create a workbook
			Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AutoFitSample.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Autofit the second row of the worksheet
            sheet.AutoFitRow(2, 1, 2, false);

            //Save the document
            string output = "AutoFitRowInRange.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

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
