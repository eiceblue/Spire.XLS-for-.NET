using System;
using System.Windows.Forms;
using Spire.Xls;

namespace HideFormulas
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FormulasSample.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide the formulas in the used range
            sheet.AllocatedRange.IsFormulaHidden = true;

            // Protect the worksheet with password
            sheet.Protect("e-iceblue");

            // Specify the output filename for the workbook
            string output = "HideFormulas.xlsx";

            // Save the modified workbook to a file (in Excel 2013 format)
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the Excel file
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
