using System;
using Spire.Xls;
using System.Windows.Forms;


namespace ReplacePartialText
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Set value for cell "A1"
            sheet.Range["A1"].Text = "Hello World";

            // Automatically adjusting the column width to fit the content.
            sheet.Range["A1"].AutoFitColumns();

            // Replace Partial Text
            sheet.CellList[0].TextPartReplace("World", "Spire");

            // Saving the modified workbook to a file named "replaced.xlsx" in the Excel 2016 format.
            workbook.SaveToFile("replaced.xlsx", ExcelVersion.Version2016);

            // Dispose of the workbook object to free up resources
            workbook.Dispose();

            // Launch the MS Excel file.
            ExcelDocViewer("replaced.xlsx");
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
