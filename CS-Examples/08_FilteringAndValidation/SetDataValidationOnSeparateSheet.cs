using System;
using System.Windows.Forms;
using Spire.Xls;

namespace SetDataValidationOnSeparateSheet
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{

            // Create a workbook to store Excel data
            Workbook workbook = new Workbook();

            // Load the Excel document from disk into the workbook
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetDataValidationOnSeparateSheet.xlsx");

            // Access the first sheet in the workbook
            Worksheet sheet1 = workbook.Worksheets[0];

            // Set text in cell B10 on the first sheet
            sheet1.Range["B10"].Text = "Here is a dataValidation example.";

            // Access the second sheet in the workbook
            Worksheet sheet2 = workbook.Worksheets[1];

            // Enable the option to allow data from a different sheet in data validation
            sheet2.ParentWorkbook.Allow3DRangesInDataValidation = true;

            // Set the data range for data validation on cell B11 of the first sheet,
            // using the range A1:A7 from the second sheet as the source of data
            sheet1.Range["B11"].DataValidation.DataRange = sheet2.Range["A1:A7"];

            // Save the modified workbook with data validation to a new file named "result.xlsx"
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("result.xlsx");
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
