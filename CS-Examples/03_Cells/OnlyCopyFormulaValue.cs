using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace OnlyCopyFormulaValue
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CopyOnlyFormulaValue1.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set the copy option to only copy formula values
            CopyRangeOptions copyOptions = CopyRangeOptions.OnlyCopyFormulaValue;

            // Define the source range to be copied
            CellRange sourceRange = sheet.Range["A6:E6"];

            // Copy the source range to a destination range using the specified copy options
            sheet.Copy(sourceRange, sheet.Range["A8:E8"], copyOptions);

            // Copy the source range to another destination range using the same copy options
            sourceRange.Copy(sheet.Range["A10:E10"], copyOptions);

            // Specify the output file name for the result
            string result = "Result-OnlyCopyFormulaValue.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the MS Excel file.
            ExcelDocViewer(result);
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
