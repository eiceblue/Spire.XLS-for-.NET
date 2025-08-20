using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core;

namespace AddTotalRowToTable
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

            // Load the Excel file from the specified path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AddATotalRowToTable.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Create a table with the data from the specified cell range
            IListObject table = sheet.ListObjects.Create("Table", sheet.Range["A1:D4"]);

            // Display the total row in the table
            table.DisplayTotalRow = true;

            // Add a total row to the table
            table.Columns[0].TotalsRowLabel = "Total";
            // Calculate the sum for column 1 in the total row
            table.Columns[1].TotalsCalculation = ExcelTotalsCalculation.Sum;
            // Calculate the sum for column 2 in the total row
            table.Columns[2].TotalsCalculation = ExcelTotalsCalculation.Sum;
            // Calculate the sum for column 3 in the total row
            table.Columns[3].TotalsCalculation = ExcelTotalsCalculation.Sum;

            // Specify the filename for the resulting Excel file
            String result = "Result-AddATotalRowToTable.xlsx"; // Specify the name for the resulting Excel file

            // Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object
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
