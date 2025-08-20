using System;
using System.Windows.Forms;

using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace ShowDataFieldInRow
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a new excel document
            Workbook workbook = new Workbook();
            //Load an excel document with Pivot table from the file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PivotTableExample.xlsx");

            // Get the worksheet where the pivot table is located
            Worksheet sheet = workbook.Worksheets[1];

            // Access the pivot table in the worksheet
            XlsPivotTable pivotTable = sheet.PivotTables[0] as XlsPivotTable;

            // Show the data field in the row area of the pivot table
            pivotTable.ShowDataFieldInRow = true;

            // Calculate the data in the pivot table
            pivotTable.CalculateData();

            // Save the modified workbook to a file
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2016);

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
