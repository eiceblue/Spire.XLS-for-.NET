using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace ExpandOrCollapseRows
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

            // Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_7.xlsx");

            // Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            // Get the first pivot table from the sheet
            Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTable = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;

            // Calculate data.
            pivotTable.CalculateData();

            // Collapse the rows.
            (pivotTable.PivotFields["Vendor No"] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3501", true);

            // Expand the rows.
            (pivotTable.PivotFields["Vendor No"] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3502", false);

            // Specify the filename for the resulting workbook
            String result = "Result-ExpandOrCollapseRowsInPivotTable.xlsx";

            // Save the modified workbook to a file using Excel 2013 format
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
