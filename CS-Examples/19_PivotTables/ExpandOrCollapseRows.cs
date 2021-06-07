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
            //Create a workbook.
			Workbook workbook = new Workbook();

            //Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_7.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Get the data in Pivot Table.
            Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTable = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;

            //Calculate Data.
            pivotTable.CalculateData();

            //Collapse the rows.
            (pivotTable.PivotFields["Vendor No"] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3501", true);

            //Expand the rows.
            (pivotTable.PivotFields["Vendor No"] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3502", false);

            String result = "Result-ExpandOrCollapseRowsInPivotTable.xlsx";

            //Save to file.
            workbook.SaveToFile(result, ExcelVersion.Version2013);

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
