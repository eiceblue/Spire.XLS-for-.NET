using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace RefreshPivotTable
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
			Worksheet sheet = workbook.Worksheets[1];

            // Update the data source of PivotTable.
            sheet.Range["D2"].Value = "999";

            // Get the PivotTable that was built on the data source.
            XlsPivotTable pt = workbook.Worksheets[0].PivotTables[0] as XlsPivotTable;

            // Refresh the data of PivotTable.
            pt.Cache.IsRefreshOnLoad = true;

            // Specify the filename for the resulting file
            String result = "Result-RefreshPivotTable.xlsx";

            // Save the modified workbook to a file
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
