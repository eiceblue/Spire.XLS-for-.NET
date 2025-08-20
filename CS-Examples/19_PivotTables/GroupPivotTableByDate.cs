using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using Spire.Xls.Core;

namespace GroupPivotTableByDate
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new Workbook object
            Workbook workbook = new Workbook();
          
            // Load the workbook from the specified file path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\GroupPivotTableByDate.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first pivot table in the worksheet
            XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;
         
            // Get the first row field in the pivot table
            IPivotField field = pt.RowFields[0];

            // Set the start and end dates for grouping
            DateTime start = new DateTime(2023, 1, 5);
            DateTime end = new DateTime(2023, 3, 2);

            // Set the group by type to days
            PivotGroupByTypes[] types = new PivotGroupByTypes[] { PivotGroupByTypes.Days };
           
            // Create a new group with the specified start and end dates, group by type, and interval
            field.CreateGroup(start, end, types, 10);

            // Calculate the pivot table data
            pt.CalculateData();
          
            // Refresh the pivot table cache
            pt.Cache.IsRefreshOnLoad = true;

            // Set the output file name
            String result = "GroupPivotTableByDate_output.xlsx";

            // Save the workbook to the specified file path with the specified Excel version
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);
		}

		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
