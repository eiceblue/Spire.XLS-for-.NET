using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace GroupRowsAndColumns
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

            // Load an existing Excel file into the workbook
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\GroupRowsAndColumns.xls");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Group rows 1 to 5 (excluding child groups)
            sheet.GroupByRows(1, 5, false);

            // Group columns 1 to 3 (excluding child groups)
            sheet.GroupByColumns(1, 3, false);

            // Save the modified workbook to a new file in Excel 2010 format
            workbook.SaveToFile("GroupRowsAndColumns.xlsx", ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("GroupRowsAndColumns.xlsx");
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
