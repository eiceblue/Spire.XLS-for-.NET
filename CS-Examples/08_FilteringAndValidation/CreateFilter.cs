using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace CreateFilter
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateFilter.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Specify the range for creating the filter (in this case, A1 to J1)
            sheet.AutoFilters.Range = sheet.Range["A1:J1"];

            // Save the modified workbook with the applied filter to a new file named "CreateFilter_out.xlsx"
            string result = "CreateFilter_out.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
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
