using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using Spire.Xls;
using Spire.Xls.Charts;

namespace CopyWithOptions
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

            // Load an existing Excel document from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet1 = workbook.Worksheets[0];

            // Add a new worksheet as the destination sheet
            Worksheet destinationSheet = workbook.Worksheets.Add("DestSheet");

            // Specify the range to be copied from the original sheet (B2:D4)
            CellRange cellRange = sheet1.Range["B2:D4"];

            // Copy the specified range to the added worksheet, keeping the original styles and updating references
            workbook.Worksheets[0].Copy(cellRange, workbook.Worksheets[1], 2, 1, true, true);

            // Specify the output file name
            string outputFile = "Output.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the output file.
            Viewer(outputFile);
		}
		private void Viewer( string fileName )
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
