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

namespace CopySingleColumnAndRow
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateTable.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet1 = workbook.Worksheets[0];

            // Specify the destination range to copy one column (column G)
            CellRange columnCells = sheet1.Range["G1:G19"];

            // Copy the second column (column index 1) to the destination range
            sheet1.Columns[1].Copy(columnCells);

            // Specify the destination range to copy one row (row 21, columns A to E)
            CellRange rowCells = sheet1.Range["A21:E21"];

            // Copy the first row (row index 0) to the destination range
            sheet1.Rows[0].Copy(rowCells);

            // Specify the output file name
            string outputFile = "Output.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launching the output file.
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
