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

namespace CopyCellsRange
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a workbook
			Workbook workbook = new Workbook();

            // Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateTable.xlsx");
      
            // Get the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];

            // Specify a destination range 
            CellRange cells = sheet1.Range["G1:H19"];

            // Copy the selected range to destination range 
            sheet1.Range["B1:C19"].Copy(cells);

            // Specify the name for the resulting Excel file 
            String outputFile = "Output.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013);

            // Dispose of the workbook object
            workbook.Dispose();

            // Launchthe output file.
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
