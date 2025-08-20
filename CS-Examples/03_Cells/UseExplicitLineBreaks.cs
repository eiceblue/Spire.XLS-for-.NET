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

namespace UseExplicitLineBreaks
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

            // Get the first default worksheet in the workbook
            Worksheet sheet1 = workbook.Worksheets[0];

            // Specify a cell range
            CellRange c5 = sheet1.Range["C5"];

            // Set the cell width for the specified range
            sheet1.SetColumnWidth(c5.Column, 70);

            // Put the string value with explicit line breaks into the cell
            c5.Value = "Spire.XLS for .NET is a professional Excel .NET API\n that can be used to create, read, \nwrite, convert and print Excel files in any type \nof .NET(C#, VB.NET, ASP.NET, .NET Core) application. \nSpire.XLS for .NET offers object model\n Excel API for speeding up Excel programming in .NET platform -\n create new Excel documents from template, edit existing \nExcel documents and \nconvert Excel files.";

            // Enable text wrap for the cell
            c5.IsWrapText = true;

            // Specify the output file name as an xlsx file
            string outputFile = "Output.xlsx";

            // Save the workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launching the output file.
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
