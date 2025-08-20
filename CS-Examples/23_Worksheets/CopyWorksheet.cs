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
using System.Text;
using System.Collections.Generic;

namespace CopyWorksheet
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a source workbook object
            Workbook sourceWorkbook = new Workbook();

            // Load the source Excel document from disk
            sourceWorkbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadImages.xlsx");

            // Get the first worksheet from the source workbook
            Worksheet srcWorksheet = sourceWorkbook.Worksheets[0];

            // Create a target workbook object
            Workbook targetWorkbook = new Workbook();

            // Load the target Excel document from disk
            targetWorkbook.LoadFromFile(@"..\..\..\..\..\..\Data\sample.xlsx");

            // Add a new worksheet to the target workbook
            Worksheet targetWorksheet = targetWorkbook.Worksheets.Add("added");

            // Copy the first worksheet of the source workbook to the new added worksheet in the target workbook
            targetWorksheet.CopyFrom(srcWorksheet);

            // Specify the output file path and name
            String outputFile = "Output.xlsx";

            // Save the modified target workbook to a file
            targetWorkbook.SaveToFile(outputFile, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            sourceWorkbook.Dispose();
            targetWorkbook.Dispose();

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
