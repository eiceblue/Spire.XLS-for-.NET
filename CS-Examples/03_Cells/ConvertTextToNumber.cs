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
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace ConvertTextToNumber
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.xlsx");

            //Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            //Convert text string format to number format
            worksheet.Range["D2:D8"].ConvertToNumber();

            //Specify the filename for the resulting Excel file
            String outputFile = "Output.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
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
