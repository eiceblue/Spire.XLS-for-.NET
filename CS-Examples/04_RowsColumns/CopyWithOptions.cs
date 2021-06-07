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
            //Create a workbook
			Workbook workbook = new Workbook();

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.xlsx");
      
            //Get the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];

            //Add a new worksheet as destination sheet
            Worksheet destinationSheet = workbook.Worksheets.Add("DestSheet");

            //Specify a copy range of original sheet
            CellRange cellRange = sheet1.Range["B2:D4"];

            //Copy the specified range to added worksheet and keep original styles and update reference
            workbook.Worksheets[0].Copy(cellRange, workbook.Worksheets[1], 2, 1, true, true);

            //String for output file 
            String outputFile = "Output.xlsx";

            //Save the file
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013);

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
