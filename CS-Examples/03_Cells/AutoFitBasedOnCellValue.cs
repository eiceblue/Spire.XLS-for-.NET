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

namespace AutoFitBasedOnCellValue
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadImages.xlsx");

            //Get first worksheet of the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            //Set value for B8
            CellRange cell = worksheet.Range["B8"];
            cell.Text = "Welcome to Spire.XLS!";

            //Set the cell style
            CellStyle style = cell.Style;
            style.Font.Size = 16;
            style.Font.IsBold = true;

            //Autofit column width and row height based on cell value
            cell.AutoFitColumns();
            cell.AutoFitRows();

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
