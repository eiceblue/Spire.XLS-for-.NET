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
            //Create a workbook
			Workbook workbook = new Workbook();

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateTable.xlsx");
      
            //Get the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];

            //Specify a destination range to copy one column 
            CellRange columnCells = sheet1.Range["G1:G19"];

            //Copy the second column to destination range 
            sheet1.Columns[1].Copy(columnCells);

            //Specify a destination range to copy one row 
            CellRange rowCells = sheet1.Range["A21:E21"];

            //Copy the first row to destination range 
            sheet1.Rows[0].Copy(rowCells);

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
