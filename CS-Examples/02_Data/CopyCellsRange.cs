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
            //Create a workbook
			Workbook workbook = new Workbook();

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateTable.xlsx");
      
            //Get the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];

            //Specify a destination range 
            CellRange cells = sheet1.Range["G1:H19"];

            //Copy the selected range to destination range 
            sheet1.Range["B1:C19"].Copy(cells);

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
