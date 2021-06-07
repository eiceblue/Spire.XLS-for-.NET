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

namespace FindAndReplaceData
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
            Worksheet worksheet = workbook.Worksheets[0];

            //Find the "Brazil" string
            CellRange[] ranges = worksheet.FindAllString("Area", false, false);

            //Traverse the found ranges
            foreach (CellRange range in ranges)
            {
                //Replace it with "China"
                range.Text = "Area Code";
                //Highlight the color
                range.Style.Color = Color.Yellow;
            }
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
