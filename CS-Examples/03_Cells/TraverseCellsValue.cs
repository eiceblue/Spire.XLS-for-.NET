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

namespace TraverseCellsValue
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CellValues.xlsx");

            //Get first worksheet of the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            //Get the cell range collection 
            CellRange[] cellRangeCollection = worksheet.Cells;

            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();
            content.AppendLine("Values of the first sheet:");

            //Traverse cells value
            foreach (CellRange cellRange in cellRangeCollection)
            {
                //Set string format for displaying
                string result = string.Format("Cell: " + cellRange.RangeAddress + "   Value: " + cellRange.Value);

                //Add result string to StringBuilder
                content.AppendLine(result);
            }
            //String for output file 
            String outputFile = "Output.txt";

            //Save them to a txt file
            File.WriteAllText(outputFile, content.ToString());

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
