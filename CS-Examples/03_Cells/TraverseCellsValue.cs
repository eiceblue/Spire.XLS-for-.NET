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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CellValues.xlsx");

            // Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of cell ranges in the worksheet
            CellRange[] cellRangeCollection = worksheet.Cells;

            // Create a StringBuilder to save the content
            StringBuilder content = new StringBuilder();
            content.AppendLine("Values of the first sheet:");

            // Traverse through the cells and retrieve their values
            foreach (CellRange cellRange in cellRangeCollection)
            {
                // Set the string format for displaying the cell address and value
                string result = string.Format("Cell: " + cellRange.RangeAddress + "   Value: " + cellRange.Value);

                // Add the result string to the StringBuilder
                content.AppendLine(result);
            }

            // Specify the output file name as a txt file
            string outputFile = "Output.txt";

            // Save the content to a txt file
            File.WriteAllText(outputFile, content.ToString());

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
