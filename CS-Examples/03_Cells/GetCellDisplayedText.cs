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

namespace GetCellDisplayedText 
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
            cell.NumberValue = 0.012345;

            //Set the cell style
            CellStyle style = cell.Style;
            style.NumberFormat = "0.00";

            //Get the cell value
            string cellValue = cell.Value;

            //Get the displayed text of the cell
            string displayedText = cell.DisplayedText;

            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();

            //Set string format for displaying
            string result = string.Format("B8 Value: " + cellValue + "\r\nB8 displayed text: " + displayedText);

            //Add result string to StringBuilder
            content.AppendLine(result);

            //Specify the filename for the resulting file
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
