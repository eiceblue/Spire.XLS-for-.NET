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

namespace DetectEmptyWorksheet
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

            //Get the first worksheet
            Worksheet worksheet1 = workbook.Worksheets[0];

            //Detect the first worksheet is empty or not
            bool detect1 = worksheet1.IsEmpty;

            //Get the second worksheet
            Worksheet worksheet2 = workbook.Worksheets[1];

            //Detect the second worksheet is empty or not
            bool detect2 = worksheet2.IsEmpty;

            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();

            //Set string format for displaying
            string result = string.Format("The first worksheet is empty or not: " + detect1 + "\r\nThe second worksheet is empty or not: " + detect2);

            //Add result string to StringBuilder
            content.AppendLine(result);

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
