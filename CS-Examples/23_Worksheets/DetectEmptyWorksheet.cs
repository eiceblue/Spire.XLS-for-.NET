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
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadImages.xlsx");

            // Get the first worksheet from the workbook
            Worksheet worksheet1 = workbook.Worksheets[0];

            // Detect if the first worksheet is empty
            bool detect1 = worksheet1.IsEmpty;

            // Get the second worksheet from the workbook
            Worksheet worksheet2 = workbook.Worksheets[1];

            // Detect if the second worksheet is empty
            bool detect2 = worksheet2.IsEmpty;

            // Create a StringBuilder to save the content
            StringBuilder content = new StringBuilder();

            // Format the result string for displaying
            string result = string.Format("The first worksheet is empty or not: {0}\r\nThe second worksheet is empty or not: {1}", detect1, detect2);

            // Add the result string to the StringBuilder
            content.AppendLine(result);

            // Specify the output file path and name
            string outputFile = "Output.txt";

            // Save the result to a text file
            File.WriteAllText(outputFile, content.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();
  
            // Launch the file.
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
