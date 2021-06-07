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

namespace GetSettingsOfDataValidation
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

            //Get first worksheet of the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            //Cell B4 has the Decimal Validation
            CellRange cell = worksheet.Range["B4"];

            //Get the valditation of this cell
            Validation validation = cell.DataValidation;

            //Get the settings
            string allowType = validation.AllowType.ToString();
            string data = validation.CompareOperator.ToString();
            string minimum = validation.Formula1.ToString();
            string maximum = validation.Formula2.ToString();
            string ignoreBlank = validation.IgnoreBlank.ToString();
           
            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();

            //Set string format for displaying
            string result = string.Format("Settings of Validation: \r\nAllow Type: " + allowType + "\r\nData: " + data + "\r\nMinimum: " + minimum +"\r\nMaximum: " + maximum + "\r\nIgnoreBlank: "+ignoreBlank);

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
