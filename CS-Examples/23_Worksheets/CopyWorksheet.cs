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

namespace CopyWorksheet
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
            Workbook sourceWorkbook = new Workbook();

            //Load the source Excel document from disk
            sourceWorkbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadImages.xlsx");

            //Get the first worksheet
            Worksheet srcWorksheet = sourceWorkbook.Worksheets[0];

            //Create a workbook
            Workbook targetWorkbook = new Workbook();

            //Load the target Excel document from disk
            targetWorkbook.LoadFromFile(@"..\..\..\..\..\..\Data\sample.xlsx");

            //Add a new worksheet
            Worksheet targetWorksheet = targetWorkbook.Worksheets.Add("added");

            //Copy the first worksheet of source Excel document to the new added worksheet of target Excel document
            targetWorksheet.CopyFrom(srcWorksheet);

            //String for output file 
            String outputFile = "Output.xlsx";

            //Save the file
            targetWorkbook.SaveToFile(outputFile, ExcelVersion.Version2013);

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
