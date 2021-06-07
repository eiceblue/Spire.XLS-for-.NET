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

namespace FillDataInWorksheet
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

            //Get first worksheet of the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            //Fill data
            worksheet.Range["A1"].Style.Font.IsBold = true;
            worksheet.Range["B1"].Style.Font.IsBold = true;
            worksheet.Range["C1"].Style.Font.IsBold = true;
            worksheet.Range["A1"].Text = "Month";
            worksheet.Range["A2"].Text = "January";
            worksheet.Range["A3"].Text = "February";
            worksheet.Range["A4"].Text = "March";
            worksheet.Range["A5"].Text = "April";
            worksheet.Range["B1"].Text = "Payments";
            worksheet.Range["B2"].NumberValue = 251;
            worksheet.Range["B3"].NumberValue = 515;
            worksheet.Range["B4"].NumberValue = 454;
            worksheet.Range["B5"].NumberValue = 874;
            worksheet.Range["C1"].Text = "Sample";
            worksheet.Range["C2"].Text = "Sample1";
            worksheet.Range["C3"].Text = "Sample2";
            worksheet.Range["C4"].Text = "Sample3";
            worksheet.Range["C5"].Text = "Sample4";

            //Set width for the second column
            worksheet.SetColumnWidth(2, 10);

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
