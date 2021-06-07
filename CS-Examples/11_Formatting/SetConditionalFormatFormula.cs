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
using System.Collections.Generic;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using System.Drawing.Imaging;
using Spire.Xls.Core;
using Spire.Xls.Core.Spreadsheet;
using System.Text;
using Spire.Xls.Core.Spreadsheet.Collections;

namespace SetConditionalFormatFormula
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

            //Get the default first  worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Add ConditionalFormat
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();

            //Define the range
            xcfs.AddRange(sheet.Range["B5"]);

            //Add condition
            IConditionalFormat format = xcfs.AddCondition();
            format.FormatType = ConditionalFormatType.CellValue;

            //If greater than 1000
            format.FirstFormula = "1000";
            format.Operator = ComparisonOperatorType.Greater;
            format.BackColor = Color.Orange;

            sheet.Range["B1"].NumberValue=40;
            sheet.Range["B2"].NumberValue=500;
            sheet.Range["B3"].NumberValue=300;
            sheet.Range["B4"].NumberValue=400;
            
            //Set a SUM formula for B5
            sheet.Range["B5"].Formula = "=SUM(B1:B4)";

            //Add text
            sheet.Range["C5"].Text = "If Sum of B1:B4 is greater than 1000, B5 will have orange background.";
       
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
