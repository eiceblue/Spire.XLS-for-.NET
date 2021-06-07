using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SetDataValidationOnSeparateSheet
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetDataValidationOnSeparateSheet.xlsx");

            //This is the first sheet
            Worksheet sheet1 = workbook.Worksheets[0];

            sheet1.Range["B10"].Text = "Here is a dataValidation example.";

            //This is the second sheet
			Worksheet sheet2 = workbook.Worksheets[1];

            //The property is to enable the data can be from different sheet.
            sheet2.ParentWorkbook.Allow3DRangesInDataValidation = true;
            sheet1.Range["B11"].DataValidation.DataRange = sheet2.Range["A1:A7"];

            workbook.SaveToFile("result.xlsx",ExcelVersion.Version2013);
            ExcelDocViewer("result.xlsx");
		}
        private void ExcelDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
	}
}
