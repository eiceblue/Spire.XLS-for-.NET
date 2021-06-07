using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace WrapOrUnwrapTextInCells
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a workbook.
			Workbook workbook = new Workbook();

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Wrap the excel text;
            sheet.Range["C1"].Text = "e-iceblue is in facebook and welcome to like us";
            sheet.Range["C1"].Style.WrapText = true;
            sheet.Range["D1"].Text = "e-iceblue is in twitter and welcome to follow us";
            sheet.Range["D1"].Style.WrapText = true;

            //Unwrap the excel text;
            sheet.Range["C2"].Text = "http://www.facebook.com/pages/e-iceblue/139657096082266";
            sheet.Range["C2"].Style.WrapText = false;
            sheet.Range["D2"].Text = "https://twitter.com/eiceblue";
            sheet.Range["D2"].Style.WrapText = false;

            //Set the text color of Range["C1:D1"]
            sheet.Range["C1:D1"].Style.Font.Size = 15;
            sheet.Range["C1:D1"].Style.Font.Color = Color.Blue;
            //Set the text color of Range["C2:D2"]
            sheet.Range["C2:D2"].Style.Font.Size = 15;
            sheet.Range["C2:D2"].Style.Font.Color = Color.DeepSkyBlue;

            String result = "Result-WrapOrUnwrapTextInExcelCells.xlsx";

            //Save to file.
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            //Launch the MS Excel file.
            ExcelDocViewer(result);
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
