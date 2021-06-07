using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace LockSpecificRowInNewExcel
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

            //Create an empty worksheet.
            workbook.CreateEmptySheet();

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Loop through all the rows in the worksheet and unlock them.
            for (int i = 0; i < 255; i++)
            {
                sheet.Rows[i].Style.Locked = false;
            }

            //Lock the third row in the worksheet.
            sheet.Rows[2].Text = "Locked";
            sheet.Rows[2].Style.Locked = true;

            //Set the password.
            sheet.Protect("123", SheetProtectionType.All);

            String result = "Result-LockSpecificRowInNewlyXlsFile.xlsx";

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
