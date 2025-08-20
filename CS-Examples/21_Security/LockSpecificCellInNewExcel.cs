using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace LockSpecificCellInNewExcel
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a workbook.
			Workbook workbook = new Workbook();

            // Create an empty worksheet.
            workbook.CreateEmptySheet();

            // Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            // Loop through all the rows in the worksheet and unlock them.
            for (int i = 0; i < 255; i++)
            {
                sheet.Rows[i].Style.Locked = false;
            }

            // Lock specific cell in the worksheet.
            sheet.Range["A1"].Text = "Locked";
            sheet.Range["A1"].Style.Locked = true;

            // Lock specific cell range in the worksheet.
            sheet.Range["C1:E3"].Text = "Locked";
            sheet.Range["C1:E3"].Style.Locked = true;

            // Set the password.
            sheet.Protect("123", SheetProtectionType.All);

            // Specify the output filename for the workbook
            String result = "Result-LockSpecificCellInNewlyXlsFile.xlsx";

            // Save the modified workbook to a file (in Excel 2013 format)
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

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
