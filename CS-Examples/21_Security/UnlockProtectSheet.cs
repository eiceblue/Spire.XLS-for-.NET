using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace UnlockProtectSheet
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

            // Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\UnprotectProtectSheet.xlsx");

            // Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Unprotect the worksheet with password.
            sheet.Unprotect("e-iceblue");

            // Specify the output filename for the workbook
            String result = "Result-UnprotectProtectSheet.xlsx";

            // Save the modified workbook to a file
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
