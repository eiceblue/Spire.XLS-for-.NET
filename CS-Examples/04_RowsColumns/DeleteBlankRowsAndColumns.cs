using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace DeleteBlankRowsAndColumns
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_2.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Delete blank rows from the worksheet
            for (int i = sheet.Rows.Length - 1; i >= 0; i--)
            {
                if (sheet.Rows[i].IsBlank)
                {
                    sheet.DeleteRow(i + 1);
                }
            }

            // Delete blank columns from the worksheet
            for (int j = sheet.Columns.Length - 1; j >= 0; j--)
            {
                if (sheet.Columns[j].IsBlank)
                {
                    sheet.DeleteColumn(j + 1);
                }
            }

            // Specify the output file name
            string result = "Result-DeleteBlankRowsAndColumns.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the MS Excel file.
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
