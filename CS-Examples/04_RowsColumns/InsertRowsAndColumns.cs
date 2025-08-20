using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace InsertRowsAndColumns
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }		
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load an existing file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\InsertRowsAndColumns.xls");

            // Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert a row into the worksheet at index 2
            worksheet.InsertRow(2);

            // Insert a column into the worksheet at index 2
            worksheet.InsertColumn(2);

            // Insert multiple rows into the worksheet starting at index 5, with a count of 2
            worksheet.InsertRow(5, 2);

            // Insert multiple columns into the worksheet starting at index 5, with a count of 2
            worksheet.InsertColumn(5, 2);

            // Specify the output file name
            string result = "InsertRowsAndColumns_out.xlsx";

            // Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // View the file
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
