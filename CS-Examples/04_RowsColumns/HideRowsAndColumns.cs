using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace HideRowsAndColumns
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

            // Load an existing Excel file named "HideRowsAndColumns.xls"
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\HideRowsAndColumns.xls");

            // Get the first worksheet from the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide the second column of the worksheet
            worksheet.HideColumn(2);

            // Hide the fourth row of the worksheet
            worksheet.HideRow(4);

            // Save the modified workbook to a new file named "HideRowsAndColumns.xlsx" in Excel 2010 format
            workbook.SaveToFile("HideRowsAndColumns.xlsx", ExcelVersion.Version2010);

            // View the file
            ExcelDocViewer("HideRowsAndColumns.xlsx");
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
