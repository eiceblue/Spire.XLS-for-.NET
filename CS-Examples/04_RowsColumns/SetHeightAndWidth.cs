using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SetHeightAndWidth
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

            // Load an existing workbook from file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetHeightAndWidth.xls");

            // Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the width of column 4 to 30 units
            worksheet.SetColumnWidth(4, 30);

            // Set the height of row 4 to 30 units
            worksheet.SetRowHeight(4, 30);

            // Specify the output file name
            string result = "SetHeightAndWidth_out.xlsx";

            // Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
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
