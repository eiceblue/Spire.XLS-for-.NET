using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace AddWorksheet
{
	public partial class Form1 :Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load an existing Excel document from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AddWorksheet.xlsx");

            // Add a new worksheet named "AddedSheet"
            Worksheet sheet = workbook.Worksheets.Add("AddedSheet");
            sheet.Range["C5"].Text = "This is a new sheet.";

            // Save the modified workbook to a file
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("Output.xlsx");
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
