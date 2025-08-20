using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace DifferentHeaderFooterOnFirstPage
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

            // Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            // Set value for the range
            sheet.Range["A1"].Text="Hello World";
            sheet.Range["F30"].Text = "Hello World";
            sheet.Range["G150"].Text = "Hello World";
   
            // Set the value to show the headers/footers for first page are different from the other pages.
            sheet.PageSetup.DifferentFirst = 1;
             
            // Set the header and footer for the first page.
            sheet.PageSetup.FirstHeaderString = "Different First page";
            sheet.PageSetup.FirstFooterString = "Different First footer";

            // Set the other pages' header and footer. 
            sheet.PageSetup.LeftHeader = "Demo of Spire.XLS";
            sheet.PageSetup.CenterFooter = "Footer by Spire.XLS";

            // Specify the file name for the resulting Excel file
            String result = "Result-AddDifferentHeaderFooterForTheFirstPage.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
	}
}
