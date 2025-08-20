using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace SetOtherPrintingOptions
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
            workbook.LoadFromFile(@"../../../../../../Data/Template_Xls_1.xlsx");

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Get the reference of the PageSetup of the worksheet.
            PageSetup pageSetup = sheet.PageSetup;

            // Allow to print gridlines.
            pageSetup.IsPrintGridlines = true;

            // Allow to print row/column headings.
            pageSetup.IsPrintHeadings = true;

            // Allow to print worksheet in black & white mode.
            pageSetup.BlackAndWhite = true;

            // Allow to print comments as displayed on worksheet.
            pageSetup.PrintComments = PrintCommentType.InPlace;

            // Allow to print worksheet with draft quality.
            pageSetup.Draft = true;

            // Allow to print cell errors as N/A.
            pageSetup.PrintErrors = PrintErrorsType.NA;

            // Specify the output file name for the result
            String result = "Result-SetOtherPrintOptionsOfXlsFile.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
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
