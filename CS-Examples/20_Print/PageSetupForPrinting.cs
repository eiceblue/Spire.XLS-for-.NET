using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using Spire.Xls;
using Spire.Xls.Charts;
using System.Text;
using System.Collections.Generic;

namespace PageSetupForPrinting
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateTable.xlsx");

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Specifying the print area
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.PrintArea = "A1:E19";

            // Define column A & E as title columns
            pageSetup.PrintTitleColumns = "$A:$E";

            // Define row numbers 1 as title rows
            pageSetup.PrintTitleRows = "$1:$2";

            // Allow to print with gridlines
            pageSetup.IsPrintGridlines = true;

            // Allow to print with row/column headings
            pageSetup.IsPrintHeadings = true;

            // Allow to print worksheet in black & white mode
            pageSetup.BlackAndWhite = true;

            // Allow to print comments as displayed on worksheet
            pageSetup.PrintComments = PrintCommentType.InPlace;

            // Set printing quality
            pageSetup.PrintQuality = 150;

            // Allow to print cell errors as N/A
            pageSetup.PrintErrors = PrintErrorsType.NA;

            // Set the printing order 
            pageSetup.Order = OrderType.OverThenDown;

            // Print file
            workbook.PrintDocument.Print();

            // Dispose of the workbook object to release resources
            workbook.Dispose();
		}
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
