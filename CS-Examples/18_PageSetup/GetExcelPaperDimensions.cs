using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using System.Text;
using System.IO;

namespace GetExcelPaperDimensions
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

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Create a StringBuilder to store the result
            StringBuilder content = new StringBuilder();

            // Get the dimensions of A2 paper
            sheet.PageSetup.PaperSize = PaperSizeType.A2Paper;
            content.AppendLine("A2Paper: " + sheet.PageSetup.PageWidth + " x " + sheet.PageSetup.PageHeight);

            // Get the dimensions of A3 paper
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            content.AppendLine("PaperA3: " + sheet.PageSetup.PageWidth + " x " + sheet.PageSetup.PageHeight);

            // Get the dimensions of A4 paper
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            content.AppendLine("PaperA4: " + sheet.PageSetup.PageWidth + " x " + sheet.PageSetup.PageHeight);

            // Get the dimensions of letter-sized paper
            sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            content.AppendLine("PaperLetter: " + sheet.PageSetup.PageWidth + " x " + sheet.PageSetup.PageHeight);

            // Specify the output file name for the result
            string result = "Result-GetExcelPaperDimensions.txt";

            // Write the content of the StringBuilder to the result file
            File.WriteAllText(result, content.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the file.
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
