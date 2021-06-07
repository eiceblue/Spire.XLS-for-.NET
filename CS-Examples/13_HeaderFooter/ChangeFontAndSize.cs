using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace ChangeFontAndSize
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a workbook.
			Workbook workbook = new Workbook();

            //Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeFontAndSizeForHeaderAndFooter.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Set the new font and size for the header and footer
            string text = sheet.PageSetup.LeftHeader;
            //"Arial Unicode MS" is font name, "18" is font size
            text = "&\"Arial Unicode MS\"&18 Header Footer Sample by Spire.XLS ";
            sheet.PageSetup.LeftHeader = text;
            sheet.PageSetup.RightFooter = text;

            String result = "Result-ChangeFontAndSizeForHeaderAndFooter.xlsx";

            //Save to file.
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
	}
}
