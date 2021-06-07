using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace SetMarginsOfExcelSheet
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_4.xlsx");

            //Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            //Get the PageSetup object of the first worksheet.
            PageSetup pageSetup = sheet.PageSetup;

            //Set bottom,left,right and top page margins.
            pageSetup.BottomMargin = 2;
            pageSetup.LeftMargin = 1;
            pageSetup.RightMargin = 1;
            pageSetup.TopMargin = 3;

            String result = "Result-SetMarginsOfExcelSheet.xlsx";

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
