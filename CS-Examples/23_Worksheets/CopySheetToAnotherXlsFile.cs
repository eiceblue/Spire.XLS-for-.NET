using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace CopySheetToAnotherXlsFile
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

            //Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            //Put some data into header rows (A1:A4)
            for (int i = 1; i < 6; i++)
            {
                sheet.Range["A" + i].Text = string.Format("Header Row {0}", i);
                //sheet.Cells[i].Value = string.Format("Header Row {0}",i);
            }

            //Put some detail data (A5:A99)
            for (int i = 5; i < 100; i++)
            {
                sheet.Range["A" + i].Text = string.Format("Detail Row {0}", i);
                //sheet.Cells[i].Value = string.Format("Detail Row {0}",i);
            }

            //Define a pagesetup object based on the first worksheet.
            PageSetup pageSetup = sheet.PageSetup;

            //The first five rows are repeated in each page. It can be seen in print preview.
            pageSetup.PrintTitleRows = "$1:$5";

            //Create another Workbook.
            Workbook workbook1 = new Workbook();

            //Get the first worksheet in the book.
            Worksheet sheet1 = workbook1.Worksheets[0];

            //Copy worksheet to destination worsheet in another Excel file.
            sheet1.CopyFrom(sheet);

            String result = "Result-sourceFile.xlsx";
            String result1 = "Result-CopySheetToAnotherXlsFile.xlsx";

            //Save the source file we created.
            workbook.SaveToFile(result,ExcelVersion.Version2013);

            //Save the destination file.
            workbook1.SaveToFile(result1, ExcelVersion.Version2013);

            //Launch the MS Excel files.
            ExcelDocViewer(result);
            ExcelDocViewer(result1);
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
