using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace EditExcelComment
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_8.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Get the first comment.
            ExcelComment comment = sheet.Comments[0];

            //Edit the comment.
            comment.Text = "This comment has been edited by Spire.XLS.";

            String result = "Result-EditExcelComment.xlsx";

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
