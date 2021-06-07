using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace CreateNestedGroup
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

            //Set the style.
            CellStyle style = workbook.Styles.Add("style");
            style.Font.Color = Color.CadetBlue;
            style.Font.IsBold = true;

            //Set the summary rows appear above detail rows.
            sheet.PageSetup.IsSummaryRowBelow = false;

            //Insert sample data to cells.
            sheet.Range["A1"].Value = "Project plan for project X";
            sheet.Range["A1"].CellStyleName = style.Name;

            sheet.Range["A3"].Value = "Set up";
            sheet.Range["A3"].CellStyleName = style.Name;
            sheet.Range["A4"].Value = "Task 1";
            sheet.Range["A5"].Value = "Task 2";
            sheet.Range["A4:A5"].BorderAround(LineStyleType.Thin);
            sheet.Range["A4:A5"].BorderInside(LineStyleType.Thin);

            sheet.Range["A7"].Value = "Launch";
            sheet.Range["A7"].CellStyleName = style.Name;
            sheet.Range["A8"].Value = "Task 1";
            sheet.Range["A9"].Value = "Task 2";
            sheet.Range["A8:A9"].BorderAround(LineStyleType.Thin);
            sheet.Range["A8:A9"].BorderInside(LineStyleType.Thin);

            //Group the rows that you want to group.
            sheet.GroupByRows(2, 9, false);
            sheet.GroupByRows(4, 5, false);
            sheet.GroupByRows(8, 9, false);

            String result = "Result-CreateNestedGroup.xlsx";

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
