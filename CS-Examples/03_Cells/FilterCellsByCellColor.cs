using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core.Spreadsheet.AutoFilter;

namespace FilterCellsByCellColor
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_3.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Create an auto filter in the sheet and specify the range to be filterd
            sheet.AutoFilters.Range = sheet.Range["G1:G19"];

            //Get the column to be filterd
            FilterColumn filtercolumn = (FilterColumn)sheet.AutoFilters[0];

            //Add a color filter to filter the column based on cell color
            sheet.AutoFilters.AddFillColorFilter(filtercolumn, Color.Red);

            //Filter the data.
            sheet.AutoFilters.Filter();

            String result = "Result-FilterCellsByCellColor.xlsx";

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
