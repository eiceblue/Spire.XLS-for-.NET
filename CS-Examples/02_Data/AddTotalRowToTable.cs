using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core;

namespace AddTotalRowToTable
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AddATotalRowToTable.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Create a table with the data from the specific cell range.
            IListObject table = sheet.ListObjects.Create("Table", sheet.Range["A1:D4"]);

            //Display total row.
            table.DisplayTotalRow = true;

            //Add a total row.
            table.Columns[0].TotalsRowLabel = "Total";
            table.Columns[1].TotalsCalculation = ExcelTotalsCalculation.Sum;
            table.Columns[2].TotalsCalculation = ExcelTotalsCalculation.Sum;
            table.Columns[3].TotalsCalculation = ExcelTotalsCalculation.Sum;

            String result = "Result-AddATotalRowToTable.xlsx";

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
