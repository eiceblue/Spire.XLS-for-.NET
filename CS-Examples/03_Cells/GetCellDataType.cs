using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace GetCellDataType
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_2.xlsx");

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            //Get the cell types of the cells in range ¡°C13:F13¡±
            foreach (CellRange range in sheet.Range["H2:H7"])
            {
                XlsWorksheet.TRangeValueType cellType = sheet.GetCellType(range.Row, range.Column, false);
                sheet[range.Row, range.Column + 1].Text = cellType.ToString();
                sheet[range.Row, range.Column + 1].Style.Font.Color = Color.Red;
                sheet[range.Row, range.Column + 1].Style.Font.IsBold = true;
            }

            // Specify the filename for the resulting Excel file
            String result = "Result-GetCellDataType.xlsx";
            // Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the MS Excel file.
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
