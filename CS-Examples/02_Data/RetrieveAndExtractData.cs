using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace RetrieveAndExtractData 
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a new workbook instance
            Workbook newBook = new Workbook();

            //Get the first worksheet.
            Worksheet newSheet = newBook.Worksheets[0];

            //Create a new workbook instance and load the sample Excel file.
			Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_3.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Retrieve data and extract it to the first worksheet of the new excel workbook.
            int i = 1;
            int columnCount = sheet.Columns.Length;
            foreach (CellRange range in sheet.Columns[0])
            {
                if (range.Text == "teacher")
                {
                    CellRange sourceRange = sheet.Range[range.Row, 1, range.Row, columnCount];
                    CellRange destRange = newSheet.Range[i, 1, i, columnCount];
                    sheet.Copy(sourceRange, destRange,true);
                    i++;
                }
            }

            // Specify the name for the resulting Excel file
            String result = "Result-RetrieveAndExtractDataToNewExcelFile.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
            newBook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            newBook.Dispose();
            workbook.Dispose();

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
