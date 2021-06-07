using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace SetArrayOfValuesIntoRange
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

            //Create an empty worksheet.
            workbook.CreateEmptySheets(1);

            //Get the worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            //Set the value of max row and column.
            int maxRow = 10000;
            //int maxRow = 5;
            int maxCol = 200;
            //int maxCol = 5;

            //Output an array of data to a range of worksheet.
            object[,] myarray = new object[maxRow + 1, maxCol + 1];
            bool[,] isred = new bool[maxRow + 1, maxCol + 1];
            for (int i = 0; i <= maxRow; i++)
                for (int j = 0; j <= maxCol; j++)
                {
                    myarray[i, j] = i + j;
                    if ((int)myarray[i, j] > 8)
                        isred[i, j] = true;
                }

            sheet.InsertArray(myarray, 1, 1);

            String result = "Result-SetArrayOfValuesIntoRange.xlsx";

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
