using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace SplitDataIntoMultipleColumns
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SplitExcelDataIntoMultipleCols.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Split data into separate columns by the delimited characters ¨C space.
            string[] splitText = null;
            string text = null;
            for (int i = 1; i < sheet.LastRow; i++)
            {
                text = sheet.Range[i + 1, 1].Text;
                splitText = text.Split(' ');
                for (int j = 0; j < splitText.Length; j++)
                {
                    sheet.Range[i + 1, 1 + j + 1].Text = splitText[j];
                }
            }

            String result = "Result-SplitExcelDataIntoMultipleCols.xlsx";

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
