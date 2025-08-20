using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using System.Text;
using System.IO;

namespace CopyMultipeSheetsToSingleSheet
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_11.xlsx");

            // Get the first worksheet.
			Worksheet sheet1 = workbook.Worksheets[0];

            // Copy all objects(such as text, shape, image...) from sheet2 to sheet1
            for (int i = 1; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet2 = workbook.Worksheets[i];
                sheet2.Copy((CellRange)sheet2.MaxDisplayRange, sheet1, sheet1.LastRow + 1, sheet2.FirstColumn, true);
            }

            // Save to file
            string fileName = "CopyMultipeSheetsToSingleSheet_result.xlsx";
            workbook.SaveToFile(fileName,ExcelVersion.Version2016);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(fileName);
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
