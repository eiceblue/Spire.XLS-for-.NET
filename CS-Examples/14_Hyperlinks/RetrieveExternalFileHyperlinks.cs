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

namespace RetrieveExternalFileHyperlinks
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\RetrieveExternalFileHyperlinks.xlsx");

            // Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            StringBuilder content = new StringBuilder();

            //Retrieve external file hyperlinks.
            foreach (HyperLink item in sheet.HyperLinks)
            {
                String address = item.Address;
                String sheetName = item.Range.WorksheetName;
                CellRange range = item.Range;
                content.AppendLine(String.Format("Cell[{0},{1}] in sheet \"" + sheetName + "\" contains File URL: {2}", range.Row, range.Column, address));
            }

            // Specify the output file name for the modified workbook
            String result = "Result-RetrieveExternalFileHyperlinks.txt";

            //Save to file.
            File.WriteAllText(result, content.ToString());

            // Dispose of the workbook object to release resources
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
