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

namespace GetIntersectionOfTwoRanges
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_1.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Get the two ranges.
            CellRange range = sheet.Range["A2:D7"].Intersect(sheet.Range["B2:E8"]);

            StringBuilder content = new StringBuilder();
            content.AppendLine("The intersection of the two ranges \"A2:D7\" and \"B2:E8\" is:");

            //Get the intersection of the two ranges.
            foreach (CellRange r in range)
            {
                content.AppendLine(r.Value.ToString());
            }

            String result = "Result-GetTheIntersectionOfTwoRanges.txt";

            //Save to file.
            File.WriteAllText(result,content.ToString());

            //Launch the file.
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
